"""Read and normalise the ONS constrained-off datasets."""

from __future__ import annotations

import re

import numpy as np
import pandas as pd

from .fetch import list_prefix, read_csv
from .sources import ONS_DATASETS, ONS_S3, RAZAO_ORDER

MONTH_RE = re.compile(r"_(\d{4})_(\d{2})\.csv$", re.I)

# 30-minute settlement periods -> energy in MWh
HALF_HOUR = 0.5


def months_available(dataset: str) -> dict[str, dict]:
    """{'2026_09': {'name':..., 'url':..., 'etag':...}, ...} for one dataset."""
    prefix, stem = ONS_DATASETS[dataset]
    listing = list_prefix(prefix)
    out: dict[str, dict] = {}
    for name, meta in listing.items():
        if not name.lower().endswith(".csv"):
            continue
        m = MONTH_RE.search(name)
        if not m or not name.upper().startswith(stem.upper()):
            continue
        out[f"{m.group(1)}_{m.group(2)}"] = {
            "name": name,
            "url": f"{ONS_S3}/{prefix}/{name}",
            "etag": meta["etag"],
            "size": meta["size"],
        }
    return dict(sorted(out.items()))


# --------------------------------------------------------------------------
# constrained-off "TM" files (one row per plant/conjunto per half hour)
# --------------------------------------------------------------------------
TM_COLS = [
    "id_subsistema",
    "id_estado",
    "nom_estado",
    "nom_usina",
    "id_ons",
    "ceg",
    "din_instante",
    "val_geracao",
    "val_geracaolimitada",
    "val_disponibilidade",
    "val_geracaoreferencia",
    "val_geracaoreferenciafinal",
    "cod_razaorestricao",
    "cod_origemrestricao",
    "dsc_restricao",
    "nom_pontoconexao",
    "nom_agenteoperador",
    "val_geracaonaorealizadaapurada",
    "num_minutos_rel",
    "num_minutos_cnf",
    "num_minutos_ene",
    "num_minutos_restricao",
]


def load_tm(url: str, fonte: str) -> pd.DataFrame:
    """Load one monthly constrained-off file and add the derived energy columns."""
    df = read_csv(url)
    keep = [c for c in TM_COLS if c in df.columns]
    df = df[keep].copy()
    # Os arquivos antigos não têm todas as colunas: dsc_restricao só existe a partir
    # de set/2025 e os num_minutos_* aparecem depois do início da série eólica.
    # Criamos as que faltam para o resto do pipeline não precisar checar cada uma.
    for c in TM_COLS:
        if c not in df.columns:
            df[c] = pd.NA
    df["din_instante"] = pd.to_datetime(df["din_instante"], errors="coerce")
    df = df.dropna(subset=["din_instante", "id_ons"])

    for c in (
        "val_geracao",
        "val_geracaolimitada",
        "val_disponibilidade",
        "val_geracaoreferencia",
        "val_geracaoreferenciafinal",
        "val_geracaonaorealizadaapurada",
        "num_minutos_rel",
        "num_minutos_cnf",
        "num_minutos_ene",
        "num_minutos_restricao",
    ):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df["fonte"] = fonte  # 'Solar' | 'Eólica'
    for c in ("cod_razaorestricao", "cod_origemrestricao"):
        df[c] = df[c].astype("object").fillna("").astype(str).str.strip().str.upper()
        df.loc[df[c].isin(("NAN", "<NA>", "NONE")), c] = ""

    # O patamar só conta como cortado quando o ONS declarou razão de restrição.
    df["restrito"] = df["cod_razaorestricao"].ne("")

    # GNRa (val_geracaonaorealizadaapurada) só existe nos arquivos a partir de
    # jan/2026.  Para o histórico anterior reproduzimos a definição do próprio
    # ONS — "diferença entre a geração de referência e a geração verificada
    # (se menor que zero, GNRa = 0), nos períodos em que houve limitação" — e
    # marcamos a origem, para a planilha poder distinguir apurado de calculado.
    gnr_pub = df["val_geracaonaorealizadaapurada"]
    gnr_calc = (df["val_geracaoreferencia"] - df["val_geracao"]).clip(lower=0)
    df["gnr_origem"] = np.where(gnr_pub.notna(), "publicada", "calculada")
    gnr = gnr_pub.fillna(gnr_calc).fillna(0.0).clip(lower=0)

    df["mwh_gnr"] = gnr * HALF_HOUR
    df["mwh_gerado"] = df["val_geracao"].fillna(0.0) * HALF_HOUR
    df["mwh_referencia"] = df["val_geracaoreferencia"].fillna(0.0) * HALF_HOUR
    df.loc[~df["restrito"], "mwh_gnr"] = 0.0
    df.loc[~df["restrito"], "gnr_origem"] = ""

    # Os num_minutos_* também só existem a partir de jan/2026.  Antes disso o
    # patamar semi-horário inteiro conta como restrito (30 min) na razão
    # declarada — é a melhor aproximação possível com o que o ONS publicou.
    falta_min = df["num_minutos_restricao"].isna()
    df.loc[falta_min, "num_minutos_restricao"] = np.where(df.loc[falta_min, "restrito"], 30.0, 0.0)
    for code, col in (("REL", "num_minutos_rel"), ("CNF", "num_minutos_cnf"),
                      ("ENE", "num_minutos_ene")):
        falta = df[col].isna()
        df.loc[falta, col] = np.where(
            df.loc[falta, "restrito"] & df.loc[falta, "cod_razaorestricao"].eq(code), 30.0, 0.0)
    df["minutos_origem"] = np.where(falta_min, "estimado", "publicado")

    df["data"] = df["din_instante"].dt.date
    df["hora"] = df["din_instante"].dt.floor("h")
    return df


def aggregate_daily(df: pd.DataFrame) -> pd.DataFrame:
    """conjunto x day x reason -> energy, minutes and R$ (if priced)."""
    money = "rs_gnr" if "rs_gnr" in df.columns else None
    gcols = ["fonte", "id_subsistema", "id_estado", "id_ons", "nom_usina", "data",
             "cod_razaorestricao"]
    agg = {
        "mwh_gnr": "sum",
        "mwh_gerado": "sum",
        "mwh_referencia": "sum",
        "num_minutos_restricao": "sum",
        "restrito": "sum",
    }
    for c in ("num_minutos_rel", "num_minutos_cnf", "num_minutos_ene"):
        if c in df.columns:
            agg[c] = "sum"
    for c in ("gnr_origem", "minutos_origem"):
        if c in df.columns:
            agg[c] = "max"  # 'publicada' > 'calculada'; 'publicado' > 'estimado'
    if money:
        agg[money] = "sum"
    out = df.groupby(gcols, dropna=False, observed=True).agg(agg).reset_index()
    out = out.rename(columns={
        "restrito": "n_semihoras", "num_minutos_restricao": "minutos",
        "num_minutos_rel": "min_REL", "num_minutos_cnf": "min_CNF",
        "num_minutos_ene": "min_ENE",
    })
    for c in ("min_REL", "min_CNF", "min_ENE"):
        if c not in out.columns:
            out[c] = 0.0
    out["data"] = pd.to_datetime(out["data"])
    return out


# --------------------------------------------------------------------------
# capacity-factor file: the conjunto master (name, capacity, lat/lon)
# --------------------------------------------------------------------------
FC_COLS = {
    "id_subsistema": "id_subsistema",
    "nom_subsistema": "nom_subsistema",
    "id_estado": "id_estado",
    "nom_estado": "nom_estado",
    "nom_pontoconexao": "nom_pontoconexao",
    "val_latitudesecoletora": "lat",
    "val_longitudesecoletora": "lon",
    "nom_modalidadeoperacao": "modalidade",
    "nom_tipousina": "tipo",
    "nom_usina_conjunto": "nome",
    "id_ons": "id_ons",
    "val_capacidadeinstalada": "capacidade_mw",
}


def load_conjuntos(urls: list[str]) -> pd.DataFrame:
    """Build the conjunto master from one or more FATOR_CAPACIDADE-2 months."""
    frames: list[pd.DataFrame] = []
    vistos, estaveis = -1, 0
    for url in urls:
        df = read_csv(url, usecols=lambda c: c in FC_COLS)
        df = df.rename(columns=FC_COLS)
        for c in ("lat", "lon", "capacidade_mw"):
            df[c] = pd.to_numeric(df[c], errors="coerce")
        df = df.dropna(subset=["id_ons"])
        df = (
            df.sort_values("capacidade_mw")
            .groupby("id_ons", as_index=False)
            .last()  # newest capacity / coordinates win
        )
        frames.append(df)
        # Os arquivos de fator de capacidade são grandes (um mês inteiro de dados
        # horários por usina). Consolidamos a cada mês e paramos assim que dois
        # meses seguidos não trouxerem nenhum conjunto novo — o que interessa
        # aqui é só o cadastro (nome, coordenada, capacidade), não a série.
        parcial = pd.concat(frames, ignore_index=True)
        parcial = parcial.sort_values("capacidade_mw").groupby("id_ons", as_index=False).last()
        frames = [parcial]
        if len(parcial) == vistos:
            estaveis += 1
            if estaveis >= 2:
                break
        else:
            estaveis, vistos = 0, len(parcial)
    out = frames[0] if frames else pd.DataFrame()
    if out.empty:
        raise SystemExit("não consegui montar o cadastro de conjuntos (fator de capacidade)")
    out["fonte"] = out["tipo"].map({"Solar": "Solar", "Eólica": "Eólica"}).fillna(out["tipo"])
    return out


# --------------------------------------------------------------------------
# detalhamento por usina (dataset *_detail)
# --------------------------------------------------------------------------
# Schema bem diferente do TM: traz irradiância, geração estimada e verificada
# por usina, mas NÃO traz razão de restrição, geração limitada nem GNRa.  As
# colunas id_ons_conjuntousina e flg_geracaorestrita só existem a partir de
# jan/2026; antes disso o vínculo com o conjunto é o nom_conjuntousina e o
# patamar restrito vem da razão declarada no arquivo de conjunto.
DETAIL_COLS = [
    "id_subsistema", "id_estado", "nom_modalidadeoperacao", "nom_conjuntousina",
    "nom_usina", "id_ons", "ceg", "din_instante", "val_irradianciaverificado",
    "val_geracaoestimada", "val_geracaoverificada", "id_ons_conjuntousina",
    "nom_origemgeracaoreferencia", "val_geracaoreferenciafinal", "flg_geracaorestrita",
]


def load_detail(url: str, conjuntos: list[str] | None = None) -> pd.DataFrame:
    """Carrega um mês do detalhamento por usina, opcionalmente já filtrado."""
    df = read_csv(url)
    for c in DETAIL_COLS:
        if c not in df.columns:
            df[c] = pd.NA
    df = df[DETAIL_COLS].copy()
    df["din_instante"] = pd.to_datetime(df["din_instante"], errors="coerce")
    df = df.dropna(subset=["din_instante", "id_ons"])
    for c in ("val_irradianciaverificado", "val_geracaoestimada", "val_geracaoverificada",
              "val_geracaoreferenciafinal"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["flg_geracaorestrita"] = pd.to_numeric(df["flg_geracaorestrita"], errors="coerce")

    if conjuntos:
        alvo = set(conjuntos)
        vinc = df["id_ons_conjuntousina"].astype("object")
        df = df[vinc.isin(alvo) | df["nom_conjuntousina"].isin(alvo)]
    df["data"] = df["din_instante"].dt.date
    df["hora"] = df["din_instante"].dt.floor("h")
    return df


def merge_detail_razao(detail: pd.DataFrame, tm: pd.DataFrame) -> pd.DataFrame:
    """Traz a razão/origem da restrição do conjunto para cada linha de usina.

    O arquivo por usina não publica cod_razaorestricao. A restrição é declarada
    no nível do conjunto, então casamos por (conjunto, data/hora). O vínculo usa
    id_ons_conjuntousina quando existe e cai no nome do conjunto quando não.
    """
    chave_tm = tm[["id_ons", "nom_usina", "din_instante", "cod_razaorestricao",
                   "cod_origemrestricao", "dsc_restricao", "val_geracaolimitada",
                   "val_disponibilidade"]].rename(
        columns={"id_ons": "_cj_id", "nom_usina": "_cj_nome"})

    d = detail.copy()
    d["_cj_id"] = d["id_ons_conjuntousina"]
    out = d.merge(chave_tm.drop(columns=["_cj_nome"]), on=["_cj_id", "din_instante"], how="left")

    falta = out["cod_razaorestricao"].isna()
    if falta.any():
        por_nome = chave_tm.drop(columns=["_cj_id"]).rename(
            columns={"_cj_nome": "nom_conjuntousina"})
        rep = out.loc[falta, ["nom_conjuntousina", "din_instante"]].merge(
            por_nome, on=["nom_conjuntousina", "din_instante"], how="left")
        for c in ("cod_razaorestricao", "cod_origemrestricao", "dsc_restricao",
                  "val_geracaolimitada", "val_disponibilidade"):
            out.loc[falta, c] = rep[c].to_numpy()

    out["cod_razaorestricao"] = out["cod_razaorestricao"].fillna("")
    out["cod_origemrestricao"] = out["cod_origemrestricao"].fillna("")
    restrito = out["cod_razaorestricao"].ne("")
    if out["flg_geracaorestrita"].notna().any():
        restrito = restrito | out["flg_geracaorestrita"].fillna(0).eq(1)
    out["restrito"] = restrito

    # Curtailment por usina: geração estimada − verificada nos patamares restritos.
    gnr = (out["val_geracaoestimada"] - out["val_geracaoverificada"]).clip(lower=0)
    out["mwh_gnr"] = np.where(out["restrito"], gnr.fillna(0.0), 0.0) * HALF_HOUR
    out["mwh_gerado"] = out["val_geracaoverificada"].fillna(0.0) * HALF_HOUR
    out["mwh_referencia"] = out["val_geracaoestimada"].fillna(0.0) * HALF_HOUR
    out["minutos"] = np.where(out["restrito"], 30.0, 0.0)
    return out


def aggregate_daily_usina(df: pd.DataFrame) -> pd.DataFrame:
    """usina × dia × razão — a granularidade escolhida para o histórico por usina."""
    gcols = ["nom_conjuntousina", "id_ons_conjuntousina", "nom_usina", "id_ons",
             "id_subsistema", "id_estado", "data", "cod_razaorestricao"]
    agg = {"mwh_gnr": "sum", "mwh_gerado": "sum", "mwh_referencia": "sum",
           "minutos": "sum", "restrito": "sum"}
    if "rs_gnr" in df.columns:
        agg["rs_gnr"] = "sum"
    out = df.groupby(gcols, dropna=False, observed=True).agg(agg).reset_index()
    out = out.rename(columns={"restrito": "n_semihoras"})
    out["data"] = pd.to_datetime(out["data"])
    return out


def reason_matrix(daily: pd.DataFrame) -> pd.DataFrame:
    """Wide table: one column of MWh per restriction reason."""
    p = (
        daily.pivot_table(
            index=["fonte", "id_subsistema", "id_ons", "nom_usina", "data"],
            columns="cod_razaorestricao",
            values="mwh_gnr",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )
    for code in RAZAO_ORDER:
        if code not in p.columns:
            p[code] = 0.0
    p["total"] = p[RAZAO_ORDER].sum(axis=1)
    return p
