"""Read and normalise the ONS constrained-off datasets."""

from __future__ import annotations

import re

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
    df["cod_razaorestricao"] = df["cod_razaorestricao"].fillna("").str.strip().str.upper()
    df["cod_origemrestricao"] = df.get(
        "cod_origemrestricao", pd.Series(index=df.index, dtype=object)
    ).fillna("").str.strip().str.upper()

    # Energia não gerada, em MWh.  GNRa is a MWmed average over the 30-minute
    # settlement period, so MWh = MWmed * 0.5.
    df["mwh_gnr"] = df["val_geracaonaorealizadaapurada"].fillna(0.0).clip(lower=0) * HALF_HOUR
    df["mwh_gerado"] = df["val_geracao"].fillna(0.0) * HALF_HOUR
    df["mwh_referencia"] = df["val_geracaoreferencia"].fillna(0.0) * HALF_HOUR

    # The half hour only counts as curtailed when ONS flagged a reason.
    df["restrito"] = df["cod_razaorestricao"].ne("")
    df.loc[~df["restrito"], "mwh_gnr"] = 0.0

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
    if money:
        agg[money] = "sum"
    out = df.groupby(gcols, dropna=False, observed=True).agg(agg).reset_index()
    out = out.rename(columns={"restrito": "n_semihoras", "num_minutos_restricao": "minutos"})
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
    frames = []
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
    out = pd.concat(frames, ignore_index=True)
    out = out.sort_values("capacidade_mw").groupby("id_ons", as_index=False).last()
    out["fonte"] = out["tipo"].map({"Solar": "Solar", "Eólica": "Eólica"}).fillna(out["tipo"])
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
