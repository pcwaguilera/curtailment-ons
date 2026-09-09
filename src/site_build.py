"""Render the public dashboard (docs/index.html) from the aggregated tables."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

from .sources import RAZAO_ORDER

TEMPLATE = Path(__file__).with_name("template.html")
FONTE_CODE = {"Solar": "S", "Eólica": "E"}


def _round(x, n=1):
    try:
        return round(float(x), n)
    except (TypeError, ValueError):
        return 0.0


def build_payload(daily: pd.DataFrame, conjuntos: pd.DataFrame, prices: pd.DataFrame,
                  xlsx_name: str | None = None) -> dict:
    d = daily.copy()
    d["fonte_cod"] = d["fonte"].map(FONTE_CODE).fillna("S")
    d["mes"] = d["data"].dt.strftime("%Y-%m")

    meta = conjuntos.set_index("id_ons")
    ids = sorted(d["id_ons"].unique())
    names = d.groupby("id_ons")["nom_usina"].last()
    subs = d.groupby("id_ons")["id_subsistema"].last()
    ufs = d.groupby("id_ons")["id_estado"].last()
    fontes = d.groupby("id_ons")["fonte_cod"].last()

    conj = []
    for i in ids:
        row = meta.loc[i] if i in meta.index else None
        lat = lon = cap = None
        if row is not None:
            lat = None if pd.isna(row["lat"]) else round(float(row["lat"]), 4)
            lon = None if pd.isna(row["lon"]) else round(float(row["lon"]), 4)
            cap = None if pd.isna(row["capacidade_mw"]) else round(float(row["capacidade_mw"]), 1)
        conj.append({
            "id": i,
            "nome": str(names.get(i, i)),
            "fonte": fontes.get(i, "S"),
            "sub": subs.get(i, ""),
            "uf": ufs.get(i, ""),
            "lat": lat, "lon": lon, "cap": cap,
        })
    idx = {c["id"]: n for n, c in enumerate(conj)}

    # ---- monthly per conjunto, split by reason -------------------------
    wide = d.pivot_table(index=["id_ons", "mes"], columns="cod_razaorestricao",
                         values="mwh_gnr", aggfunc="sum", fill_value=0.0)
    for code in RAZAO_ORDER:
        if code not in wide.columns:
            wide[code] = 0.0
    extra = d.groupby(["id_ons", "mes"]).agg(
        mwh=("mwh_gnr", "sum"), rs=("rs_gnr", "sum"), ref=("mwh_referencia", "sum")
    )
    wide = wide.join(extra)
    mensal = [
        [idx[i], mes, _round(r["mwh"]), _round(r["REL"]), _round(r["CNF"]),
         _round(r["ENE"]), _round(r["PAR"]), _round(r["rs"]), _round(r["ref"])]
        for (i, mes), r in wide.iterrows()
    ]

    # ---- daily per fonte x submercado ----------------------------------
    dd = d.pivot_table(index=["fonte_cod", "data", "id_subsistema"],
                       columns="cod_razaorestricao", values="mwh_gnr",
                       aggfunc="sum", fill_value=0.0)
    for code in RAZAO_ORDER:
        if code not in dd.columns:
            dd[code] = 0.0
    dd["t"] = dd[RAZAO_ORDER].sum(axis=1)
    diario = [
        [f, ts.strftime("%Y-%m-%d"), sub, _round(r["t"]), _round(r["REL"]),
         _round(r["CNF"]), _round(r["ENE"]), _round(r["PAR"])]
        for (f, ts, sub), r in dd.iterrows()
    ]

    aviso = ""
    if not prices.empty:
        est = prices[prices["pld_fonte"] == "semanal_media"]
        if not est.empty:
            ini = est["hora"].min().strftime("%d/%m/%Y")
            aviso = (
                f"<b>PLD estimado a partir de {ini}.</b> A CCEE só publica o PLD horário depois "
                "do fechamento mensal; até lá o valor usa a média semanal do PLD do submercado. "
                "Os números são recalculados automaticamente quando o PLD horário sai."
            )

    return {
        "meta": {
            "periodo": f"{d['data'].min():%d/%m/%Y} – {d['data'].max():%d/%m/%Y}",
            "atualizado": datetime.now(timezone.utc).strftime("%d/%m/%Y %H:%M UTC"),
            "xlsx": xlsx_name,
            "aviso": aviso,
        },
        "conjuntos": conj,
        "mensal": mensal,
        "diario": diario,
    }


def render(payload: dict) -> str:
    html = TEMPLATE.read_text(encoding="utf-8")
    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    # keep the JSON safe inside a <script> block
    blob = blob.replace("</", "<\\/")
    return html.replace("__PAYLOAD__", blob)


def write_site(docs: Path, daily: pd.DataFrame, conjuntos: pd.DataFrame,
               prices: pd.DataFrame, selected: list[str], xlsx_name: str | None) -> None:
    payload = build_payload(daily, conjuntos, prices, xlsx_name)
    docs.mkdir(parents=True, exist_ok=True)
    (docs / "index.html").write_text(render(payload), encoding="utf-8")
    (docs / "data").mkdir(exist_ok=True)
    (docs / "data" / "payload.json").write_text(
        json.dumps(payload, ensure_ascii=False), encoding="utf-8"
    )
    (docs / ".nojekyll").write_text("")
    size = (docs / "index.html").stat().st_size / 1e6
    print(f"site written: {docs/'index.html'} ({size:.1f} MB)")
