"""Hourly PLD (Preço de Liquidação das Diferenças) from CCEE Dados Abertos.

Two series are used:

* ``pld_horario_submercado`` - the settled hourly PLD per submercado.  CCEE
  only publishes it after the monthly accounting closes, so the last one or
  two months are always missing.
* ``pld_media_semanal``      - the weekly average PLD, published within days.
  Used only to fill the gap left by the hourly file, and always flagged as
  such so nobody mistakes an estimate for a settled price.

Both are exposed through the portal's CKAN datastore, which we read with the
``datastore/dump`` endpoint (plain CSV, no pagination, no API key).
"""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd

from .fetch import get
from .sources import CCEE_CKAN, CCEE_PLD_WEEKLY, SUBMERCADO_MAP


def _package(pkg: str) -> dict:
    r = get(f"{CCEE_CKAN}/api/3/action/package_show?id={pkg}")
    return r.json()["result"]


def _dump(resource_id: str) -> pd.DataFrame:
    """Read a CKAN datastore resource as CSV."""
    url = f"{CCEE_CKAN}/datastore/dump/{resource_id}"
    raw = get(url).content
    for sep in (",", ";"):
        try:
            df = pd.read_csv(io.BytesIO(raw), sep=sep, low_memory=False)
        except Exception:  # noqa: BLE001
            continue
        if df.shape[1] > 1:
            return df
    raise RuntimeError(f"could not parse CCEE resource {resource_id}")


def _norm_sub(s: pd.Series) -> pd.Series:
    return (
        s.astype(str).str.strip().str.upper().map(SUBMERCADO_MAP).fillna(s.astype(str).str.strip())
    )


def load_pld_horario(years: list[int]) -> pd.DataFrame:
    """Hourly PLD -> columns [hora, submercado, pld, pld_fonte]."""
    pkg = _package("pld_horario_submercado")
    frames = []
    for res in pkg["resources"]:
        name = (res.get("name") or "").lower()
        if not any(str(y) in name for y in years):
            continue
        df = _dump(res["id"])
        cols = {c.upper(): c for c in df.columns}
        need = {"MES_REFERENCIA", "SUBMERCADO", "PERIODO_COMERCIALIZACAO", "PLD"}
        if not need.issubset(cols):
            continue
        d = pd.DataFrame(
            {
                "mes": df[cols["MES_REFERENCIA"]].astype(str).str.strip(),
                "submercado": _norm_sub(df[cols["SUBMERCADO"]]),
                "periodo": pd.to_numeric(df[cols["PERIODO_COMERCIALIZACAO"]], errors="coerce"),
                "pld": pd.to_numeric(
                    df[cols["PLD"]].astype(str).str.replace(",", ".", regex=False),
                    errors="coerce",
                ),
            }
        ).dropna(subset=["periodo", "pld"])
        # PERIODO_COMERCIALIZACAO = 1-based hour inside the reference month.
        base = pd.to_datetime(d["mes"], format="%Y%m", errors="coerce")
        d = d[base.notna()]
        d["hora"] = base[base.notna()] + pd.to_timedelta(d["periodo"] - 1, unit="h")
        frames.append(d[["hora", "submercado", "pld"]])

    if not frames:
        return pd.DataFrame(columns=["hora", "submercado", "pld", "pld_fonte"])
    out = pd.concat(frames, ignore_index=True).drop_duplicates(["hora", "submercado"], keep="last")
    out["pld_fonte"] = "horario"
    return out


def load_pld_semanal(years: list[int]) -> pd.DataFrame:
    """Weekly average PLD -> [semana_inicio, submercado, pld]."""
    pkg = _package(CCEE_PLD_WEEKLY)
    frames = []
    for res in pkg["resources"]:
        name = (res.get("name") or "").lower()
        if not any(str(y) in name for y in years):
            continue
        df = _dump(res["id"])
        cols = {c.upper(): c for c in df.columns}
        if not {"SEMANA", "SUBMERCADO"}.issubset(cols):
            continue
        val = next((cols[c] for c in cols if c.startswith("PLD")), None)
        if val is None:
            continue
        d = pd.DataFrame(
            {
                "semana_inicio": pd.to_datetime(
                    df[cols["SEMANA"]], format="%d/%m/%Y", errors="coerce"
                ),
                "submercado": _norm_sub(df[cols["SUBMERCADO"]]),
                "pld": pd.to_numeric(
                    df[val].astype(str).str.replace(",", ".", regex=False), errors="coerce"
                ),
            }
        ).dropna()
        frames.append(d)
    if not frames:
        return pd.DataFrame(columns=["semana_inicio", "submercado", "pld"])
    return pd.concat(frames, ignore_index=True).drop_duplicates(
        ["semana_inicio", "submercado"], keep="last"
    )


def build_price_table(first: datetime, last: datetime) -> pd.DataFrame:
    """One row per hour x submercado covering [first, last], with a source flag.

    Hourly settled PLD wherever CCEE has published it; the weekly average
    carried across the remaining hours, flagged ``semanal_media``.
    """
    years = list(range(first.year, last.year + 1))
    hourly = load_pld_horario(years)
    weekly = load_pld_semanal(years)

    idx = pd.MultiIndex.from_product(
        [
            pd.date_range(first.normalize(), last.normalize() + pd.Timedelta(hours=23), freq="h"),
            ["N", "NE", "S", "SE"],
        ],
        names=["hora", "submercado"],
    )
    grid = pd.DataFrame(index=idx).reset_index()
    grid = grid.merge(hourly, on=["hora", "submercado"], how="left")

    if not weekly.empty:
        weekly = weekly.sort_values("semana_inicio")
        grid = grid.sort_values("hora")
        filled = []
        for sub, g in grid.groupby("submercado", sort=False):
            w = weekly[weekly["submercado"] == sub][["semana_inicio", "pld"]]
            if w.empty:
                filled.append(g)
                continue
            g = pd.merge_asof(
                g.sort_values("hora"),
                w.rename(columns={"pld": "pld_semana", "semana_inicio": "hora"}).sort_values(
                    "hora"
                ),
                on="hora",
                direction="backward",
            )
            need = g["pld"].isna() & g["pld_semana"].notna()
            g.loc[need, "pld"] = g.loc[need, "pld_semana"]
            g.loc[need, "pld_fonte"] = "semanal_media"
            filled.append(g.drop(columns=["pld_semana"]))
        grid = pd.concat(filled, ignore_index=True)

    grid["pld_fonte"] = grid["pld_fonte"].fillna("indisponivel")
    return grid[["hora", "submercado", "pld", "pld_fonte"]]


def apply_price(df: pd.DataFrame, prices: pd.DataFrame) -> pd.DataFrame:
    """Attach PLD and the monetary value of the curtailed energy.

    ``rs_gnr`` is the market value of the energy that was not generated
    (MWh x PLD).  It is an opportunity-cost figure, not the constrained-off
    compensation actually settled by CCEE - see the README.
    """
    out = df.merge(
        prices.rename(columns={"submercado": "id_subsistema"}),
        on=["hora", "id_subsistema"],
        how="left",
    )
    out["rs_gnr"] = out["mwh_gnr"] * out["pld"]
    return out
