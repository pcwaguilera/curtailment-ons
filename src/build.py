"""Daily pipeline: download -> price -> aggregate -> publish.

Usage
-----
    python -m src.build            # incremental (only months whose ETag moved)
    python -m src.build --full     # rebuild everything from 2021-10 / 2024-04
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import yaml

from . import ccee, condicionantes, ons
from .excel_report import PLD_INICIO, write_excel
from .fetch import State
from .site_build import write_site
from .sources import ONS_DATASETS, ONS_S3

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
DOCS = ROOT / "docs"
CONFIG = ROOT / "config" / "conjuntos.yml"

TM_SETS = {"solar_tm": "Solar", "wind_tm": "Eólica"}


# --------------------------------------------------------------------------
def load_config() -> dict:
    if CONFIG.exists():
        return yaml.safe_load(CONFIG.read_text(encoding="utf-8")) or {}
    return {}


# As tabelas processadas ficam particionadas por mês (data/<nome>/YYYY_MM.parquet).
# Um arquivo único seria reescrito inteiro todo dia e o histórico do Git cresceria
# dezenas de MB por execução; assim só o mês que mudou entra no commit.
def _read_partitions(folder: Path) -> pd.DataFrame:
    if not folder.exists():
        return pd.DataFrame()
    parts = [pd.read_parquet(p) for p in sorted(folder.glob("*.parquet"))]
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def _write_partitions(folder: Path, df: pd.DataFrame, months: set[str]) -> None:
    """Grava (ou apaga) apenas as partições dos meses reprocessados."""
    folder.mkdir(parents=True, exist_ok=True)
    for mes in months:
        part = folder / f"{mes}.parquet"
        chunk = df[df["mes"] == mes] if not df.empty else df
        if chunk is None or chunk.empty:
            part.unlink(missing_ok=True)
        else:
            chunk.to_parquet(part, index=False)


def month_key(ts: pd.Timestamp) -> str:
    return f"{ts.year:04d}_{ts.month:02d}"


# --------------------------------------------------------------------------
def main(full: bool = False) -> None:
    DATA.mkdir(exist_ok=True)
    DOCS.mkdir(exist_ok=True)
    (DOCS / "data").mkdir(exist_ok=True)

    cfg = load_config()
    selected = [str(x).strip() for x in (cfg.get("conjuntos_solar") or [])]
    usinas_de = [str(x).strip() for x in (cfg.get("usinas_detalhe") or [])]
    state = State(DATA / "state.json")

    # ---------------------------------------------------------------- 1. index
    listing = {ds: ons.months_available(ds) for ds in TM_SETS}
    det_listing = ons.months_available("solar_detail") if usinas_de else {}
    all_months = sorted({m for d in listing.values() for m in d})
    if not all_months:
        raise SystemExit("ONS listing came back empty - aborting without touching data/")
    first = pd.Timestamp(all_months[0].replace("_", "-") + "-01")
    last_month = all_months[-1]

    # ---------------------------------------------------- 2. prices (rebuilt daily)
    # A aba de PLD do Excel começa em set/2023 mesmo que a série do ONS comece
    # depois, então a grade de preços nunca pode começar tarde demais.
    prices = ccee.build_price_table(min(first, PLD_INICIO),
                                    pd.Timestamp.utcnow().tz_localize(None))
    prices.to_parquet(DATA / "pld.parquet", index=False)

    pld_status: dict[str, str] = {}
    for mes, g in prices.assign(mes=prices["hora"].map(month_key)).groupby("mes"):
        if (g["pld_fonte"] == "horario").mean() > 0.99:
            pld_status[mes] = "horario"
        elif (g["pld_fonte"] != "indisponivel").any():
            pld_status[mes] = "semanal_media"
        else:
            pld_status[mes] = "indisponivel"
    prev_status = state.data.get("pld_status", {})

    # ------------------------------------------------- 3. which months to reload
    dirty: dict[str, set[str]] = {}
    for ds in TM_SETS:
        d = set()
        for mes, meta in listing[ds].items():
            if full:
                d.add(mes)
            elif state.etag(ds, mes) != meta["etag"]:
                d.add(mes)  # ONS republished the file
            elif prev_status.get(mes) != pld_status.get(mes):
                d.add(mes)  # CCEE published a better price for that month
            elif (ds == "solar_tm" and mes in det_listing
                  and state.etag("solar_detail", mes) != det_listing[mes]["etag"]):
                d.add(mes)  # o arquivo por usina mudou; precisamos do TM junto
        dirty[ds] = d

    print(f"months to (re)process: " + ", ".join(f"{k}={len(v)}" for k, v in dirty.items()))

    # ------------------------------------------------------------ 4. process
    DAILY_DIR = DATA / "agg_daily"
    DETAIL_DIR = DATA / "detail_selected"
    USINA_DIR = DATA / "usinas_daily"
    if full:
        for folder in (DAILY_DIR, DETAIL_DIR, USINA_DIR):
            for p in folder.glob("*.parquet") if folder.exists() else []:
                p.unlink()

    daily_new, detail_new, usina_new = [], [], []
    for ds, fonte in TM_SETS.items():
        for mes in sorted(dirty[ds]):
            meta = listing[ds][mes]
            print(f"  .. {ds} {mes} ({meta['size']/1e6:.1f} MB)")
            df = ons.load_tm(meta["url"], fonte)
            if df.empty:
                continue
            df = ccee.apply_price(df, prices)
            df["mes"] = mes

            d = ons.aggregate_daily(df)
            d["mes"] = mes
            # dominant price source of the day, so the UI can flag estimates
            src = (
                df.groupby(["id_ons", "data"])["pld_fonte"]
                .agg(lambda s: s.mode().iat[0] if len(s.mode()) else "indisponivel")
                .rename("pld_fonte")
                .reset_index()
            )
            src["data"] = pd.to_datetime(src["data"])
            d = d.merge(src, on=["id_ons", "data"], how="left")
            daily_new.append(d)

            if fonte == "Solar" and selected:
                # Guardamos TODOS os patamares dos conjuntos selecionados, não só os
                # restritos: a geração e a geração esperada mensais precisam do mês
                # inteiro, e o recálculo das condicionantes precisa dos campos de
                # limitação e disponibilidade patamar a patamar.
                sel = df[df["id_ons"].isin(selected) | df["nom_usina"].isin(selected)]
                if not sel.empty:
                    detail_new.append(condicionantes.aplicar(sel))

            # ---- detalhamento por usina (base diária) -----------------------
            if fonte == "Solar" and usinas_de and mes in det_listing:
                dmeta = det_listing[mes]
                print(f"     + usinas {mes} ({dmeta['size']/1e6:.1f} MB)")
                det = ons.load_detail(dmeta["url"], usinas_de)
                if not det.empty:
                    det = ons.merge_detail_razao(det, df)
                    det = ccee.apply_price(det, prices)
                    u = ons.aggregate_daily_usina(det)
                    u["mes"] = mes
                    usina_new.append(u)
                state.set_etag("solar_detail", mes, dmeta["etag"])

            state.set_etag(ds, mes, meta["etag"])

    _write_partitions(
        DAILY_DIR,
        pd.concat(daily_new, ignore_index=True) if daily_new else pd.DataFrame(),
        {m for s in dirty.values() for m in s},
    )
    _write_partitions(
        DETAIL_DIR,
        pd.concat(detail_new, ignore_index=True) if detail_new else pd.DataFrame(),
        dirty["solar_tm"],
    )
    _write_partitions(
        USINA_DIR,
        pd.concat(usina_new, ignore_index=True) if usina_new else pd.DataFrame(),
        dirty["solar_tm"],
    )
    daily = _read_partitions(DAILY_DIR)
    detail = _read_partitions(DETAIL_DIR)
    usinas = _read_partitions(USINA_DIR)

    if daily.empty:
        raise SystemExit("no aggregated data produced - aborting")

    daily = daily.sort_values(["fonte", "id_ons", "data"]).reset_index(drop=True)
    if not detail.empty:
        detail = detail.sort_values(["id_ons", "din_instante"]).reset_index(drop=True)

    # ------------------------------------------------- 5. conjunto master (map)
    fc = ons.months_available("fatorcap")
    fc_months = list(fc)[-3:] if not full else list(fc)[-24:]
    conjuntos = ons.load_conjuntos([fc[m]["url"] for m in fc_months])
    conjuntos.to_parquet(DATA / "conjuntos.parquet", index=False)

    # ------------------------------------------------------------ 6. outputs
    xlsx = DOCS / "curtailment_solar.xlsx"
    write_excel(xlsx, daily, detail, conjuntos, selected, prices, usinas)
    write_site(DOCS, daily, conjuntos, prices, selected, xlsx.name)

    state.save(
        pld_status=pld_status,
        last_run=datetime.now(timezone.utc).isoformat(timespec="seconds"),
        last_month=last_month,
    )
    print(f"done - {len(daily):,} daily rows, {len(detail):,} semi-hourly rows for the Excel")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--full", action="store_true", help="rebuild the whole history")
    main(**vars(ap.parse_args()))
