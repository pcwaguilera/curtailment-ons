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
from pathlib import Path
from datetime import datetime

import pandas as pd

from .fetch import UA, get
from .sources import CCEE_CKAN, CCEE_PLD_WEEKLY, SUBMERCADO_MAP


_SESSION = None


def _sessao() -> "requests.Session":
    """Sessão aquecida: o WAF da CCEE costuma exigir uma visita à home antes.

    A primeira requisição ao portal recebe os cookies do firewall; sem eles as
    chamadas de API são recusadas com 403 quando vêm de um datacenter.
    """
    global _SESSION
    if _SESSION is not None:
        return _SESSION
    import requests

    s = requests.Session()
    s.headers.update(UA)
    try:
        s.get(CCEE_CKAN, timeout=60)
    except Exception as exc:  # noqa: BLE001
        print(f"  ! não consegui aquecer a sessão da CCEE ({exc})")
    s.headers.update({"Referer": CCEE_CKAN + "/", "X-Requested-With": "XMLHttpRequest"})
    _SESSION = s
    return s


def diagnostico() -> list[tuple[str, str]]:
    """Testa cada porta de entrada da CCEE e devolve [(url, resultado)].

    Serve para descobrir, de dentro do runner, se o bloqueio é do WAF em toda a
    origem ou só em alguns caminhos — e se algum host alternativo responde.
    """
    import requests

    alvos = [
        (CCEE_CKAN + "/", "home do portal"),
        (CCEE_CKAN + "/api/3/action/status_show", "API CKAN (status)"),
        (CCEE_CKAN + "/api/3/action/package_show?id=pld_horario_submercado", "API CKAN (package)"),
        (CCEE_CKAN + "/dataset/pld_horario_submercado", "página do dataset"),
        ("https://pda-download.ccee.org.br/6pVzKbCKRHaGCCjrBbK29A/content", "download direto"),
    ]
    s = _sessao()
    out = []
    for url, rotulo in alvos:
        try:
            r = s.get(url, timeout=45, stream=True)
            amostra = next(r.iter_content(200), b"")[:80]
            out.append((rotulo, f"HTTP {r.status_code} · {r.headers.get('content-type','?')} "
                                f"· {amostra[:60]!r}"))
            r.close()
        except Exception as exc:  # noqa: BLE001
            out.append((rotulo, f"ERRO {type(exc).__name__}: {exc}"))
    return out


def _package(pkg: str) -> dict:
    r = get(f"{CCEE_CKAN}/api/3/action/package_show?id={pkg}", session=_sessao())
    return r.json()["result"]


def _dump(resource_id: str) -> pd.DataFrame:
    """Read a CKAN datastore resource as CSV."""
    url = f"{CCEE_CKAN}/datastore/dump/{resource_id}"
    raw = get(url, session=_sessao()).content
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


# --------------------------------------------------------------------------
# snapshot local
# --------------------------------------------------------------------------
# A CCEE bloqueia por faixa de IP: do runner do GitHub toda a origem responde
# 403 ("Acesso bloqueado"), inclusive o host de download. Como o PLD horário só
# muda uma vez por mês, quando a contabilidade fecha, a série fica versionada
# aqui e é atualizada a partir de uma máquina que a CCEE aceita.
PLD_DIR = Path(__file__).resolve().parent.parent / "data" / "pld"
PLD_HORARIO_CSV = PLD_DIR / "pld_horario.csv"      # hora,N,NE,S,SE  (wide)
PLD_SEMANAL_CSV = PLD_DIR / "pld_semanal.csv"      # semana,submercado,pld


def snapshot_horario() -> pd.DataFrame:
    """Lê o snapshot horário versionado -> [hora, submercado, pld, pld_fonte]."""
    if not PLD_HORARIO_CSV.exists():
        return pd.DataFrame(columns=["hora", "submercado", "pld", "pld_fonte"])
    w = pd.read_csv(PLD_HORARIO_CSV)
    w["hora"] = pd.to_datetime(w["hora"], format="%Y-%m-%dT%H", errors="coerce")
    w = w.dropna(subset=["hora"])
    d = w.melt(id_vars="hora", value_vars=["N", "NE", "S", "SE"],
               var_name="submercado", value_name="pld").dropna(subset=["pld"])
    d["pld"] = pd.to_numeric(d["pld"], errors="coerce")
    d["pld_fonte"] = "horario"
    return d.dropna(subset=["pld"])


def snapshot_semanal() -> pd.DataFrame:
    if not PLD_SEMANAL_CSV.exists():
        return pd.DataFrame(columns=["semana_inicio", "submercado", "pld"])
    d = pd.read_csv(PLD_SEMANAL_CSV)
    d["semana_inicio"] = pd.to_datetime(d["semana"], errors="coerce")
    d["pld"] = pd.to_numeric(d["pld"], errors="coerce")
    return d.dropna(subset=["semana_inicio", "pld"])[["semana_inicio", "submercado", "pld"]]


def salvar_snapshot(hourly: pd.DataFrame, weekly: pd.DataFrame) -> None:
    """Regrava os arquivos versionados a partir do que veio da rede."""
    PLD_DIR.mkdir(parents=True, exist_ok=True)
    if not hourly.empty:
        w = hourly.pivot_table(index="hora", columns="submercado", values="pld",
                               aggfunc="last").reset_index()
        for s in ("N", "NE", "S", "SE"):
            if s not in w.columns:
                w[s] = pd.NA
        w = w[["hora", "N", "NE", "S", "SE"]].sort_values("hora")
        w["hora"] = w["hora"].dt.strftime("%Y-%m-%dT%H")
        w.to_csv(PLD_HORARIO_CSV, index=False)
    if not weekly.empty:
        s = weekly.rename(columns={"semana_inicio": "semana"}).sort_values(
            ["semana", "submercado"])
        s["semana"] = pd.to_datetime(s["semana"]).dt.strftime("%Y-%m-%d")
        s[["semana", "submercado", "pld"]].to_csv(PLD_SEMANAL_CSV, index=False)


def _combinar(snap: pd.DataFrame, rede: pd.DataFrame, chaves: list[str]) -> pd.DataFrame:
    """Rede por cima do snapshot (mais fresca), snapshot preenchendo o resto."""
    if rede.empty:
        return snap
    if snap.empty:
        return rede
    return (pd.concat([snap, rede], ignore_index=True)
            .drop_duplicates(chaves, keep="last"))


def build_price_table(first: datetime, last: datetime) -> pd.DataFrame:
    """One row per hour x submercado covering [first, last], with a source flag.

    Hourly settled PLD wherever CCEE has published it; the weekly average
    carried across the remaining hours, flagged ``semanal_media``.
    """
    years = list(range(first.year, last.year + 1))

    # O PLD é enriquecimento: se a CCEE estiver fora do ar ou mudar o formato, o
    # build segue com os dados do ONS e o valor em R$ fica em branco, marcado
    # como indisponível. Não vale derrubar a apuração inteira por causa do preço.
    hourly, weekly = snapshot_horario(), snapshot_semanal()
    if not hourly.empty:
        print(f"  · PLD do snapshot: horário até {hourly['hora'].max():%d/%m/%Y}, "
              f"semanal até {weekly['semana_inicio'].max():%d/%m/%Y}"
              if not weekly.empty else "  · PLD do snapshot (só horário)")

    # Se esta máquina alcançar a CCEE (roda local, não no runner), atualiza o
    # snapshot na hora. No GitHub Actions isso falha com 403 e seguimos com o
    # arquivo versionado.
    try:
        rede_h = load_pld_horario(years)
        hourly = _combinar(hourly, rede_h, ["hora", "submercado"])
        atualizou = not rede_h.empty
    except Exception as exc:  # noqa: BLE001
        print(f"  ! CCEE inacessível para o PLD horário ({exc}); usando o snapshot")
        atualizou = False
    try:
        rede_s = load_pld_semanal(years)
        weekly = _combinar(weekly, rede_s, ["semana_inicio", "submercado"])
        atualizou = atualizou or not rede_s.empty
    except Exception as exc:  # noqa: BLE001
        print(f"  ! CCEE inacessível para o PLD semanal ({exc}); usando o snapshot")
    if atualizou:
        salvar_snapshot(hourly, weekly)
        print("  · snapshot de PLD atualizado a partir da CCEE")

    idx = pd.MultiIndex.from_product(
        [
            pd.date_range(first.normalize(), last.normalize() + pd.Timedelta(hours=23), freq="h"),
            ["N", "NE", "S", "SE"],
        ],
        names=["hora", "submercado"],
    )
    grid = pd.DataFrame(index=idx).reset_index()
    if hourly.empty:
        # Um merge com frame vazio quebra por dtype (object × datetime64).
        grid["pld"] = pd.Series(dtype="float64")
        grid["pld_fonte"] = pd.Series(dtype="object")
    else:
        hourly = hourly.copy()
        hourly["hora"] = pd.to_datetime(hourly["hora"])
        grid = grid.merge(hourly, on=["hora", "submercado"], how="left")

    if not weekly.empty:
        weekly = weekly.copy()
        weekly["semana_inicio"] = pd.to_datetime(weekly["semana_inicio"])
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
    if prices is None or prices.empty:
        out = df.copy()
        out["pld"] = float("nan")
        out["pld_fonte"] = "indisponivel"
    else:
        out = df.merge(
            prices.rename(columns={"submercado": "id_subsistema"}),
            on=["hora", "id_subsistema"],
            how="left",
        )
        out["pld_fonte"] = out["pld_fonte"].fillna("indisponivel")
    out["rs_gnr"] = out["mwh_gnr"] * out["pld"]
    return out
