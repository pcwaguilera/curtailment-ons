"""Build the Excel deliverable: solar only, restricted to the chosen conjuntos."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .sources import RAZAO, RAZAO_ORDER, SUBSISTEMA_NOME

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY = Font(name=FONT, size=10)
TITLE = Font(name=FONT, bold=True, size=13, color="1F3864")
NOTE = Font(name=FONT, size=9, italic=True, color="595959")
INPUT_FONT = Font(name=FONT, size=10, color="0000FF")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MWH = "#,##0.00"
BRL = 'R$ #,##0.00;[Red]-R$ #,##0.00;"-"'
PCT = "0.0%"
DATE = "dd/mm/yyyy"
DTM = "dd/mm/yyyy hh:mm"

EXCEL_ROW_LIMIT = 1_048_575  # minus the header


def _sheet(wb: Workbook, name: str):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    return ws


def _write_table(ws, df: pd.DataFrame, start_row: int, formats: dict[str, str],
                 widths: dict[str, int] | None = None) -> int:
    """Write a dataframe as a formatted table. Returns the last row written."""
    cols = list(df.columns)
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=start_row, column=j, value=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        letter = get_column_letter(j)
        ws.column_dimensions[letter].width = (widths or {}).get(c, max(11, min(30, len(c) + 4)))
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)

    values = df.itertuples(index=False, name=None)
    for i, row in enumerate(values, start=start_row + 1):
        for j, v in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=v)
            cell.font = BODY
            fmt = formats.get(cols[j - 1])
            if fmt:
                cell.number_format = fmt
    last = start_row + len(df)
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(cols))}{max(last, start_row)}"
    return last


# --------------------------------------------------------------------------
def _readme(wb: Workbook, selected: list[str], daily: pd.DataFrame,
            detail: pd.DataFrame, prices: pd.DataFrame) -> None:
    ws = wb.active
    ws.title = "Leia-me"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 96

    rows = [
        ("Curtailment (constrained-off) — usinas fotovoltaicas", None, TITLE),
        (None, None, None),
        ("Gerado em (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"), None),
        ("Conjuntos selecionados", ", ".join(selected) if selected
         else "(nenhum — edite config/conjuntos.yml)", None),
        ("Período coberto",
         f"{daily['data'].min():%d/%m/%Y} a {daily['data'].max():%d/%m/%Y}" if not daily.empty
         else "—", None),
        (None, None, None),
        ("Fontes de dados", None, Font(name=FONT, bold=True, size=11)),
        ("ONS — Restrição constrained-off FV",
         "https://dados.ons.org.br/dataset/restricao_coff_fotovoltaica", None),
        ("ONS — Detalhamento por usina",
         "https://dados.ons.org.br/dataset/restricao_coff_fotovoltaica_detail", None),
        ("ONS — Intra-semi-hora",
         "https://dados.ons.org.br/dataset/coff_fotovoltaica_intrasemihora", None),
        ("ONS — Fator de capacidade 2 (lat/long)",
         "https://dados.ons.org.br/dataset/fator-capacidade-2", None),
        ("CCEE — PLD horário por submercado",
         "https://dadosabertos.ccee.org.br/dataset/pld_horario_submercado", None),
        (None, None, None),
        ("Definições", None, Font(name=FONT, bold=True, size=11)),
        ("GNRa (MWmed)",
         "Geração Não Realizada Apurada: geração de referência menos geração verificada nos "
         "patamares com limitação (campo val_geracaonaorealizadaapurada do ONS).", None),
        ("Energia não gerada (MWh)",
         "GNRa × 0,5 h — cada registro do ONS é um patamar semi-horário em MWmed.", None),
        ("Valor (R$)",
         "Energia não gerada (MWh) × PLD horário do submercado na hora do evento. "
         "É o valor de mercado da energia perdida (custo de oportunidade), NÃO o "
         "ressarcimento de constrained-off apurado pela CCEE, que segue regras próprias "
         "e em geral remunera apenas parte das restrições.", None),
        (None, None, None),
        ("Razão da restrição", None, Font(name=FONT, bold=True, size=11)),
        *[(code, RAZAO[code], None) for code in RAZAO_ORDER],
        (None, None, None),
        ("Avisos", None, Font(name=FONT, bold=True, size=11)),
        ("Dados provisórios",
         "O ONS revisa os dados após a publicação; meses recentes mudam. Este arquivo é "
         "reconstruído todo dia a partir da fonte.", None),
        ("PLD estimado",
         "A CCEE só publica o PLD horário depois do fechamento mensal. Enquanto isso, as "
         "horas mais recentes usam a média semanal do PLD e vêm marcadas como "
         "'semanal_media' na coluna PLD (fonte). Elas são recalculadas automaticamente "
         "quando o PLD horário sai.", None),
        ("Cobertura do PLD neste arquivo",
         (prices["pld_fonte"].value_counts(normalize=True).mul(100).round(1).to_string()
          .replace("\n", " | ") if not prices.empty else "—"), None),
    ]
    r = 1
    for label, value, font in rows:
        if label is not None:
            c = ws.cell(row=r, column=1, value=label)
            c.font = font or Font(name=FONT, bold=True, size=10)
            c.alignment = Alignment(vertical="top")
        if value is not None:
            c = ws.cell(row=r, column=2, value=value)
            c.font = BODY
            c.alignment = Alignment(vertical="top", wrap_text=True)
        r += 1
    ws.cell(row=r + 1, column=1, value="Licença: dados ONS e CCEE sob CC-BY. "
                                       "Este arquivo é gerado automaticamente.").font = NOTE


# --------------------------------------------------------------------------
def write_excel(path: Path, daily: pd.DataFrame, detail: pd.DataFrame,
                conjuntos: pd.DataFrame, selected: list[str], prices: pd.DataFrame) -> None:
    solar = daily[daily["fonte"] == "Solar"].copy()
    if selected:
        solar = solar[solar["id_ons"].isin(selected) | solar["nom_usina"].isin(selected)]

    wb = Workbook()
    _readme(wb, selected, solar, detail, prices)

    # ---------------------------------------------------------------- Diário
    dia = (
        solar.pivot_table(
            index=["id_ons", "nom_usina", "id_subsistema", "id_estado", "data", "pld_fonte"],
            columns="cod_razaorestricao",
            values="mwh_gnr",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )
    for code in RAZAO_ORDER:
        if code not in dia.columns:
            dia[code] = 0.0
    money = solar.groupby(["id_ons", "data"], as_index=False)["rs_gnr"].sum()
    minutes = solar.groupby(["id_ons", "data"], as_index=False)["minutos"].sum()
    dia = dia.merge(money, on=["id_ons", "data"], how="left").merge(
        minutes, on=["id_ons", "data"], how="left"
    )
    dia["MWh total"] = dia[RAZAO_ORDER].sum(axis=1)
    dia = dia.rename(
        columns={
            "id_ons": "ID ONS", "nom_usina": "Conjunto", "id_subsistema": "Submercado",
            "id_estado": "UF", "data": "Data", "pld_fonte": "PLD (fonte)",
            "rs_gnr": "Valor (R$)", "minutos": "Minutos em restrição",
        }
    )
    dia = dia[["ID ONS", "Conjunto", "Submercado", "UF", "Data", *RAZAO_ORDER,
               "MWh total", "Valor (R$)", "Minutos em restrição", "PLD (fonte)"]]
    dia = dia.sort_values(["Conjunto", "Data"])
    ws_dia = _sheet(wb, "Diário")
    ws_dia.cell(row=1, column=1, value="Energia não gerada por conjunto e por dia (MWh) "
                                       "e valor a PLD").font = TITLE
    fmt = {c: MWH for c in [*RAZAO_ORDER, "MWh total"]}
    fmt.update({"Valor (R$)": BRL, "Data": DATE, "Minutos em restrição": "#,##0"})
    last_dia = _write_table(ws_dia, dia, 3, fmt, {"Conjunto": 30, "ID ONS": 16, "Data": 12})

    # ---------------------------------------------------------------- Resumo
    ws = _sheet(wb, "Resumo")
    ws.cell(row=1, column=1, value="Resumo por conjunto — usinas fotovoltaicas").font = TITLE
    ws.cell(row=2, column=1,
            value="Todos os valores são fórmulas SUMIFS sobre a aba 'Diário'.").font = NOTE
    heads = ["Conjunto", "ID ONS", "Submercado", "UF", "Capacidade (MW)",
             *[f"{c} (MWh)" for c in RAZAO_ORDER], "MWh total", "Valor (R$)",
             "R$/MWh médio", "Latitude", "Longitude"]
    for j, h in enumerate(heads, start=1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = 15
    ws.column_dimensions["A"].width = 32
    ws.freeze_panes = "A5"

    cj = conjuntos.set_index("id_ons")
    names = dia[["Conjunto", "ID ONS", "Submercado", "UF"]].drop_duplicates("ID ONS")
    rng_id = f"'Diário'!$A${4}:$A${max(last_dia,4)}"
    for i, row in enumerate(names.itertuples(index=False), start=5):
        conj, idons, sub, uf = row
        ws.cell(row=i, column=1, value=conj).font = BODY
        ws.cell(row=i, column=2, value=idons).font = BODY
        ws.cell(row=i, column=3, value=sub).font = BODY
        ws.cell(row=i, column=4, value=uf).font = BODY
        cap = float(cj["capacidade_mw"].get(idons, float("nan")))
        c = ws.cell(row=i, column=5, value=None if pd.isna(cap) else round(cap, 1))
        c.font, c.number_format = BODY, "#,##0.0"
        for k, code in enumerate(RAZAO_ORDER):
            col = get_column_letter(6 + k)
            src = get_column_letter(6 + k)  # same offset on 'Diário'
            c = ws.cell(row=i, column=6 + k,
                        value=f"=SUMIFS('Diário'!${src}$4:${src}${max(last_dia,4)},{rng_id},$B{i})")
            c.font, c.number_format = BODY, MWH
        c = ws.cell(row=i, column=10, value=f"=SUM(F{i}:I{i})")
        c.font, c.number_format = BODY, MWH
        c = ws.cell(row=i, column=11,
                    value=f"=SUMIFS('Diário'!$K$4:$K${max(last_dia,4)},{rng_id},$B{i})")
        c.font, c.number_format = BODY, BRL
        c = ws.cell(row=i, column=12, value=f'=IFERROR(K{i}/J{i},"")')
        c.font, c.number_format = BODY, BRL
        for k, key in enumerate(("lat", "lon")):
            v = cj[key].get(idons, None)
            c = ws.cell(row=i, column=13 + k, value=None if pd.isna(v) else float(v))
            c.font, c.number_format = BODY, "0.000000"
    tot = len(names) + 4
    if len(names):
        c = ws.cell(row=tot + 1, column=1, value="TOTAL")
        c.font = Font(name=FONT, bold=True, size=10)
        for col in "FGHIJK":
            c = ws.cell(row=tot + 1, column="ABCDEFGHIJK".index(col) + 1,
                        value=f"=SUM({col}5:{col}{tot})")
            c.font = Font(name=FONT, bold=True, size=10)
            c.number_format = BRL if col == "K" else MWH

    # ---------------------------------------------------------------- Mensal
    mensal = solar.copy()
    mensal["Mês"] = mensal["data"].dt.to_period("M").dt.to_timestamp()
    mensal = (
        mensal.groupby(["nom_usina", "id_ons", "Mês"], as_index=False)
        .agg(**{"MWh não gerado": ("mwh_gnr", "sum"), "Valor (R$)": ("rs_gnr", "sum"),
                "Minutos em restrição": ("minutos", "sum")})
        .rename(columns={"nom_usina": "Conjunto", "id_ons": "ID ONS"})
        .sort_values(["Conjunto", "Mês"])
    )
    ws_m = _sheet(wb, "Mensal")
    ws_m.cell(row=1, column=1, value="Totais mensais por conjunto").font = TITLE
    _write_table(ws_m, mensal, 3,
                 {"Mês": "mmm/yyyy", "MWh não gerado": MWH, "Valor (R$)": BRL,
                  "Minutos em restrição": "#,##0"}, {"Conjunto": 30})

    # -------------------------------------------------------- Semi-horário
    if detail is not None and not detail.empty:
        det = detail.copy()
        if selected:
            det = det[det["id_ons"].isin(selected) | det["nom_usina"].isin(selected)]
        det = det.sort_values(["nom_usina", "din_instante"])
        keep = {
            "din_instante": "Data/hora", "nom_usina": "Conjunto", "id_ons": "ID ONS",
            "id_subsistema": "Submercado", "id_estado": "UF",
            "val_geracaoreferencia": "Ger. referência (MWmed)",
            "val_geracao": "Ger. verificada (MWmed)",
            "val_geracaolimitada": "Ger. limitada (MWmed)",
            "val_geracaonaorealizadaapurada": "GNRa (MWmed)",
            "mwh_gnr": "Energia não gerada (MWh)",
            "cod_razaorestricao": "Razão", "cod_origemrestricao": "Origem",
            "dsc_restricao": "Descrição da restrição",
            "num_minutos_restricao": "Minutos em restrição",
            "pld": "PLD (R$/MWh)", "pld_fonte": "PLD (fonte)", "rs_gnr": "Valor (R$)",
        }
        det = det[[c for c in keep if c in det.columns]].rename(columns=keep)
        truncated = len(det) > EXCEL_ROW_LIMIT
        det = det.head(EXCEL_ROW_LIMIT)
        ws_d = _sheet(wb, "Semi-horário")
        ws_d.cell(row=1, column=1,
                  value="Detalhe semi-horário dos conjuntos selecionados").font = TITLE
        if truncated:
            ws_d.cell(row=2, column=1,
                      value="ATENÇÃO: excede o limite de linhas do Excel — truncado. "
                            "Reduza a lista de conjuntos ou use o .parquet do repositório."
                      ).font = Font(name=FONT, size=9, bold=True, color="C00000")
        _write_table(ws_d, det, 3,
                     {"Data/hora": DTM, "Energia não gerada (MWh)": MWH,
                      "PLD (R$/MWh)": BRL, "Valor (R$)": BRL,
                      "GNRa (MWmed)": MWH, "Ger. referência (MWmed)": MWH,
                      "Ger. verificada (MWmed)": MWH, "Ger. limitada (MWmed)": MWH,
                      "Minutos em restrição": "#,##0"},
                     {"Conjunto": 28, "Data/hora": 18, "Descrição da restrição": 60})

    # ---------------------------------------------------------------- PLD
    if not prices.empty:
        p = prices.rename(columns={"hora": "Hora", "submercado": "Submercado",
                                   "pld": "PLD (R$/MWh)", "pld_fonte": "Fonte"})
        p = p[p["Hora"] >= (solar["data"].min() if not solar.empty else p["Hora"].min())]
        ws_p = _sheet(wb, "PLD horário")
        ws_p.cell(row=1, column=1,
                  value="PLD horário por submercado — CCEE Dados Abertos").font = TITLE
        _write_table(ws_p, p.head(EXCEL_ROW_LIMIT), 3,
                     {"Hora": DTM, "PLD (R$/MWh)": BRL}, {"Hora": 18})

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"excel written: {path} ({path.stat().st_size/1e6:.1f} MB)")
