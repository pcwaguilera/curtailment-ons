"""Build the Excel deliverable.

Estrutura do arquivo
--------------------
Leia-me · Meus conjuntos · Mês a mês · Condicionantes · Perdas R$ · Minutos ·
Conjunto x usinas · Usinas — resumo · uma aba por conjunto · abas de histórico
(2024, 2025 S1, 2025 S2, 2026) · uma aba por usina · Resumo/Mensal/Diário
(todos os conjuntos FV) · PLD horário.

A cadeia de números tem um sentido só: as abas de histórico guardam os dados
crus do ONS, as tabelas consolidadas somam delas por SUMIFS, e as abas por
conjunto somam das consolidadas.  Nada é digitado duas vezes, e qualquer
número pode ser rastreado até o patamar semi-horário que o gerou.
"""

from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .condicionantes import TOL_BASE, TOL_MW, TOL_PCT, VIGENCIA_NOVA, mensal as cond_mensal
from .sources import RAZAO, RAZAO_ORDER, SUBSISTEMA_NOME

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="DDE5F0")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY = Font(name=FONT, size=10)
BOLD = Font(name=FONT, bold=True, size=10)
TITLE = Font(name=FONT, bold=True, size=13, color="1F3864")
SECTION = Font(name=FONT, bold=True, size=11, color="1F3864")
NOTE = Font(name=FONT, size=9, italic=True, color="595959")
BLUE = Font(name=FONT, size=10, color="0000FF")
RED = Font(name=FONT, bold=True, size=10, color="C00000")

MWH = "#,##0.0"
BRL = 'R$ #,##0.00;[Red]-R$ #,##0.00;"-"'
PCT = "0.0%"
DATE = "dd/mm/yyyy"
MES = "mmm/yyyy"
DTM = "dd/mm/yyyy hh:mm"
INT = "#,##0"

EXCEL_ROW_LIMIT = 1_048_000

# A aba de PLD começa aqui, independentemente de quando começa a série do ONS.
PLD_INICIO = pd.Timestamp("2023-09-01")

DIA_HEADERS = ["ID ONS", "Conjunto", "Submercado", "UF", "Data",
               *RAZAO_ORDER, "MWh total", "Valor (R$)", "Minutos em restrição",
               "PLD (fonte)"]
DIA_ID = "A"
DIA_RZ = {c: get_column_letter(6 + i) for i, c in enumerate(RAZAO_ORDER)}
DIA_RS = "K"


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def _sheet(wb: Workbook, name: str):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    return ws


def _safe_name(nome: str, taken: set[str]) -> str:
    s = re.sub(r"^Conj\.\s*", "", nome)
    s = re.sub(r"[\[\]:*?/\\]", "-", s)[:31].strip()
    base, n = s, 2
    while s in taken or not s:
        s = f"{base[:28]} {n}"
        n += 1
    taken.add(s)
    return s


def _head(ws, headers, row, widths=None):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=j, value=h)
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(h, 14)


def _table(ws, df: pd.DataFrame, start_row: int, formats: dict,
           widths: dict | None = None, autofilter: bool = True) -> int:
    _head(ws, list(df.columns), start_row, widths)
    for i, row in enumerate(df.itertuples(index=False, name=None), start=start_row + 1):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = BODY
            fmt = formats.get(df.columns[j - 1])
            if fmt:
                c.number_format = fmt
    last = start_row + len(df)
    if autofilter:
        ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(df.columns))}{max(last, start_row)}"
    return last


def _pct_col(ws, col: int, header: str, first: int, last: int,
             num_col: str, den_col: str) -> None:
    c = ws.cell(row=first - 1, column=col, value=header)
    c.font, c.fill = HDR_FONT, HDR_FILL
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.column_dimensions[get_column_letter(col)].width = 12
    for r in range(first, last + 1):
        cell = ws.cell(row=r, column=col, value=f'=IFERROR({num_col}{r}/{den_col}{r},"")')
        cell.font, cell.number_format = BODY, PCT


# --------------------------------------------------------------------------
# Leia-me
# --------------------------------------------------------------------------
def _readme(ws, sel_meta: pd.DataFrame, missing: list[str], solar: pd.DataFrame,
            prices: pd.DataFrame) -> None:
    ws.title = "Leia-me"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 104

    cobertura = "—"
    if not prices.empty:
        share = prices["pld_fonte"].value_counts(normalize=True).mul(100).round(1)
        cobertura = " | ".join(f"{k}: {v}%" for k, v in share.items())
    base_tol = {"verificada": "Geração Verificada", "limitada": "Geração Limitada",
                "capacidade": "capacidade instalada"}[TOL_BASE]

    rows = [
        ("Curtailment (constrained-off) — usinas fotovoltaicas", None, TITLE),
        (None, None, None),
        ("Gerado em (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"), None),
        ("Período coberto",
         f"{solar['data'].min():%d/%m/%Y} a {solar['data'].max():%d/%m/%Y}"
         if not solar.empty else "—", None),
        (None, None, None),
        ("As abas", None, SECTION),
        ("Meus conjuntos", "Totais e mês a mês dos seus conjuntos, em uma página.", None),
        ("Mês a mês", "Geração, geração esperada e curtailment por razão, com os %.", None),
        ("Condicionantes", "Curtailment publicado × recalculado com o teste de tolerância.", None),
        ("Perdas R$", "Valor da energia não gerada, aberto por razão.", None),
        ("Minutos", "Minutos em restrição por razão, mês a mês.", None),
        ("(nome do conjunto)", "Uma aba por conjunto, com as quatro tabelas empilhadas.", None),
        ("Histórico 2024 / 2025 S1 / 2025 S2 / 2026",
         "Os dados crus do ONS, patamar a patamar (30 em 30 min), nas colunas originais "
         "mais as calculadas. TODAS as abas de resumo somam daqui por SUMIFS — mexeu no "
         "histórico, o resumo acompanha.", None),
        ("U. (nome da usina)", "Histórico diário de cada usina, do arquivo de "
                               "detalhamento por usina.", None),
        ("Usinas — resumo", "Mês a mês de cada usina, somado da aba dela.", None),
        ("Conjunto x usinas", "A comparação: curtailment do conjunto × soma das usinas.", None),
        ("Resumo / Mensal / Diário", "Todos os conjuntos fotovoltaicos, para comparação.", None),
        ("PLD horário",
         "PLD de set/2023 em diante, uma linha por hora e os quatro submercados lado a "
         "lado. A coluna Fonte diz se o preço é o horário liquidado ou a média semanal "
         "usada enquanto o mês não fecha.", None),
        (None, None, None),
        ("Seus conjuntos", None, SECTION),
    ]
    if not sel_meta.empty:
        for _, r in sel_meta.iterrows():
            rows.append((r["id_ons"],
                         f"{r['nom_usina']} — {r['id_estado']}, "
                         f"{SUBSISTEMA_NOME.get(r['id_subsistema'], r['id_subsistema'])}", None))
    else:
        rows.append(("(vazio)", "Nada listado em config/conjuntos.yml.", None))
    if missing:
        rows.append(("NÃO ENCONTRADOS",
                     "Itens de config/conjuntos.yml sem correspondência no período: "
                     + ", ".join(missing), RED))

    rows += [
        (None, None, None),
        ("Definições", None, SECTION),
        ("Geração (MWh)", "val_geracao × 0,5 h — geração verificada, todos os patamares.", None),
        ("Geração esperada (MWh)",
         "val_geracaoreferencia × 0,5 h — a Geração de Referência do ONS, ou seja, quanto a "
         "usina teria gerado sem limitação.", None),
        ("Curtailment (MWh)",
         "val_geracaonaorealizadaapurada (GNRa) × 0,5 h. GNRa = max(0, Geração de Referência − "
         "Geração Verificada) nos patamares em que o ONS declarou razão de restrição. "
         "É o número publicado, SEM condicionante.", None),
        ("% de curtailment", "Curtailment ÷ geração esperada, no mesmo mês e conjunto.", None),
        ("Perda (R$)",
         "Curtailment × PLD horário do submercado na hora do evento. Valor de mercado da "
         "energia perdida (custo de oportunidade), NÃO o ressarcimento de constrained-off "
         "apurado pela CCEE, que segue regras próprias e em geral remunera apenas parte "
         "das restrições.", None),
        (None, None, None),
        ("Condicionantes — como foi calculado", None, SECTION),
        ("Documento",
         "RO-AO.BR.13 Rev.09 (vigência 18/06/2026), itens 4.13 e 5.2.2.10. "
         "https://www.ons.org.br → MPO → Rotinas Operacionais SM 5.13 → Pós-Operação → "
         "Apuração de Dados", None),
        ("Teste de tolerância (item 4.13)",
         f"tolerância = min({TOL_PCT:.0%} × {base_tol} ; {TOL_MW:.0f} MW). "
         "desvio = Geração Limitada − Geração Verificada. "
         "Atendido quando |desvio| ≤ tolerância.", None),
        ("Regra antiga (até jul/2025)",
         "Tolerância não atendida ⇒ o patamar deixava de contar: curtailment = 0. "
         "Atendida ⇒ curtailment = max(0, Ger. de Referência − Ger. Verificada).", None),
        ("Metodologia nova (item 5.2.2.10)",
         "a) G_ref_Disp = min(Ger. de Referência ; Disponibilidade eletromecânica). "
         "b) tolerância atendida ⇒ G_Ref_Final = G_ref_Disp; não atendida ⇒ "
         "G_Ref_Final = G_ref_Disp − (Ger. Limitada − Ger. Verificada). "
         "c) G_Ref_Final < 0 ⇒ 0. Curtailment = max(0, G_Ref_Final − Ger. Verificada). "
         "Em vez de zerar o patamar, desconta o desvio.", None),
        ("Regra vigente",
         f"Antiga até jul/2025; metodologia nova de {VIGENCIA_NOVA:%m/%Y} em diante. "
         "A aba Condicionantes traz também a metodologia nova aplicada a TODOS os meses e "
         "a diferença entre as duas.", None),
        ("Isto é reconstrução, não apuração oficial",
         "Partimos da Geração de Referência já publicada pelo ONS — que é exatamente onde o "
         "item 5.2.2.10 começa. O ONS calcula essa referência por função de produtividade e "
         "faz o rateio por usina do item 5.2.2.11. Diferenças contra a apuração oficial são "
         "esperadas; use estes números como estimativa e ordem de grandeza.", RED),
        ("Patamares sem Geração Limitada",
         "Quando o ONS declara razão de restrição mas não publica val_geracaolimitada, o teste "
         "de tolerância não é aplicável e o patamar é tratado como dentro da tolerância. "
         "A contagem desses casos está na aba Condicionantes.", None),
        (None, None, None),
        ("Razão da restrição", None, SECTION),
        *[(code, RAZAO[code], None) for code in RAZAO_ORDER],
        (None, None, None),
        ("Fontes", None, SECTION),
        ("ONS — constrained-off FV",
         "https://dados.ons.org.br/dataset/restricao_coff_fotovoltaica", None),
        ("ONS — detalhamento por usina",
         "https://dados.ons.org.br/dataset/restricao_coff_fotovoltaica_detail", None),
        ("ONS — intra-semi-hora",
         "https://dados.ons.org.br/dataset/coff_fotovoltaica_intrasemihora", None),
        ("ONS — fator de capacidade 2 (lat/long)",
         "https://dados.ons.org.br/dataset/fator-capacidade-2", None),
        ("CCEE — PLD horário por submercado",
         "https://dadosabertos.ccee.org.br/dataset/pld_horario_submercado", None),
        ("CCEE — PLD médio semanal",
         "https://dadosabertos.ccee.org.br/dataset/pld_media_semanal", None),
        (None, None, None),
        ("Avisos", None, SECTION),
        ("Dados provisórios",
         "O ONS revisa os dados após a publicação; meses recentes mudam. O arquivo é "
         "reconstruído todo dia a partir da fonte.", None),
        ("PLD estimado",
         "A CCEE só publica o PLD horário depois do fechamento mensal. Até lá as horas mais "
         "recentes usam a média semanal, marcadas como 'semanal_media'. São recalculadas "
         "sozinhas quando o PLD horário sai.", None),
        ("Cobertura do PLD neste arquivo", cobertura, None),
        ("Cores", "Números em azul vêm dos arquivos do ONS/CCEE; os pretos são fórmulas.", NOTE),
    ]

    r = 1
    for label, value, font in rows:
        if label is not None:
            c = ws.cell(row=r, column=1, value=label)
            c.font = font if font in (TITLE, SECTION, RED) else (font or BOLD)
            c.alignment = Alignment(vertical="top")
        if value is not None:
            c = ws.cell(row=r, column=2, value=value)
            c.font = NOTE if font is NOTE else BODY
            c.alignment = Alignment(vertical="top", wrap_text=True)
        r += 1


# --------------------------------------------------------------------------
# abas de histórico (semi-horário, dados crus do ONS + colunas calculadas)
# --------------------------------------------------------------------------
# As 11 primeiras colunas são exatamente os campos do ONS pedidos pela usuária,
# na ordem dela.  As seguintes são calculadas e existem para as abas de resumo
# poderem somar por SUMIFS — nenhum número da planilha é digitado duas vezes.
HIST_ONS = ["nom_usina", "din_instante", "val_geracao", "val_geracaolimitada",
            "val_disponibilidade", "val_geracaoreferencia", "val_geracaoreferenciafinal",
            "cod_razaorestricao", "cod_origemrestricao", "dsc_restricao",
            "val_geracaonaorealizadaapurada"]
HIST_CALC = ["id_ons", "mes_ref", "mwh_gerado", "mwh_referencia", "mwh_gnr", "minutos",
             "pld", "perda_rs", "mwh_gnr_vigente", "mwh_gnr_nova", "mwh_gnr_antiga",
             "tol_atendida", "gnr_origem"]
HIST_COLS = HIST_ONS + HIST_CALC
# letras (1-based) das colunas usadas nas fórmulas
HC = {c: get_column_letter(i + 1) for i, c in enumerate(HIST_COLS)}
HIST_ROW0 = 3  # cabeçalho na linha 2, dados a partir da 3

HIST_TITULOS = {
    "nom_usina": "Conjunto", "din_instante": "Data/hora",
    "val_geracao": "Geração verificada (MWmed)", "val_geracaolimitada": "Geração limitada (MWmed)",
    "val_disponibilidade": "Disponibilidade (MWmed)",
    "val_geracaoreferencia": "Geração de referência (MWmed)",
    "val_geracaoreferenciafinal": "Geração de referência final (MWmed)",
    "cod_razaorestricao": "Razão", "cod_origemrestricao": "Origem",
    "dsc_restricao": "Descrição da restrição",
    "val_geracaonaorealizadaapurada": "GNRa (MWmed)",
    "id_ons": "ID ONS", "mes_ref": "Mês (chave)", "mwh_gerado": "Geração (MWh)",
    "mwh_referencia": "Geração esperada (MWh)", "mwh_gnr": "Curtailment (MWh)",
    "minutos": "Minutos", "pld": "PLD (R$/MWh)", "perda_rs": "Perda (R$)",
    "mwh_gnr_vigente": "Curt. condicionantes vigente (MWh)",
    "mwh_gnr_nova": "Curt. metodologia nova (MWh)",
    "mwh_gnr_antiga": "Curt. regra antiga (MWh)",
    "tol_atendida": "Tolerância atendida", "gnr_origem": "Origem do GNRa",
}


def _periodo(ts: pd.Timestamp) -> str:
    """Qual aba de histórico guarda esta data."""
    if ts.year <= 2024:
        return "Histórico 2024"
    if ts.year == 2025:
        return "Histórico 2025 S1" if ts.month <= 6 else "Histórico 2025 S2"
    return f"Histórico {ts.year}"


def _hist_frame(detail: pd.DataFrame) -> pd.DataFrame:
    d = detail.copy()
    d["mes_ref"] = d["din_instante"].dt.strftime("%Y-%m")
    for c in HIST_COLS:
        if c not in d.columns:
            d[c] = pd.NA
    d = d.rename(columns={"num_minutos_restricao": "minutos"}) \
        if "minutos" not in d.columns else d
    if "perda_rs" not in detail.columns:
        d["perda_rs"] = d.get("rs_gnr", pd.Series(0.0, index=d.index))
    d["minutos"] = d.get("minutos", d.get("num_minutos_restricao", 0.0))
    d["tol_atendida"] = d.get("tol_atendida", pd.Series(True, index=d.index))
    return d[HIST_COLS].sort_values(["nom_usina", "din_instante"])


def _write_hist(wb: Workbook, detail: pd.DataFrame) -> dict[str, dict]:
    """Escreve as abas de histórico. Devolve {mes_ref: {'aba':…, 'fim':…}}."""
    if detail is None or detail.empty:
        return {}
    df = _hist_frame(detail)
    df["_aba"] = df["din_instante"].map(_periodo)
    mapa: dict[str, dict] = {}
    for aba, g in df.groupby("_aba", sort=True):
        ws = _sheet(wb, aba[:31])
        ws.cell(row=1, column=1,
                value=f"Histórico semi-horário do ONS — {aba.replace('Histórico ','')} "
                      f"({len(g):,} patamares)".replace(",", ".")).font = TITLE
        heads = [HIST_TITULOS[c] for c in HIST_COLS]
        for j, h in enumerate(heads, start=1):
            c = ws.cell(row=2, column=j, value=h)
            c.font, c.fill = HDR_FONT, HDR_FILL
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(j)].width = 15
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions[HC["dsc_restricao"]].width = 48
        ws.freeze_panes = "C3"

        g = g.drop(columns=["_aba"])
        # sem estilo por célula: são centenas de milhares de linhas e o estilo
        # individual multiplicaria o tamanho do arquivo. A fonte vem do estilo
        # Normal (Arial), definido em write_excel.
        for row in g.itertuples(index=False, name=None):
            ws.append(list(row))
        fim = 2 + len(g)
        for c, fmt in (("din_instante", DTM), ("mwh_gerado", MWH), ("mwh_referencia", MWH),
                       ("mwh_gnr", MWH), ("pld", BRL), ("perda_rs", BRL),
                       ("mwh_gnr_vigente", MWH), ("mwh_gnr_nova", MWH),
                       ("mwh_gnr_antiga", MWH), ("minutos", INT)):
            ws.column_dimensions[HC[c]].number_format = fmt
        ws.auto_filter.ref = f"A2:{get_column_letter(len(HIST_COLS))}{fim}"
        for mes in sorted(g["mes_ref"].dropna().unique()):
            mapa[str(mes)] = {"aba": aba[:31], "fim": fim}
    return mapa


def _sumifs(mapa: dict, mes: str, col: str, idons: str, razao: str | None = None) -> str:
    """SUMIFS de uma coluna do histórico, para um conjunto e um mês."""
    info = mapa.get(mes)
    if not info:
        return "0"
    a, fim = info["aba"], max(info["fim"], HIST_ROW0)
    rng = f"'{a}'!${HC[col]}${HIST_ROW0}:${HC[col]}${fim}"
    crit = (f"'{a}'!${HC['id_ons']}${HIST_ROW0}:${HC['id_ons']}${fim},\"{idons}\","
            f"'{a}'!${HC['mes_ref']}${HIST_ROW0}:${HC['mes_ref']}${fim},\"{mes}\"")
    if razao:
        crit += (f",'{a}'!${HC['cod_razaorestricao']}${HIST_ROW0}:"
                 f"${HC['cod_razaorestricao']}${fim},\"{razao}\"")
    return f"=SUMIFS({rng},{crit})"


def _countifs(mapa: dict, mes: str, idons: str, extra: str = "") -> str:
    info = mapa.get(mes)
    if not info:
        return "0"
    a, fim = info["aba"], max(info["fim"], HIST_ROW0)
    crit = (f"'{a}'!${HC['id_ons']}${HIST_ROW0}:${HC['id_ons']}${fim},\"{idons}\","
            f"'{a}'!${HC['mes_ref']}${HIST_ROW0}:${HC['mes_ref']}${fim},\"{mes}\"")
    return f"=COUNTIFS({crit}{extra})"


def _crit_col(mapa: dict, mes: str, col: str, valor: str) -> str:
    info = mapa[mes]
    a, fim = info["aba"], max(info["fim"], HIST_ROW0)
    return f",'{a}'!${HC[col]}${HIST_ROW0}:${HC[col]}${fim},{valor}"


# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# abas consolidadas — agora tudo por fórmula sobre as abas de histórico
# --------------------------------------------------------------------------
def _consolidada(ws, titulo: str, subtitulo: str, sel_meta: pd.DataFrame, meses: list[str],
                 heads: list[str], celulas, formatos: dict, widths: dict) -> int:
    """Emite uma tabela conjunto × mês onde cada número é uma fórmula.

    `celulas(idons, mes)` devolve a lista de fórmulas/valores das colunas D em diante.
    """
    ws.cell(row=1, column=1, value=titulo).font = TITLE
    ws.cell(row=2, column=1, value=subtitulo).font = NOTE
    _head(ws, heads, 4, widths)
    ws.freeze_panes = "D5"
    r = 5
    for _, m in sel_meta.iterrows():
        for mes in meses:
            ws.cell(row=r, column=1, value=m["nom_usina"]).font = BODY
            ws.cell(row=r, column=2, value=m["id_ons"]).font = BODY
            c = ws.cell(row=r, column=3,
                        value=datetime.strptime(mes + "-01", "%Y-%m-%d"))
            c.font, c.number_format = BODY, MES
            for j, val in enumerate(celulas(m["id_ons"], mes, r), start=4):
                c = ws.cell(row=r, column=j, value=val)
                c.font = BODY
                fmt = formatos.get(heads[j - 1])
                if fmt:
                    c.number_format = fmt
            r += 1
    last = r - 1
    ws.auto_filter.ref = f"A4:{get_column_letter(len(heads))}{max(last,4)}"
    return last


# --------------------------------------------------------------------------
# abas por usina
# --------------------------------------------------------------------------
U_COLS = ["nom_usina", "data", "mwh_gerado", "mwh_referencia", "mwh_gnr", "minutos",
          "rs_gnr", "cod_razaorestricao", "id_ons", "mes_ref", "nom_conjuntousina"]
U_TIT = {"nom_usina": "Usina", "data": "Data", "mwh_gerado": "Geração (MWh)",
         "mwh_referencia": "Geração estimada (MWh)", "mwh_gnr": "Curtailment (MWh)",
         "minutos": "Minutos", "rs_gnr": "Perda (R$)", "cod_razaorestricao": "Razão",
         "id_ons": "ID ONS", "mes_ref": "Mês (chave)",
         "nom_conjuntousina": "Conjunto"}
UC = {c: get_column_letter(i + 1) for i, c in enumerate(U_COLS)}
U_ROW0 = 3


def _usina_tabs(wb: Workbook, usinas: pd.DataFrame, taken: set[str]) -> dict:
    """Uma aba de histórico diário por usina. Devolve {id_ons: (aba, última linha)}."""
    if usinas is None or usinas.empty:
        return {}
    u = usinas.copy()
    u["data"] = pd.to_datetime(u["data"])
    u["mes_ref"] = u["data"].dt.strftime("%Y-%m")
    for c in U_COLS:
        if c not in u.columns:
            u[c] = pd.NA
    mapa = {}
    ordem = (u.groupby(["nom_conjuntousina", "nom_usina", "id_ons"], as_index=False)["mwh_gnr"]
             .sum().sort_values(["nom_conjuntousina", "nom_usina"]))
    for _, row in ordem.iterrows():
        g = u[u["id_ons"] == row["id_ons"]][U_COLS].sort_values("data")
        nome = _safe_name("U. " + str(row["nom_usina"]), taken)
        ws = _sheet(wb, nome)
        ws.cell(row=1, column=1,
                value=f"{row['nom_usina']} · {row['id_ons']} · {row['nom_conjuntousina']} "
                      f"— histórico diário").font = TITLE
        for j, c in enumerate(U_COLS, start=1):
            cell = ws.cell(row=2, column=j, value=U_TIT[c])
            cell.font, cell.fill = HDR_FONT, HDR_FILL
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(j)].width = 15
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions[UC["nom_conjuntousina"]].width = 28
        ws.freeze_panes = "C3"
        for rec in g.itertuples(index=False, name=None):
            ws.append(list(rec))
        fim = 2 + len(g)
        for c, fmt in (("data", DATE), ("mwh_gerado", MWH), ("mwh_referencia", MWH),
                       ("mwh_gnr", MWH), ("rs_gnr", BRL), ("minutos", INT)):
            ws.column_dimensions[UC[c]].number_format = fmt
        ws.auto_filter.ref = f"A2:{get_column_letter(len(U_COLS))}{fim}"
        mapa[row["id_ons"]] = {"aba": nome, "fim": fim, "usina": row["nom_usina"],
                               "conjunto": row["nom_conjuntousina"]}
    return mapa


def _usina_sumifs(info: dict, col: str, mes: str, razao: str | None = None) -> str:
    a, fim = info["aba"], max(info["fim"], U_ROW0)
    rng = f"'{a}'!${UC[col]}${U_ROW0}:${UC[col]}${fim}"
    crit = f"'{a}'!${UC['mes_ref']}${U_ROW0}:${UC['mes_ref']}${fim},\"{mes}\""
    if razao:
        crit += (f",'{a}'!${UC['cod_razaorestricao']}${U_ROW0}:"
                 f"${UC['cod_razaorestricao']}${fim},\"{razao}\"")
    return f"=SUMIFS({rng},{crit})"


def _usina_resumo(ws, umapa: dict, meses: list[str], mm_map: dict) -> None:
    ws.cell(row=1, column=1, value="Usinas — resumo mensal").font = TITLE
    ws.cell(row=2, column=1,
            value="Cada número é um SUMIFS na aba da própria usina. A coluna "
                  "'Conjunto (MWh)' vem da aba 'Mês a mês', para a comparação.").font = NOTE
    heads = ["Conjunto", "Usina", "ID ONS", "Mês", "Geração (MWh)",
             "Geração estimada (MWh)", "Curtailment (MWh)", "% do potencial",
             "Curt. REL (MWh)", "Curt. CNF (MWh)", "Curt. ENE (MWh)", "Perda (R$)",
             "Minutos"]
    _head(ws, heads, 4, {"Conjunto": 28, "Usina": 26, "ID ONS": 13, "Mês": 11,
                         "Geração estimada (MWh)": 16, "Curtailment (MWh)": 15})
    ws.freeze_panes = "E5"
    r = 5
    for idons, info in umapa.items():
        for mes in meses:
            ws.cell(row=r, column=1, value=info["conjunto"]).font = BODY
            ws.cell(row=r, column=2, value=info["usina"]).font = BODY
            ws.cell(row=r, column=3, value=idons).font = BODY
            c = ws.cell(row=r, column=4, value=datetime.strptime(mes + "-01", "%Y-%m-%d"))
            c.font, c.number_format = BODY, MES
            vals = [
                (_usina_sumifs(info, "mwh_gerado", mes), MWH),
                (_usina_sumifs(info, "mwh_referencia", mes), MWH),
                (_usina_sumifs(info, "mwh_gnr", mes), MWH),
                (f'=IFERROR(G{r}/F{r},"")', PCT),
                (_usina_sumifs(info, "mwh_gnr", mes, "REL"), MWH),
                (_usina_sumifs(info, "mwh_gnr", mes, "CNF"), MWH),
                (_usina_sumifs(info, "mwh_gnr", mes, "ENE"), MWH),
                (_usina_sumifs(info, "rs_gnr", mes), BRL),
                (_usina_sumifs(info, "minutos", mes), INT),
            ]
            for j, (v, fmt) in enumerate(vals, start=5):
                c = ws.cell(row=r, column=j, value=v)
                c.font, c.number_format = BODY, fmt
            r += 1
    ws.auto_filter.ref = f"A4:{get_column_letter(len(heads))}{max(r-1,4)}"


def _usina_comparacao(ws, umapa: dict, sel_meta: pd.DataFrame, meses: list[str],
                      mm_last: int) -> None:
    """Soma das usinas × total do conjunto, mês a mês — a comparação pedida."""
    ws.cell(row=1, column=1, value="Conjunto × soma das usinas").font = TITLE
    ws.cell(row=2, column=1,
            value="O curtailment do conjunto sai de val_geracaonaorealizadaapurada; o das "
                  "usinas, de geração estimada − verificada no arquivo de detalhamento. "
                  "São apurações diferentes, então alguma diferença é esperada.").font = NOTE
    heads = ["Conjunto", "ID ONS", "Mês", "Curtailment do conjunto (MWh)",
             "Soma das usinas (MWh)", "Diferença (MWh)", "Diferença (%)", "Usinas"]
    _head(ws, heads, 4, {"Conjunto": 30, "ID ONS": 14, "Mês": 11,
                         "Curtailment do conjunto (MWh)": 18, "Soma das usinas (MWh)": 17,
                         "Diferença (MWh)": 14, "Diferença (%)": 12, "Usinas": 8})
    ws.freeze_panes = "D5"
    porconj: dict[str, list] = {}
    for idons, info in umapa.items():
        porconj.setdefault(info["conjunto"], []).append(info)
    r = 5
    for _, m in sel_meta.iterrows():
        infos = porconj.get(m["nom_usina"])
        if not infos:
            continue
        for mes in meses:
            ws.cell(row=r, column=1, value=m["nom_usina"]).font = BODY
            ws.cell(row=r, column=2, value=m["id_ons"]).font = BODY
            c = ws.cell(row=r, column=3, value=datetime.strptime(mes + "-01", "%Y-%m-%d"))
            c.font, c.number_format = BODY, MES
            c = ws.cell(row=r, column=4,
                        value=f"=SUMIFS('Mês a mês'!$F$5:$F${max(mm_last,5)},"
                              f"'Mês a mês'!$B$5:$B${max(mm_last,5)},$B{r},"
                              f"'Mês a mês'!$C$5:$C${max(mm_last,5)},$C{r})")
            c.font, c.number_format = BODY, MWH
            soma = "+".join(_usina_sumifs(i, "mwh_gnr", mes)[1:] for i in infos)
            c = ws.cell(row=r, column=5, value="=" + soma)
            c.font, c.number_format = BODY, MWH
            c = ws.cell(row=r, column=6, value=f"=E{r}-D{r}")
            c.font, c.number_format = BODY, MWH
            c = ws.cell(row=r, column=7, value=f'=IFERROR(F{r}/D{r},"")')
            c.font, c.number_format = BODY, PCT
            c = ws.cell(row=r, column=8, value=len(infos))
            c.font, c.number_format = BODY, INT
            r += 1
    ws.auto_filter.ref = f"A4:H{max(r-1,4)}"


# --------------------------------------------------------------------------
def write_excel(path: Path, daily: pd.DataFrame, detail: pd.DataFrame,
                conjuntos: pd.DataFrame, selected: list[str], prices: pd.DataFrame,
                usinas: pd.DataFrame | None = None) -> None:
    solar = daily[daily["fonte"] == "Solar"].copy()
    if solar.empty:
        raise SystemExit("nenhum dado fotovoltaico para escrever no Excel")
    for c in ("min_REL", "min_CNF", "min_ENE"):
        if c not in solar.columns:
            solar[c] = 0.0

    ids = set(solar["id_ons"])
    nomes = set(solar["nom_usina"])
    sel_ids = [s for s in selected if s in ids]
    sel_ids += solar[solar["nom_usina"].isin(selected)]["id_ons"].unique().tolist()
    sel_ids = list(dict.fromkeys(sel_ids))
    missing = [s for s in selected if s not in ids and s not in nomes]
    sel_meta = (solar[solar["id_ons"].isin(sel_ids)]
                .groupby("id_ons", as_index=False)
                .agg(nom_usina=("nom_usina", "last"), id_estado=("id_estado", "last"),
                     id_subsistema=("id_subsistema", "last")))
    solar_sel = solar[solar["id_ons"].isin(sel_ids)] if sel_ids else solar
    meses = sorted(solar_sel["data"].dt.strftime("%Y-%m").unique())

    wb = Workbook()
    # Arial em tudo pelo estilo Normal: as abas de histórico têm centenas de
    # milhares de linhas e estilizar célula a célula multiplicaria o arquivo.
    wb._named_styles["Normal"].font = Font(name=FONT, size=10)

    ws_readme = wb.active
    ws_meus = _sheet(wb, "Meus conjuntos")
    ws_mm = _sheet(wb, "Mês a mês")
    ws_cd = _sheet(wb, "Condicionantes")
    ws_pr = _sheet(wb, "Perdas R$")
    ws_mn = _sheet(wb, "Minutos")
    ws_uc = _sheet(wb, "Conjunto x usinas")
    ws_ur = _sheet(wb, "Usinas — resumo")

    conj_sheets: dict[str, str] = {}
    taken: set[str] = {"Leia-me", "Meus conjuntos", "Mês a mês", "Condicionantes",
                       "Perdas R$", "Minutos", "Conjunto x usinas", "Usinas — resumo"}
    for _, m in sel_meta.iterrows():
        nm = _safe_name(m["nom_usina"], taken)
        conj_sheets[m["id_ons"]] = nm
        _sheet(wb, nm)

    mapa = _write_hist(wb, detail)
    umapa = _usina_tabs(wb, usinas, taken)

    ws_resumo = _sheet(wb, "Resumo")
    ws_mensal = _sheet(wb, "Mensal")
    ws_dia = _sheet(wb, "Diário")
    ws_pld = _sheet(wb, "PLD horário")

    _readme(ws_readme, sel_meta, missing, solar, prices)

    tem_hist = bool(mapa)

    # ---- Mês a mês ---------------------------------------------------------
    heads_mm = ["Conjunto", "ID ONS", "Mês", "Geração (MWh)", "Geração esperada (MWh)",
                "Curtailment total (MWh)", "Curt. REL (MWh)", "Curt. CNF (MWh)",
                "Curt. ENE (MWh)", "Curt. PAR (MWh)"]
    fmt_mm = {h: MWH for h in heads_mm if "MWh" in h}

    def cel_mm(idons, mes, r):
        if not tem_hist:
            return [0] * 7
        return [_sumifs(mapa, mes, "mwh_gerado", idons),
                _sumifs(mapa, mes, "mwh_referencia", idons),
                _sumifs(mapa, mes, "mwh_gnr", idons),
                *[_sumifs(mapa, mes, "mwh_gnr", idons, cod) for cod in RAZAO_ORDER]]

    mm_last = _consolidada(
        ws_mm, "Geração, geração esperada e curtailment — mês a mês",
        "Cada número é um SUMIFS nas abas de histórico. Os percentuais dividem pela "
        "geração esperada do próprio mês.",
        sel_meta, meses, heads_mm, cel_mm, fmt_mm,
        {"Conjunto": 30, "ID ONS": 14, "Mês": 11, "Geração (MWh)": 14,
         "Geração esperada (MWh)": 16, "Curtailment total (MWh)": 16})
    for off, (h, num) in enumerate([("% curt. total", "F"), ("% curt. REL", "G"),
                                    ("% curt. CNF", "H"), ("% curt. ENE", "I"),
                                    ("% curt. PAR", "J")]):
        _pct_col(ws_mm, 11 + off, h, 5, mm_last, num, "E")

    # ---- Condicionantes ----------------------------------------------------
    heads_cd = ["Conjunto", "ID ONS", "Mês", "Curtailment total (MWh)",
                "Com condicionantes — vigente (MWh)",
                "Metodologia nova em todos os meses (MWh)",
                "Regra antiga em todos os meses (MWh)", "Patamares com restrição",
                "Patamares fora da tolerância", "Patamares sem Geração Limitada"]
    fmt_cd = {h: MWH for h in heads_cd if "MWh" in h}
    fmt_cd.update({h: INT for h in heads_cd if h.startswith("Patamares")})

    def cel_cd(idons, mes, r):
        if not tem_hist:
            return [0] * 7
        return [_sumifs(mapa, mes, "mwh_gnr", idons),
                _sumifs(mapa, mes, "mwh_gnr_vigente", idons),
                _sumifs(mapa, mes, "mwh_gnr_nova", idons),
                _sumifs(mapa, mes, "mwh_gnr_antiga", idons),
                _countifs(mapa, mes, idons, _crit_col(mapa, mes, "cod_razaorestricao", '"<>"')),
                _countifs(mapa, mes, idons, _crit_col(mapa, mes, "tol_atendida", "FALSE")),
                _countifs(mapa, mes, idons,
                          _crit_col(mapa, mes, "cod_razaorestricao", '"<>"')
                          + _crit_col(mapa, mes, "val_geracaolimitada", '""'))]

    cd_last = _consolidada(
        ws_cd, "Curtailment publicado × recalculado com as condicionantes",
        "RO-AO.BR.13 Rev.09, itens 4.13 e 5.2.2.10. Reconstrução a partir dos dados "
        "publicados — não é a apuração oficial. Ver Leia-me.",
        sel_meta, meses, heads_cd, cel_cd, fmt_cd,
        {"Conjunto": 30, "ID ONS": 14, "Mês": 11, "Curtailment total (MWh)": 15,
         "Com condicionantes — vigente (MWh)": 17,
         "Metodologia nova em todos os meses (MWh)": 18,
         "Regra antiga em todos os meses (MWh)": 18,
         "Patamares com restrição": 12, "Patamares fora da tolerância": 13,
         "Patamares sem Geração Limitada": 13})
    for off, (h, expr) in enumerate([("Diferença: total − vigente (MWh)", "=D{r}-E{r}"),
                                     ("Diferença: nova − vigente (MWh)", "=F{r}-E{r}"),
                                     ("% retido pela condicionante", '=IFERROR(E{r}/D{r},"")')]):
        col = 11 + off
        c = ws_cd.cell(row=4, column=col, value=h)
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws_cd.column_dimensions[get_column_letter(col)].width = 16
        for r in range(5, cd_last + 1):
            cell = ws_cd.cell(row=r, column=col, value=expr.format(r=r))
            cell.font = BODY
            cell.number_format = PCT if "%" in h else MWH

    # ---- Perdas R$ ---------------------------------------------------------
    heads_pr = ["Conjunto", "ID ONS", "Mês", "Perda total (R$)", "Perda REL (R$)",
                "Perda CNF (R$)", "Perda ENE (R$)", "Perda PAR (R$)"]

    def cel_pr(idons, mes, r):
        if not tem_hist:
            return [0] * 5
        return [_sumifs(mapa, mes, "perda_rs", idons),
                *[_sumifs(mapa, mes, "perda_rs", idons, cod) for cod in RAZAO_ORDER]]

    pr_last = _consolidada(
        ws_pr, "Perda financeira da energia não gerada",
        "Curtailment × PLD horário do submercado. Custo de oportunidade, não "
        "ressarcimento — ver Leia-me.",
        sel_meta, meses, heads_pr, cel_pr, {h: BRL for h in heads_pr if "R$" in h},
        {"Conjunto": 30, "ID ONS": 14, "Mês": 11, "Perda total (R$)": 16})

    # ---- Minutos -----------------------------------------------------------
    heads_mn = ["Conjunto", "ID ONS", "Mês", "Minutos REL", "Minutos CNF", "Minutos ENE",
                "Minutos em restrição (total)", "Patamares com restrição"]

    def cel_mn(idons, mes, r):
        if not tem_hist:
            return [0] * 5
        return [*[_sumifs(mapa, mes, "minutos", idons, cod) for cod in ("REL", "CNF", "ENE")],
                _sumifs(mapa, mes, "minutos", idons),
                _countifs(mapa, mes, idons,
                          _crit_col(mapa, mes, "cod_razaorestricao", '"<>"'))]

    mn_last = _consolidada(
        ws_mn, "Minutos em restrição por razão",
        "Antes de jan/2026 o ONS não publicava num_minutos_*; nesses meses cada patamar "
        "restrito conta 30 minutos. Ver Leia-me.",
        sel_meta, meses, heads_mn, cel_mn, {h: INT for h in heads_mn[3:]},
        {"Conjunto": 30, "ID ONS": 14, "Mês": 11, "Minutos em restrição (total)": 16,
         "Patamares com restrição": 13})
    col = len(heads_mn) + 1
    c = ws_mn.cell(row=4, column=col, value="Horas em restrição")
    c.font, c.fill = HDR_FONT, HDR_FILL
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws_mn.column_dimensions[get_column_letter(col)].width = 14
    for r in range(5, mn_last + 1):
        cell = ws_mn.cell(row=r, column=col, value=f"=G{r}/60")
        cell.font, cell.number_format = BODY, "#,##0.0"

    # ---- usinas ------------------------------------------------------------
    if umapa:
        _usina_resumo(ws_ur, umapa, meses, mapa)
        _usina_comparacao(ws_uc, umapa, sel_meta, meses, mm_last)
    else:
        for ws, txt in ((ws_ur, "Sem detalhamento por usina."),
                        (ws_uc, "Sem detalhamento por usina.")):
            ws.cell(row=1, column=1, value=txt).font = BOLD
            ws.cell(row=2, column=1,
                    value="Liste os conjuntos em 'usinas_detalhe' no config/conjuntos.yml "
                          "e rode o workflow com full = true.").font = NOTE

    # ---- abas por conjunto -------------------------------------------------
    meses_dt = [datetime.strptime(m + "-01", "%Y-%m-%d") for m in meses]
    for _, m in sel_meta.iterrows():
        _conjunto_sheet(wb[conj_sheets[m["id_ons"]]], m, meses_dt, conjuntos,
                        mm_last, cd_last, pr_last, mn_last)

    _meus(ws_meus, sel_meta, conjuntos, conj_sheets, mm_last, pr_last)
    _abas_gerais(ws_dia, ws_resumo, ws_mensal, None, ws_pld,
                 solar, detail, conjuntos, sel_ids, prices)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"excel written: {path} ({path.stat().st_size/1e6:.1f} MB) — "
          f"{len(sel_ids)} conjuntos, {len(umapa)} usinas, {len(mapa)} meses no histórico")

# --------------------------------------------------------------------------
def _conjunto_sheet(ws, meta, meses, conjuntos, mm_last, cd_last, pr_last, mn_last) -> None:
    """As quatro tabelas empilhadas, todas por SUMIFS das abas consolidadas."""
    idons = meta["id_ons"]
    cj = conjuntos.set_index("id_ons")
    cap = cj["capacidade_mw"].get(idons, float("nan"))
    ws.cell(row=1, column=1, value=meta["nom_usina"]).font = TITLE
    ws.cell(row=2, column=1,
            value=f"{idons} · {meta['id_estado']} · "
                  f"{SUBSISTEMA_NOME.get(meta['id_subsistema'], meta['id_subsistema'])}"
                  + (f" · {cap:,.1f} MW".replace(",", ".") if pd.notna(cap) else "")).font = NOTE
    ws.column_dimensions["A"].width = 12
    for j in range(2, 14):
        ws.column_dimensions[get_column_letter(j)].width = 15

    def block(title, start, headers, specs, fmts, src_last):
        ws.cell(row=start, column=1, value=title).font = SECTION
        _head(ws, headers, start + 1)
        ws.column_dimensions["A"].width = 12
        r = start + 2
        for mes in meses:
            c = ws.cell(row=r, column=1, value=pd.Timestamp(mes).to_pydatetime())
            c.font, c.number_format = BODY, MES
            for j, (sheet, col, fmt) in enumerate(specs, start=2):
                if sheet is None:  # formula relative to this row
                    c = ws.cell(row=r, column=j, value=col.format(r=r))
                else:
                    rng = f"'{sheet}'!${col}$5:${col}${max(src_last[sheet], 5)}"
                    idr = f"'{sheet}'!$B$5:$B${max(src_last[sheet], 5)}"
                    mr = f"'{sheet}'!$C$5:$C${max(src_last[sheet], 5)}"
                    c = ws.cell(row=r, column=j,
                                value=f'=SUMIFS({rng},{idr},"{idons}",{mr},$A{r})')
                c.font, c.number_format = BODY, fmt
            r += 1
        # total row
        ws.cell(row=r, column=1, value="TOTAL").font = BOLD
        for j in range(2, 2 + len(specs)):
            L = get_column_letter(j)
            fmt = fmts[j - 2]
            if fmt == PCT:
                continue
            c = ws.cell(row=r, column=j, value=f"=SUM({L}{start+2}:{L}{r-1})")
            c.font, c.number_format = BOLD, fmt
        return r + 2

    src = {"Mês a mês": mm_last, "Perdas R$": pr_last, "Minutos": mn_last}
    if cd_last:
        src["Condicionantes"] = cd_last

    row = 4
    # 1 — geração e curtailment
    heads = ["Geração (MWh)", "Geração esperada (MWh)", "Curtailment total (MWh)",
             "% curt. total", "Curt. REL (MWh)", "% REL", "Curt. CNF (MWh)", "% CNF",
             "Curt. ENE (MWh)", "% ENE"]
    specs = [("Mês a mês", "D", MWH), ("Mês a mês", "E", MWH), ("Mês a mês", "F", MWH),
             (None, '=IFERROR(D{r}/C{r},"")', PCT),
             ("Mês a mês", "G", MWH), (None, '=IFERROR(F{r}/C{r},"")', PCT),
             ("Mês a mês", "H", MWH), (None, '=IFERROR(H{r}/C{r},"")', PCT),
             ("Mês a mês", "I", MWH), (None, '=IFERROR(J{r}/C{r},"")', PCT)]
    fmts = [s[2] for s in specs]
    row = block("1. Geração e curtailment, mês a mês", row, heads, specs, fmts, src)

    # 2 — condicionantes
    if cd_last:
        heads = ["Curtailment total (MWh)", "Com condicionantes — vigente (MWh)",
                 "Metodologia nova em todos os meses (MWh)",
                 "Regra antiga em todos os meses (MWh)", "Diferença: total − vigente (MWh)",
                 "Diferença: nova − vigente (MWh)", "Patamares fora da tolerância"]
        specs = [("Condicionantes", "D", MWH), ("Condicionantes", "E", MWH),
                 ("Condicionantes", "F", MWH), ("Condicionantes", "G", MWH),
                 (None, "=B{r}-C{r}", MWH), (None, "=D{r}-C{r}", MWH),
                 ("Condicionantes", "I", INT)]
        fmts = [s[2] for s in specs]
        row = block("2. Curtailment com as condicionantes (RO-AO.BR.13)", row, heads,
                    specs, fmts, src)

    # 3 — perdas R$
    heads = ["Perda total (R$)", "Perda REL (R$)", "Perda CNF (R$)", "Perda ENE (R$)"]
    specs = [("Perdas R$", "D", BRL), ("Perdas R$", "E", BRL),
             ("Perdas R$", "F", BRL), ("Perdas R$", "G", BRL)]
    fmts = [s[2] for s in specs]
    row = block("3. Perda financeira (curtailment × PLD)", row, heads, specs, fmts, src)

    # 4 — minutos
    heads = ["Minutos REL", "Minutos CNF", "Minutos ENE", "Minutos totais", "Horas totais"]
    specs = [("Minutos", "D", INT), ("Minutos", "E", INT), ("Minutos", "F", INT),
             ("Minutos", "G", INT), (None, "=E{r}/60", "#,##0.0")]
    fmts = [s[2] for s in specs]
    block("4. Minutos em restrição por razão", row, heads, specs, fmts, src)
    ws.freeze_panes = "B6"


# --------------------------------------------------------------------------
def _meus(ws, sel_meta, conjuntos, conj_sheets, mm_last, pr_last) -> None:
    ws.cell(row=1, column=1, value="Meus conjuntos — total do período").font = TITLE
    ws.cell(row=2, column=1,
            value="SUMIFS sobre 'Mês a mês' e 'Perdas R$'. Clique no nome para abrir a aba "
                  "do conjunto.").font = NOTE
    heads = ["Conjunto", "ID ONS", "UF", "Submercado", "Capacidade (MW)", "Geração (MWh)",
             "Geração esperada (MWh)", "Curtailment (MWh)", "% do potencial",
             "Curt. REL (MWh)", "Curt. CNF (MWh)", "Curt. ENE (MWh)", "Perda total (R$)",
             "R$/MWh médio", "Latitude", "Longitude"]
    _head(ws, heads, 4, {"Conjunto": 30, "ID ONS": 14, "Geração esperada (MWh)": 16,
                         "Curtailment (MWh)": 15, "Perda total (R$)": 16})
    ws.column_dimensions["A"].width = 30
    ws.freeze_panes = "A5"
    cj = conjuntos.set_index("id_ons")
    r = 5
    for _, m in sel_meta.iterrows():
        i = m["id_ons"]
        sheet = conj_sheets.get(i)
        c = ws.cell(row=r, column=1, value=m["nom_usina"])
        c.font = BODY
        if sheet:
            c.hyperlink = f"#'{sheet}'!A1"
            c.font = Font(name=FONT, size=10, color="0563C1", underline="single")
        ws.cell(row=r, column=2, value=i).font = BODY
        ws.cell(row=r, column=3, value=m["id_estado"]).font = BODY
        ws.cell(row=r, column=4,
                value=SUBSISTEMA_NOME.get(m["id_subsistema"], m["id_subsistema"])).font = BODY
        cap = cj["capacidade_mw"].get(i, float("nan"))
        c = ws.cell(row=r, column=5, value=None if pd.isna(cap) else round(float(cap), 1))
        c.font, c.number_format = BLUE, "#,##0.0"
        mmr = lambda L: (f"=SUMIFS('Mês a mês'!${L}$5:${L}${max(mm_last,5)},"
                         f"'Mês a mês'!$B$5:$B${max(mm_last,5)},\"{i}\")")
        for j, L in ((6, "D"), (7, "E"), (8, "F"), (10, "G"), (11, "H"), (12, "I")):
            c = ws.cell(row=r, column=j, value=mmr(L))
            c.font, c.number_format = BODY, MWH
        c = ws.cell(row=r, column=9, value=f'=IFERROR(H{r}/G{r},"")')
        c.font, c.number_format = BODY, PCT
        c = ws.cell(row=r, column=13,
                    value=f"=SUMIFS('Perdas R$'!$D$5:$D${max(pr_last,5)},"
                          f"'Perdas R$'!$B$5:$B${max(pr_last,5)},\"{i}\")")
        c.font, c.number_format = BODY, BRL
        c = ws.cell(row=r, column=14, value=f'=IFERROR(M{r}/H{r},"")')
        c.font, c.number_format = BODY, BRL
        for k, key in enumerate(("lat", "lon")):
            v = cj[key].get(i, float("nan"))
            c = ws.cell(row=r, column=15 + k, value=None if pd.isna(v) else float(v))
            c.font, c.number_format = BLUE, "0.000000"
        r += 1
    ws.cell(row=r, column=1, value="TOTAL").font = BOLD
    for col in ("F", "G", "H", "J", "K", "L", "M"):
        j = "ABCDEFGHIJKLM".index(col) + 1
        c = ws.cell(row=r, column=j, value=f"=SUM({col}5:{col}{r-1})")
        c.font = BOLD
        c.number_format = BRL if col == "M" else MWH


# --------------------------------------------------------------------------
def _abas_gerais(ws_dia, ws_resumo, ws_mensal, _unused, ws_pld,
                 solar, detail, conjuntos, sel_ids, prices) -> None:
    dia = solar.pivot_table(
        index=["id_ons", "nom_usina", "id_subsistema", "id_estado", "data", "pld_fonte"],
        columns="cod_razaorestricao", values="mwh_gnr", aggfunc="sum", fill_value=0.0,
    ).reset_index()
    for code in RAZAO_ORDER:
        if code not in dia.columns:
            dia[code] = 0.0
    money = solar.groupby(["id_ons", "data"], as_index=False)["rs_gnr"].sum()
    minutes = solar.groupby(["id_ons", "data"], as_index=False)["minutos"].sum()
    dia = dia.merge(money, on=["id_ons", "data"], how="left") \
             .merge(minutes, on=["id_ons", "data"], how="left")
    dia["MWh total"] = dia[RAZAO_ORDER].sum(axis=1)
    dia = dia.rename(columns={"id_ons": "ID ONS", "nom_usina": "Conjunto",
                              "id_subsistema": "Submercado", "id_estado": "UF",
                              "data": "Data", "pld_fonte": "PLD (fonte)",
                              "rs_gnr": "Valor (R$)", "minutos": "Minutos em restrição"})
    dia = dia[DIA_HEADERS].sort_values(["Conjunto", "Data"])

    ws_dia.cell(row=1, column=1,
                value="Energia não gerada por conjunto e por dia — todos os conjuntos FV"
                ).font = TITLE
    fmt = {c: MWH for c in [*RAZAO_ORDER, "MWh total"]}
    fmt.update({"Valor (R$)": BRL, "Data": DATE, "Minutos em restrição": INT})
    dia_last = _table(ws_dia, dia, 4, fmt,
                      {"Conjunto": 30, "ID ONS": 14, "Data": 12, "MWh total": 14,
                       "Valor (R$)": 15, "Minutos em restrição": 15, "PLD (fonte)": 14})

    # Resumo (todos)
    ws_resumo.cell(row=1, column=1,
                   value="Resumo por conjunto — todos os conjuntos fotovoltaicos").font = TITLE
    heads = ["Conjunto", "ID ONS", "Submercado", "UF", "Capacidade (MW)",
             *[f"{c} (MWh)" for c in RAZAO_ORDER], "MWh não gerados", "Valor (R$)",
             "R$/MWh médio", "Meu?"]
    _head(ws_resumo, heads, 4, {"Conjunto": 32, "ID ONS": 14, "MWh não gerados": 16,
                                "Valor (R$)": 15, "Meu?": 8})
    ws_resumo.column_dimensions["A"].width = 32
    ws_resumo.freeze_panes = "A5"
    cj = conjuntos.set_index("id_ons")
    rng_id = f"'Diário'!${DIA_ID}$5:${DIA_ID}${max(dia_last,5)}"
    nomes = dia[["Conjunto", "ID ONS", "Submercado", "UF"]].drop_duplicates("ID ONS")
    r = 5
    for conj, idons, sub, uf in nomes.itertuples(index=False, name=None):
        ws_resumo.cell(row=r, column=1, value=conj).font = BODY
        ws_resumo.cell(row=r, column=2, value=idons).font = BODY
        ws_resumo.cell(row=r, column=3, value=SUBSISTEMA_NOME.get(sub, sub)).font = BODY
        ws_resumo.cell(row=r, column=4, value=uf).font = BODY
        cap = cj["capacidade_mw"].get(idons, float("nan"))
        c = ws_resumo.cell(row=r, column=5, value=None if pd.isna(cap) else round(float(cap), 1))
        c.font, c.number_format = BLUE, "#,##0.0"
        for k, code in enumerate(RAZAO_ORDER):
            L = DIA_RZ[code]
            c = ws_resumo.cell(row=r, column=6 + k,
                               value=f"=SUMIFS('Diário'!${L}$5:${L}${max(dia_last,5)},{rng_id},$B{r})")
            c.font, c.number_format = BODY, MWH
        c = ws_resumo.cell(row=r, column=10, value=f"=SUM(F{r}:I{r})")
        c.font, c.number_format = BODY, MWH
        c = ws_resumo.cell(row=r, column=11,
                           value=f"=SUMIFS('Diário'!${DIA_RS}$5:${DIA_RS}${max(dia_last,5)},{rng_id},$B{r})")
        c.font, c.number_format = BODY, BRL
        c = ws_resumo.cell(row=r, column=12, value=f'=IFERROR(K{r}/J{r},"")')
        c.font, c.number_format = BODY, BRL
        c = ws_resumo.cell(row=r, column=13, value="sim" if idons in sel_ids else "")
        c.font = BOLD if idons in sel_ids else BODY
        if idons in sel_ids:
            for col in range(1, 14):
                ws_resumo.cell(row=r, column=col).fill = SUB_FILL
        r += 1
    ws_resumo.auto_filter.ref = f"A4:M{r-1}"

    # Mensal (todos)
    mensal = solar.copy()
    mensal["Mês"] = mensal["data"].dt.to_period("M").dt.to_timestamp()
    mensal = (mensal.groupby(["nom_usina", "id_ons", "Mês"], as_index=False)
              .agg(**{"MWh não gerados": ("mwh_gnr", "sum"), "Valor (R$)": ("rs_gnr", "sum"),
                      "Geração (MWh)": ("mwh_gerado", "sum"),
                      "Geração esperada (MWh)": ("mwh_referencia", "sum"),
                      "Minutos em restrição": ("minutos", "sum")})
              .rename(columns={"nom_usina": "Conjunto", "id_ons": "ID ONS"})
              .sort_values(["Conjunto", "Mês"]))
    ws_mensal.cell(row=1, column=1, value="Totais mensais — todos os conjuntos FV").font = TITLE
    _table(ws_mensal, mensal, 3,
           {"Mês": MES, "MWh não gerados": MWH, "Valor (R$)": BRL, "Geração (MWh)": MWH,
            "Geração esperada (MWh)": MWH, "Minutos em restrição": INT},
           {"Conjunto": 32, "ID ONS": 14})

    # PLD — uma linha por hora, os quatro submercados lado a lado
    if not prices.empty:
        p = prices[prices["hora"] >= PLD_INICIO].copy()
        wide = p.pivot_table(index="hora", columns="submercado", values="pld",
                             aggfunc="last")
        for s in ("N", "NE", "S", "SE"):
            if s not in wide.columns:
                wide[s] = pd.NA
        fonte = (p.sort_values("pld_fonte")
                 .groupby("hora")["pld_fonte"].last().rename("Fonte"))
        wide = wide[["N", "NE", "S", "SE"]].join(fonte).reset_index()
        wide = wide.rename(columns={
            "hora": "Data/hora", "N": "Norte (R$/MWh)", "NE": "Nordeste (R$/MWh)",
            "S": "Sul (R$/MWh)", "SE": "Sudeste/CO (R$/MWh)"})
        wide = wide.sort_values("Data/hora")

        ws_pld.cell(row=1, column=1,
                    value="PLD horário por submercado — CCEE Dados Abertos").font = TITLE
        ws_pld.cell(
            row=2, column=1,
            value=f"Uma linha por hora, de {PLD_INICIO:%m/%Y} em diante. Fonte 'horario' = "
                  "PLD horário liquidado (pld_horario_submercado); 'semanal_media' = média "
                  "semanal, usada enquanto o mês não fecha; 'indisponivel' = a CCEE ainda "
                  "não publicou nada para aquela hora.").font = NOTE
        _table(ws_pld, wide.head(EXCEL_ROW_LIMIT), 4,
               {"Data/hora": DTM, "Norte (R$/MWh)": BRL, "Nordeste (R$/MWh)": BRL,
                "Sul (R$/MWh)": BRL, "Sudeste/CO (R$/MWh)": BRL},
               {"Data/hora": 18, "Norte (R$/MWh)": 15, "Nordeste (R$/MWh)": 15,
                "Sul (R$/MWh)": 15, "Sudeste/CO (R$/MWh)": 17, "Fonte": 14})
        ws_pld.freeze_panes = "B5"
