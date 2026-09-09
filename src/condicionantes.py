"""Recálculo do curtailment com as condicionantes da RO-AO.BR.13.

O campo ``val_geracaonaorealizadaapurada`` que o ONS publica é a Geração Não
Realizada Apurada **sem** condicionante: GNRa = max(0, G_Ref − G_Verificada)
nos patamares com limitação.  As duas variantes abaixo reconstroem o que
sobraria depois de aplicar o teste de tolerância.

Teste de tolerância (item 4.13 da RO-AO.BR.13 Rev.09)
-----------------------------------------------------
    "Para efeito de apuração diária, será admitida uma tolerância de 5% ou
     5 MW, o que for menor, da 'Geração Verificada' em relação à
     'Geração Limitada'."

    tolerancia = min(0,05 × Geração Verificada ; 5 MW)
    desvio     = Geração Limitada − Geração Verificada
    atendido   = |desvio| <= tolerancia

A base dos 5% (Geração Verificada) foi a leitura escolhida pela usuária;
troque ``TOL_BASE`` para "limitada" ou "capacidade" se a interpretação mudar.

Regra vigente até jul/2025 ("antiga")
-------------------------------------
Se a tolerância não era atendida, o ONS entendia que a usina não estava
encostada no limite e **zerava** a GNR daquele patamar.

Regra vigente a partir de ago/2025 (item 5.2.2.10, "nova")
----------------------------------------------------------
    a) G_ref_Disp   = min(Geração de Referência ; Disponibilidade eletromecânica)
    b) atendido     → G_Ref_Final = G_ref_Disp
       não atendido → G_Ref_Final = G_ref_Disp − (Geração Limitada − Geração Verificada)
    c) G_Ref_Final < 0 → 0

    GNR = max(0, G_Ref_Final − Geração Verificada)

Em vez de zerar, a metodologia nova desconta o desvio — por isso os dois
números diferem, e é essa diferença que a aba "Condicionantes" mostra.

Aviso
-----
Isto é uma **reconstrução** a partir dos dados publicados, não a apuração
oficial. O ONS calcula a Geração de Referência por função de produtividade e
faz o rateio do item 5.2.2.11 por usina; aqui partimos da Geração de
Referência já publicada, que é exatamente onde o item 5.2.2.10 começa.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

TOL_MW = 5.0
TOL_PCT = 0.05
TOL_BASE = "verificada"  # "verificada" | "limitada" | "capacidade"

# A metodologia do item 5.2.2.10 passou a valer nesta competência.
VIGENCIA_NOVA = pd.Timestamp("2025-08-01")

HALF_HOUR = 0.5


def _tolerancia(df: pd.DataFrame, capacidade: pd.Series | None = None) -> pd.Series:
    if TOL_BASE == "limitada":
        base = df["val_geracaolimitada"]
    elif TOL_BASE == "capacidade":
        base = df["id_ons"].map(capacidade) if capacidade is not None else df["val_geracao"]
    else:
        base = df["val_geracao"]
    return np.minimum(TOL_PCT * base.fillna(0.0).abs(), TOL_MW)


def aplicar(detail: pd.DataFrame, capacidade: pd.Series | None = None) -> pd.DataFrame:
    """Acrescenta ao detalhe semi-horário as colunas de curtailment recalculado.

    Colunas criadas (todas em MWh):
      mwh_gnr              — o publicado pelo ONS (já vem do carregamento)
      mwh_gnr_antiga       — regra que zerava o patamar fora da tolerância
      mwh_gnr_nova         — item 5.2.2.10 aplicado a todos os meses
      mwh_gnr_vigente      — antiga até jul/2025, nova de ago/2025 em diante
      tol_atendida         — bool do teste de tolerância
      sem_limite           — patamar com razão declarada mas sem Geração Limitada
    """
    d = detail.copy()
    for c in ("val_geracao", "val_geracaolimitada", "val_disponibilidade",
              "val_geracaoreferencia"):
        if c not in d.columns:
            d[c] = np.nan
        d[c] = pd.to_numeric(d[c], errors="coerce")

    restrito = d["cod_razaorestricao"].fillna("").ne("")
    ger = d["val_geracao"].fillna(0.0)
    lim = d["val_geracaolimitada"]
    ref = d["val_geracaoreferencia"].fillna(0.0)
    disp = d["val_disponibilidade"]

    d["sem_limite"] = restrito & lim.isna()
    desvio = (lim.fillna(ger) - ger)          # sem limite ⇒ desvio 0 ⇒ tolerância atendida
    tol = _tolerancia(d, capacidade)
    d["tol_atendida"] = desvio.abs() <= tol

    # ---- regra antiga: fora da tolerância, zera o patamar -------------------
    gnr_antiga = np.where(d["tol_atendida"], np.maximum(0.0, ref - ger), 0.0)

    # ---- item 5.2.2.10 -----------------------------------------------------
    g_ref_disp = np.minimum(ref, disp.fillna(np.inf))
    g_ref_final = np.where(d["tol_atendida"], g_ref_disp, g_ref_disp - desvio)
    g_ref_final = np.maximum(0.0, g_ref_final)
    gnr_nova = np.maximum(0.0, g_ref_final - ger)

    for name, arr in (("mwh_gnr_antiga", gnr_antiga), ("mwh_gnr_nova", gnr_nova)):
        d[name] = np.where(restrito, arr, 0.0) * HALF_HOUR

    novo_vale = d["din_instante"] >= VIGENCIA_NOVA
    d["mwh_gnr_vigente"] = np.where(novo_vale, d["mwh_gnr_nova"], d["mwh_gnr_antiga"])
    d["g_ref_final_calc"] = g_ref_final
    return d


def mensal(detail: pd.DataFrame) -> pd.DataFrame:
    """Resumo mensal por conjunto das três variantes + diferença."""
    d = detail.copy()
    d["Mês"] = d["din_instante"].dt.to_period("M").dt.to_timestamp()
    g = (d.groupby(["nom_usina", "id_ons", "Mês"], as_index=False)
         .agg(**{
             "Curtailment total (MWh)": ("mwh_gnr", "sum"),
             "Com condicionantes — regra vigente (MWh)": ("mwh_gnr_vigente", "sum"),
             "Com condicionantes — metodologia nova em todos os meses (MWh)":
                 ("mwh_gnr_nova", "sum"),
             "Com condicionantes — regra antiga em todos os meses (MWh)":
                 ("mwh_gnr_antiga", "sum"),
             "Patamares com restrição": ("cod_razaorestricao", lambda s: (s.fillna("") != "").sum()),
             "Patamares fora da tolerância": ("tol_atendida", lambda s: (~s).sum()),
             "Patamares sem Geração Limitada": ("sem_limite", "sum"),
         }))
    g = g.rename(columns={"nom_usina": "Conjunto", "id_ons": "ID ONS"})
    return g.sort_values(["Conjunto", "Mês"])
