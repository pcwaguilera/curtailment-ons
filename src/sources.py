"""Registry of the ONS / CCEE open-data sources used by this project.

All ONS constrained-off datasets live in one public S3 bucket, one CSV per
month, semicolon-separated, UTF-8, dot decimal separator.
"""

ONS_S3 = "https://ons-aws-prod-opendata.s3.amazonaws.com"

# key -> (s3 prefix, file stem)
# The monthly file name is  f"{stem}_{YYYY}_{MM}.csv"
ONS_DATASETS = {
    # ---- solar (fotovoltaica) -------------------------------------------
    "solar_tm": (
        "dataset/restricao_coff_fotovoltaica_tm",
        "RESTRICAO_COFF_FOTOVOLTAICA",
    ),
    "solar_detail": (
        "dataset/restricao_coff_fotovoltaica_detail_tm",
        "RESTRICAO_COFF_FOTOVOLTAICA_DETAIL",
    ),
    "solar_intra": (
        "dataset/restricao_coff_fotovoltaica_intrasemihora",
        "COFF_USI_FOTOVOLTAICA_INTRASEMIHORA",
    ),
    # ---- wind (eolica) ---------------------------------------------------
    "wind_tm": (
        "dataset/restricao_coff_eolica_tm",
        "RESTRICAO_COFF_EOLICA",
    ),
    "wind_detail": (
        "dataset/restricao_coff_eolica_detail_tm",
        "RESTRICAO_COFF_EOLICA_DETAIL",
    ),
    "wind_intra": (
        "dataset/restricao_coff_eolica_intrasemihora",
        "COFF_USI_EOLICAS_INTRASEMIHORA",
    ),
    # ---- capacity factor: the only ONS set carrying lat/lon --------------
    "fatorcap": (
        "dataset/fator_capacidade_2_di",
        "FATOR_CAPACIDADE-2",
    ),
}

# Human labels for the restriction-reason codes (cod_razaorestricao).
RAZAO = {
    "REL": "Indisponibilidade externa (elétrica)",
    "CNF": "Confiabilidade",
    "ENE": "Razão energética",
    "PAR": "Restrição do parecer de acesso",
}
RAZAO_ORDER = ["REL", "CNF", "ENE", "PAR"]

# Origem da restrição
ORIGEM = {"LOC": "Local", "SIS": "Sistêmica"}

# ---- CCEE ---------------------------------------------------------------
CCEE_CKAN = "https://dadosabertos.ccee.org.br"

# Hourly PLD per submercado.  `pld_horario_submercado` is the fresher of the
# two hourly series and is the one we use; `pld_horario` is kept as a
# fallback for older years.
CCEE_PLD_PACKAGES = ["pld_horario_submercado", "pld_horario"]
# Weekly average PLD - used to fill the gap between the last closed hourly
# month and today (CCEE only publishes the hourly file after monthly close).
CCEE_PLD_WEEKLY = "pld_media_semanal"

SUBMERCADO_MAP = {
    "NORTE": "N",
    "NORDESTE": "NE",
    "SUL": "S",
    "SUDESTE": "SE",
    "SUDESTE/CENTRO-OESTE": "SE",
    "N": "N",
    "NE": "NE",
    "S": "S",
    "SE": "SE",
}

SUBSISTEMA_NOME = {
    "N": "Norte",
    "NE": "Nordeste",
    "S": "Sul",
    "SE": "Sudeste/Centro-Oeste",
}
