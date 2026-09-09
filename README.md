# Curtailment no SIN — ONS + PLD

Painel público e planilha com a energia eólica e solar **cortada por ordem do ONS**
(constrained-off), atualizados sozinhos duas vezes por dia pelo GitHub Actions e
publicados no GitHub Pages.

- **Painel:** `https://<seu-usuario>.github.io/<repo>/` — mapa dos conjuntos, energia
  não gerada por razão de restrição, série diária e tabela filtrável (com download em CSV).
- **Excel:** `curtailment_solar.xlsx`, publicado junto com o painel
  (`https://<usuario>.github.io/<repo>/curtailment_solar.xlsx`). Fotovoltaica.

  A cadeia de números tem um sentido só: as **abas de histórico** (2024, 2025 S1,
  2025 S2, 2026) guardam os dados crus do ONS patamar a patamar; as **tabelas
  consolidadas** somam delas por SUMIFS; as **abas por conjunto** somam das
  consolidadas. Nada é digitado duas vezes e qualquer número é rastreável até o
  patamar que o gerou.

  As quatro tabelas mês a mês: geração e curtailment por razão com os percentuais ·
  curtailment com as condicionantes da RO-AO.BR.13 · perda em R$ por razão ·
  minutos em restrição. Mais o detalhamento **por usina** (uma aba de histórico
  diário por usina, um resumo mensal e a comparação conjunto × soma das usinas)
  para os conjuntos listados em `usinas_detalhe`.

  > A planilha **não é commitada** — ela tem dezenas de MB e é regerada a cada
  > execução, então o histórico do Git cresceria o tamanho dela por dia. Ela vai
  > direto para o GitHub Pages.

## De onde vêm os dados

| O quê | Fonte | Cobertura |
|---|---|---|
| Constrained-off fotovoltaica (semi-horário, por conjunto) | [ONS `restricao_coff_fotovoltaica`](https://dados.ons.org.br/dataset/restricao_coff_fotovoltaica) | abr/2024 → hoje |
| Constrained-off fotovoltaica, por usina | [ONS `restricao_coff_fotovoltaica_detail`](https://dados.ons.org.br/dataset/restricao_coff_fotovoltaica_detail) | abr/2024 → hoje |
| Constrained-off fotovoltaica intra-semi-hora | [ONS `coff_fotovoltaica_intrasemihora`](https://dados.ons.org.br/dataset/coff_fotovoltaica_intrasemihora) | jan/2026 → hoje |
| Constrained-off eólica (semi-horário, por conjunto) | [ONS `restricao_coff_eolica_usi`](https://dados.ons.org.br/dataset/restricao_coff_eolica_usi) | out/2021 → hoje |
| Constrained-off eólica, por usina | [ONS `restricao_coff_eolica_detail`](https://dados.ons.org.br/dataset/restricao_coff_eolica_detail) | out/2021 → hoje |
| Constrained-off eólica intra-semi-hora | [ONS `coff_eolica_usi_intrasemihora`](https://dados.ons.org.br/dataset/coff_eolica_usi_intrasemihora) | jan/2026 → hoje |
| Latitude/longitude e capacidade instalada dos conjuntos | [ONS `fator-capacidade-2`](https://dados.ons.org.br/dataset/fator-capacidade-2) | 2021 → hoje |
| PLD horário por submercado | [CCEE `pld_horario_submercado`](https://dadosabertos.ccee.org.br/dataset/pld_horario_submercado) | 2023 → último mês fechado |
| PLD médio semanal (preenche o mês em aberto) | [CCEE `pld_media_semanal`](https://dadosabertos.ccee.org.br/dataset/pld_media_semanal) | 2023 → semana passada |

Tudo sob licença CC-BY. O ONS avisa que os dados passam por consistência recorrente
e **podem mudar depois de publicados** — por isso o pipeline recarrega qualquer mês
cujo arquivo tenha sido republicado.

## Como o valor em R$ é calculado

```
energia não gerada (MWh) = val_geracaonaorealizadaapurada (MWmed) × 0,5 h
valor (R$)               = energia não gerada × PLD do submercado na hora do evento
```

Só contam os patamares em que o ONS declarou razão de restrição
(`cod_razaorestricao`): REL (indisponibilidade externa), CNF (confiabilidade),
ENE (razão energética) ou PAR (parecer de acesso).

> **Isto é custo de oportunidade, não ressarcimento.** O valor mostra quanto a energia
> perdida valeria no mercado de curto prazo. O ressarcimento de constrained-off que a
> CCEE efetivamente liquida segue regras próprias (RO-AO.BR.13 e as Regras de
> Comercialização) e em geral remunera apenas parte das restrições — tipicamente as
> por indisponibilidade externa. Para valor devido, use a apuração da CCEE.

## O curtailment antes de 2026 é reconstruído

O ONS só passou a publicar o campo `val_geracaonaorealizadaapurada` (a GNR apurada)
**a partir de janeiro de 2026** — o mesmo vale para os `num_minutos_*`. Antes disso os
arquivos trazem apenas geração verificada, geração limitada, disponibilidade, geração
de referência e a razão da restrição.

Para o histórico anterior o pipeline reproduz a definição do próprio ONS:

```
GNRa = max(0, Geração de Referência − Geração Verificada)   # só nos patamares
                                                            # com razão declarada
```

Cada linha carrega a coluna **`gnr_origem`** (`publicada` ou `calculada`) e cada mês,
nas abas de resumo, herda a origem predominante — dá para separar na hora de comparar
com a apuração oficial. Os minutos anteriores a 2026 contam 30 min por patamar restrito
e vêm marcados como estimados (`minutos_origem`).

Outras colunas que aparecem só depois: `dsc_restricao` de jan/2025;
`nom_agenteoperador`, `nom_pontoconexao` e `id_pontoconexao` de jan/2026.

## O detalhamento por usina

O dataset `restricao_coff_fotovoltaica_detail` tem outro schema: traz irradiância,
geração estimada e geração verificada **por usina**, mas não traz razão de restrição,
geração limitada nem GNRa. O `flg_geracaorestrita` e o `id_ons_conjuntousina` só
existem de 2026 em diante.

Então, por usina:

```
curtailment = max(0, Geração Estimada − Geração Verificada)   # nos patamares restritos
```

e a razão da restrição vem do arquivo de conjunto, casada por (conjunto, data/hora) —
pelo `id_ons_conjuntousina` quando existe, pelo nome do conjunto quando não.

As duas pontas são apuradas de formas diferentes, então **alguma divergência entre o
curtailment do conjunto e a soma das usinas é esperada por construção**. A aba
"Conjunto x usinas" existe justamente para medir essa diferença mês a mês.

## As condicionantes da RO-AO.BR.13

O `val_geracaonaorealizadaapurada` publicado pelo ONS é a GNR **sem** condicionante.
A aba "Condicionantes" reconstrói o que sobraria depois do teste de tolerância do
item 4.13, nas duas versões, a partir dos campos semi-horários de geração verificada,
geração limitada, geração de referência e disponibilidade eletromecânica.

```
tolerancia = min(5% × Geração Verificada ; 5 MW)      # item 4.13
desvio     = Geração Limitada − Geração Verificada
atendido   = |desvio| <= tolerancia
```

- **Regra antiga (até jul/2025):** tolerância não atendida ⇒ o patamar deixava de contar,
  curtailment = 0.
- **Metodologia nova (item 5.2.2.10, de ago/2025 em diante):**
  `G_ref_Disp = min(Ger. Referência ; Disponibilidade)`; atendido ⇒
  `G_Ref_Final = G_ref_Disp`, não atendido ⇒ `G_Ref_Final = G_ref_Disp − desvio`;
  negativo vira zero; `curtailment = max(0, G_Ref_Final − Ger. Verificada)`.
  Em vez de zerar o patamar, desconta o desvio.

A planilha traz as duas, mais a metodologia nova aplicada a **todos** os meses e a
diferença entre elas. A base dos 5% e a data de virada ficam em
`src/condicionantes.py` (`TOL_BASE`, `VIGENCIA_NOVA`), fáceis de trocar se a
interpretação mudar.

> **É reconstrução, não apuração oficial.** Partimos da Geração de Referência já
> publicada — que é onde o item 5.2.2.10 começa. O ONS calcula essa referência por
> função de produtividade e faz o rateio por usina do item 5.2.2.11. Diferenças contra
> a apuração oficial são esperadas.

### O atraso do PLD

A CCEE só publica o PLD **horário** depois do fechamento contábil do mês. Enquanto o
mês corrente não fecha, o pipeline usa o **PLD médio semanal** (publicado em poucos
dias) e marca essas linhas como `semanal_media` — no painel aparece um aviso, e na
planilha há uma coluna "PLD (fonte)". Quando o PLD horário sai, o `state.json` detecta
a mudança, reprocessa aqueles meses e o valor é substituído pelo definitivo. Nada
precisa ser feito à mão.

## Colocar no ar (uma vez)

1. Crie um repositório **público** no GitHub (público = GitHub Actions e Pages de graça)
   e envie estes arquivos:

   ```bash
   git init && git add -A
   git commit -m "curtailment ONS: pipeline inicial"
   git branch -M main
   git remote add origin https://github.com/<seu-usuario>/<repo>.git
   git push -u origin main
   ```

2. Em **Settings → Pages**, escolha *Source: GitHub Actions*.
3. Em **Settings → Actions → General → Workflow permissions**, marque
   *Read and write permissions*.
4. Em **Actions**, rode "Atualiza curtailment" manualmente com **full = true**.
   A primeira carga baixa 90 arquivos mensais do ONS e leva de 30 a 60 minutos.
   Depois disso cada execução diária leva 2 a 5 minutos, porque só recarrega os meses
   que o ONS republicou.
5. O painel fica em `https://<seu-usuario>.github.io/<repo>/`. Qualquer pessoa com o
   link abre; não precisa de conta.

Quer o painel privado? Repositório privado exige GitHub Pro para publicar Pages —
a alternativa gratuita é hospedar `docs/` no Cloudflare Pages ou no Netlify com
acesso restrito.

## Escolher os conjuntos do Excel

Edite `config/conjuntos.yml` e rode o workflow uma vez com **full = true** (o detalhe
semi-horário só é guardado para os conjuntos listados, então conjuntos novos precisam
de uma recarga). Os ids ficam na coluna `id_ons` do CSV que o painel exporta.

## Estrutura

```
src/sources.py      catálogo de datasets, códigos de razão, mapa de submercados
src/fetch.py        download, listagem do bucket S3, cache por ETag, state.json
src/ons.py          leitura e agregação dos arquivos do ONS
src/ccee.py         PLD horário + fallback semanal, e a multiplicação MWh × PLD
src/condicionantes.py  o teste de tolerância e o recálculo do item 5.2.2.10
src/excel_report.py planilha (Leia-me, Meus conjuntos, Mês a mês, Condicionantes,
                    Perdas R$, Minutos, uma aba por conjunto, Semi-horário, PLD)
src/site_build.py   monta o payload e injeta no template
src/template.html   o painel (HTML/CSS/JS puro, sem dependências externas)
src/build.py        orquestra tudo
data/agg_daily/     agregado diário, uma partição parquet por mês
data/detail_selected/  semi-horário dos seus conjuntos, uma partição por mês
data/conjuntos.parquet, data/pld.parquet, data/state.json
docs/               o que vai para o GitHub Pages: index.html + .xlsx
```

## Rodar localmente

```bash
pip install -r requirements.txt
python -m src.build --full     # primeira vez
python -m src.build            # incremental
open docs/index.html
```

## Limitações conhecidas

- Os arquivos intra-semi-hora (`*_intrasemihora`) só existem a partir de jan/2026 e
  hoje não são usados no cálculo — o valor sai do patamar semi-horário, que é a base
  da apuração. O catálogo já os traz em `src/sources.py` se você quiser o detalhe de
  minuto a minuto de cada evento.
- As coordenadas são as da **subestação coletora** do conjunto (campo do ONS), não do
  centro do parque.
- Conjuntos sem correspondência em `fator-capacidade-2` aparecem na tabela mas não no
  mapa.
