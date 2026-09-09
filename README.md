# Curtailment no SIN — ONS + PLD

Painel público e planilha com a energia eólica e solar **cortada por ordem do ONS**
(constrained-off), atualizados sozinhos duas vezes por dia pelo GitHub Actions e
publicados no GitHub Pages.

- **Painel:** `https://<seu-usuario>.github.io/<repo>/` — mapa dos conjuntos, energia
  não gerada por razão de restrição, série diária e tabela filtrável (com download em CSV).
- **Excel:** `docs/curtailment_solar.xlsx` — só fotovoltaica, só os conjuntos listados em
  `config/conjuntos.yml`, com detalhe semi-horário, PLD da hora e valor em R$.

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
src/excel_report.py planilha (Leia-me, Resumo, Mensal, Diário, Semi-horário, PLD)
src/site_build.py   monta o payload e injeta no template
src/template.html   o painel (HTML/CSS/JS puro, sem dependências externas)
src/build.py        orquestra tudo
data/               tabelas processadas (parquet) + state.json — commitadas pelo robô
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
