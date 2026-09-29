# Relatório Semanal — 10º GAC Sl

Gera o **Documento de Consulta ao Comandante** automaticamente a partir dos
exports em CSV das planilhas do Google Sheets, sem precisar editar HTML à
mão toda semana.

## Estrutura da pasta

```
relatorio_gac/
├── scripts/
│   ├── core.py               # regras de negócio (filtros, fórmulas, exclusões)
│   └── gerar_relatorio.py    # monta as seções e gera o HTML final
├── templates/
│   └── documento_consulta_comandante.template.html   # visual do relatório
├── dados/
│   └── justificativas.csv    # (criado na 1ª execução) justificativas manuais por NC/NE
└── README.md
```

## Uso pelo app (SSAC)

No app Streamlit, menu **Ferramentas → 📑 Consulta ao Cmt**. A página lê as
planilhas direto do Google Sheets (abas `160482`/`167482`,
`Corrente_Consolidado`, `RP_Adaptado` e a planilha de Material
Permanente), mostra o relatório e oferece o download do HTML. As
justificativas por NC/NE são editadas na própria página e ficam salvas na
aba `JUSTIFICATIVAS_CMT` da planilha do SSAC. As abas/planilhas podem ser
trocadas pelas variáveis `ABAS_RELATORIO_CREDITO`, `ABA_RELATORIO_CORRENTE`,
`ABA_RELATORIO_RP`, `SHEET_ID_MATERIAL`, `ABA_MATERIAL` e
`ABA_JUSTIFICATIVAS` (ver `config.py`).

A planilha de Material Permanente precisa estar compartilhada (Leitor) com
a conta de serviço do app; se não estiver, a página avisa e aceita o CSV
por upload.

## Uso semanal pela linha de comando

1. No Google Sheets, exporte como CSV as abas que você atualizou nesta
   rodada (Arquivo → Fazer download → Valores separados por vírgula):
   - **Controle de Crédito** — uma aba por UG (160482, 167482)
   - **Corrente_Consolidado** — empenhos/liquidação do ano corrente
   - **RP_Adaptado** — Restos a Pagar
   - **Aquisição de Material Permanente**

2. Rode o script apontando para os arquivos baixados. Você só precisa
   passar as fontes que atualizou nesta rodada — se omitir uma, a seção
   correspondente simplesmente não entra no relatório desta vez:

   ```bash
   python3 scripts/gerar_relatorio.py \
     --credito "Controle_Credito_160482.csv" "Controle_Credito_167482.csv" \
     --consolidado "Corrente_Consolidado.csv" \
     --rp "RP_Adaptado.csv" \
     --material "Material_Permanente.csv" \
     --saida documento_consulta_comandante.html
   ```

3. Abra o `documento_consulta_comandante.html` gerado no navegador para
   conferir.

## Justificativas manuais

As justificativas por NC/NE (o texto que explica cada pendência) não vêm de
nenhuma planilha automaticamente — são texto livre que só você sabe. Na
primeira execução, o script cria `dados/justificativas.csv` com duas
colunas: `numero` (a NC ou NE) e `justificativa`. Preencha essa planilha à
mão e rode o script de novo — ele casa automaticamente pelo número.

## Onde ajustar as regras (sem mexer no script principal)

Tudo isso está centralizado em `scripts/core.py`, na classe `Config`, no
topo do arquivo:

- **`RESP_ACEITAS`** — quais valores da coluna RESP identificam o 10º GAC Sl
  (já trata as variações de acento/º/° automaticamente).
- **`RESIDUAL_VALOR_ABS_MAX` / `RESIDUAL_PERCENTUAL_MAX`** — a partir de
  quando uma NC é considerada "saldo residual" e sai da lista principal do
  SAC.
- **`GRUPOS_DUPLICATAS`** — grupos de NCs tratadas como duplicata de
  lançamento (mantém só a primeira da lista como representante).

As fórmulas fixadas ficam como propriedades das classes `NotaCredito` e
`NotaEmpenho` no mesmo arquivo:

- `% Empenho = Empenhado / (Recebido − Recolhido)` — **sempre sobre o
  Recebido Líquido**, nunca sobre o Recebido bruto (regra explícita sua).
- `% Liquidação = (Liquidado + Pago) / (Total − Anulado)`

## Ajustar o visual

Edite `templates/documento_consulta_comandante.template.html` — cores,
cabeçalho, menu lateral, estilo dos cards (`.resumo-cards`/`.card`), das
caixas de alerta (`.alerta-box`) e comparativo (`.comparativo-box`), e das
tabelas (`.tabela-scroll`). O script não precisa ser tocado para isso.

## Limitações atuais / próximos passos possíveis

- As seções de **Viaturas**, **Descarga de Viatura** e **Controle de
  Obras (PO)** ainda não têm loader automático neste pacote — essas fontes
  (planilhas de viaturas, situação de descarga, obras) têm formato mais
  livre e foram, até agora, sempre editadas manualmente. Dá para
  automatizar do mesmo jeito assim que você quiser: é só um novo `carregar_*`
  em `core.py` + uma nova função `secao_*` em `gerar_relatorio.py`.
- A geração do **.docx** (Word) ainda é feita à parte (script Node.js com
  o pacote `docx`, gerado nas rodadas anteriores no chat). Posso portar essa
  etapa para dentro deste pacote também, se for útil.
- Caixas de **comparativo** (o que mudou desde a semana passada) e de
  **alerta** (pendências específicas) ainda não são geradas automaticamente
  — hoje dependem de comparar "à mão" com a rodada anterior. Dá para
  automatizar guardando o HTML/JSON da rodada anterior e comparando os
  totais, se quiser essa evolução.
