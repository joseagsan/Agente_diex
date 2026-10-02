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
planilhas direto do Google Sheets — abas de crédito de `ABAS_NC_ORIGEM`
(`160482`, `167482`) e `Corrente_Consolidado`/`RP_Adaptado` da planilha
Consolidado —, mostra o relatório (SAC, Empenhos/Liquidação e Indicadores)
e oferece o download do HTML. As justificativas por NC/NE são editadas na
própria página e ficam salvas na aba `JUSTIFICATIVAS_CMT` da planilha do
SSAC. As abas podem ser trocadas pelas variáveis `ABAS_NC_ORIGEM`,
`ABA_RELATORIO_CORRENTE`, `ABA_RELATORIO_RP` e `ABA_JUSTIFICATIVAS` (ver
`config.py`).

A seção de Material Permanente só existe pela linha de comando (`--material`).

## Comparativo semanal (entrou / saiu de tela)

A planilha só mostra o saldo de hoje. Por isso, cada geração do relatório
registra uma "fotografia" do saldo em tela de cada NC do 10º GAC Sl (uma por
dia) e o SAC compara com a fotografia mais recente que tenha pelo menos 7
dias — ou, enquanto o histórico for mais curto, com a mais antiga
disponível. Os cards mostram quanto **entrou** (crédito novo ou saldo que
aumentou) e quanto **saiu** de tela (empenhado, recolhido ou retirado da
planilha), e o quadro de comparativo lista as NCs que mudaram.

- No app, as fotografias ficam na aba `HISTORICO_EM_TELA` da planilha do
  SSAC (variável `ABA_HISTORICO_EM_TELA`), gravada na primeira geração do dia.
- Pela linha de comando, ficam em `dados/historico_em_tela.csv` (opção
  `--historico`). Use `--data dd/mm/aaaa` se os CSVs não forem de hoje.
- O período mínimo (7 dias) é `DIAS_COMPARATIVO` em `core.py`.

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
  lançamento: só a primeira da lista (representante) pode entrar, e o grupo
  sai da lista se a representante tiver saldo residual.

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
- O **comparativo** cobre o saldo em tela das NCs (seção SAC). Caixas de
  **alerta** (pendências específicas) e comparativos de NE/RP ainda não são
  gerados automaticamente.
