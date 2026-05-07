"""
Assistente conversacional para controle financeiro usando Claude API.
"""
import logging
from datetime import datetime

import anthropic

from config import ANTHROPIC_API_KEY, OM_PADRAO, UG_PADRAO

logger = logging.getLogger(__name__)

_SYSTEM = f"""Você é DIEX Assistant, assistente especializado em controle financeiro militar da {OM_PADRAO} (UG {UG_PADRAO}).

Suas funções:
- Responder consultas sobre saldos de NCs, status de requisições e prazos
- Calcular totais, percentuais e resumos financeiros
- Alertar sobre NCs próximas do vencimento ou com saldo crítico
- Gerar análises e recomendações baseadas nos dados disponíveis

Responda sempre em português, de forma clara e objetiva.
Use formatação markdown (tabelas, negrito, listas) quando melhorar a legibilidade.
Valores monetários no formato R$ X.XXX,XX."""


def _resumir_dados(ncs: list[dict], reqs: list[dict]) -> str:
    def _parse(s) -> float:
        try:
            return float(str(s).replace("R$", "").replace(" ", "").replace(".", "").replace(",", "."))
        except Exception:
            return 0.0

    saldo_total = sum(_parse(nc.get("SALDO NC", 0)) for nc in ncs)
    ncs_em_tela = [nc for nc in ncs if nc.get("SITU") == "EM TELA"]
    reqs_pendentes = [r for r in reqs if r.get("SITUAÇÃO") == "Pendente"]
    hoje = datetime.today().strftime("%d/%m/%Y")

    linhas = [
        f"**Dados em {hoje} — UG {UG_PADRAO} ({OM_PADRAO})**",
        f"- Total NCs: {len(ncs)} ({len(ncs) - len(ncs_em_tela)} OK, {len(ncs_em_tela)} EM TELA)",
        f"- Saldo total disponível: R$ {saldo_total:_.2f}".replace("_", ".").replace(".", ",", 1)
        if False else f"- Saldo total disponível: {_fmt(saldo_total)}",
        f"- REQs pendentes: {len(reqs_pendentes)} de {len(reqs)}",
        "",
        "**NCs:**",
    ]
    for nc in ncs:
        linhas.append(
            f"- {nc.get('NC','')} | {nc.get('ORGÃO','')} | {nc.get('FINALIDADE','')[:60]} "
            f"| Saldo: {nc.get('SALDO NC','')} | Status: {nc.get('SITU','')} | Prazo: {nc.get('PRAZO','')}"
        )

    linhas.append("\n**REQs:**")
    for req in reqs[:40]:
        linhas.append(
            f"- REQ {req.get('REQ','')} | {req.get('EMPRESA','')} "
            f"| {req.get('VALOR','')} | NC: {req.get('NC','')} | {req.get('SITUAÇÃO','')}"
        )

    return "\n".join(linhas)


def _fmt(v: float) -> str:
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def chat(mensagem: str, historico: list[dict], ncs: list[dict], reqs: list[dict]) -> str:
    """
    historico: lista de {'role': 'user'|'assistant', 'content': str}
    Retorna a resposta do assistente como string.
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    contexto = _resumir_dados(ncs, reqs)
    system = _SYSTEM + f"\n\n## Dados Atuais\n{contexto}"

    messages = list(historico) + [{"role": "user", "content": mensagem}]

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2000,
        system=system,
        messages=messages,
    )

    resposta = response.content[0].text
    logger.info("Assistente respondeu (%d chars).", len(resposta))
    return resposta
