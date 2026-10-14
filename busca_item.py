"""
Busca de item por palavra-chave nas atas de registro de preço vigentes de
uma UASG, e classificação do resultado em PARTICIPANTE (empenho imediato),
CARONA (adesão, limitada a 50% da quantidade registrada da gerenciadora) ou
SEM_OPCAO — seguindo a mesma lógica operacional usada na consulta manual ao
Contratos.gov.br (ver skill "Busca de itens em Atas de Registro de Preços").
"""
from __future__ import annotations

import unicodedata
from datetime import date, timedelta

from arp_api import consultar_arps_vencendo, consultar_itens_arp


def _normalizar(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode()
    return s.lower()


def _bate_palavra_chave(item: dict, palavra_chave: str) -> bool:
    alvo = _normalizar(palavra_chave)
    texto = _normalizar(
        (item.get("descricaoItem", "") or "") + " " + (item.get("descricaoDetalhada", "") or "")
    )
    return alvo in texto


def buscar_itens_por_uasg(codigo_ug: str, palavra_chave: str, dias_vigencia_max: int = 730) -> list[dict]:
    """Varre as atas vigentes de uma UASG gerenciadora e retorna os itens cuja
    descrição bate com a palavra-chave. Cada item vem com a chave extra
    "_ata" apontando para a ata de origem (vigência, gerenciadora etc)."""
    hoje = date.today()
    data_min = hoje.strftime("%Y-%m-%d")
    data_max = (hoje + timedelta(days=dias_vigencia_max)).strftime("%Y-%m-%d")

    atas = consultar_arps_vencendo(codigo_ug, data_min, data_max)

    encontrados = []
    cache_itens: dict[str, list[dict]] = {}
    for ata in atas:
        vig_inicial = ata.get("dataVigenciaInicial", "")
        if not vig_inicial:
            continue
        if vig_inicial not in cache_itens:
            cache_itens[vig_inicial] = consultar_itens_arp(codigo_ug, vig_inicial, vig_inicial)
        numero_ata = ata.get("numeroAtaRegistroPreco", "")
        for it in cache_itens[vig_inicial]:
            if it.get("numeroAtaRegistroPreco") != numero_ata:
                continue
            if _bate_palavra_chave(it, palavra_chave):
                encontrados.append({**it, "_ata": ata})
    return encontrados


def classificar_item(item: dict, unidades: list[dict], codigo_ug_propria: str) -> dict:
    """Classifica um item já encontrado (ver buscar_itens_por_uasg) para a
    UASG própria, usando as linhas de "unidades do item" (endpoint
    3_consultarUnidadesItem — quem participa, saldo de cada um, se aceita
    adesão)."""
    ata = item.get("_ata", {})
    ug_gerenciadora = ata.get("codigoUnidadeGerenciadora", "")

    linha_propria = next(
        (u for u in unidades if str(u.get("unidadeParticipante", "")) == str(codigo_ug_propria)), None
    )
    aceita_adesao = any(u.get("aceitaAdesao") for u in unidades) if unidades else None

    resultado = {
        "item": item,
        "ata": ata,
        "situacao": "SEM_OPCAO",
        "quantidade_utilizavel": 0.0,
        "origem_quantidade": "",
        "aceita_adesao": aceita_adesao,
    }

    somos_participantes = linha_propria is not None or str(codigo_ug_propria) == str(ug_gerenciadora)

    if somos_participantes:
        if linha_propria:
            registrada = _num(linha_propria.get("quantidadeRegistrada"))
            empenhada = _num(linha_propria.get("quantidadeEmpenhada"))
        else:
            # somos a própria gerenciadora: usa os totais do item
            registrada = _num(item.get("quantidadeRegistrada"))
            empenhada = 0.0
        resultado["situacao"] = "PARTICIPANTE"
        resultado["quantidade_utilizavel"] = max(registrada - empenhada, 0)
        resultado["origem_quantidade"] = "Quantidade disponível (participante/gerenciadora)"
    elif aceita_adesao:
        registrada_gerenciadora = _num(item.get("quantidadeRegistrada"))
        resultado["situacao"] = "CARONA"
        resultado["quantidade_utilizavel"] = int(registrada_gerenciadora * 0.5)
        resultado["origem_quantidade"] = "50% da quantidade registrada da gerenciadora"

    return resultado


def _num(v) -> float:
    try:
        return float(str(v).replace(",", "."))
    except (ValueError, TypeError):
        return 0.0
