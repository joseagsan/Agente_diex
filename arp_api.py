"""
Integração com a API pública (sem autenticação, sem certificado digital)
de Atas de Registro de Preço (ARP) do Portal de Dados Abertos de Compras
Públicas — dadosabertos.compras.gov.br/modulo-arp.

Referência: https://dadosabertos.compras.gov.br/swagger-ui/index.html
(tag "08 - ARP"). Todos os endpoints abaixo são de leitura pública.
"""
from __future__ import annotations
import logging
import random
import time

import requests

logger = logging.getLogger(__name__)

BASE = "https://dadosabertos.compras.gov.br/modulo-arp"
TIMEOUT = (5, 20)
MAX_PAGINAS = 20  # trava de segurança contra loop de paginação


TENTATIVAS = 6  # a API devolve 429 (Too Many Requests) quando apertamos demais


def _espera(tentativa: int, r: requests.Response | None) -> float:
    retry_after = r.headers.get("Retry-After", "") if r is not None else ""
    if retry_after.isdigit():
        return min(float(retry_after), 120.0)
    return min(2 ** (tentativa + 1), 60) * random.uniform(0.75, 1.25)


def _get(endpoint: str, params: dict) -> dict:
    """GET com retentativa e espera crescente em 429, erro 5xx e timeout."""
    url = f"{BASE}/{endpoint}"
    for tentativa in range(TENTATIVAS):
        r = None
        try:
            r = requests.get(url, params=params, headers={"Accept": "application/json"}, timeout=TIMEOUT)
            if r.status_code != 429 and r.status_code < 500:
                r.raise_for_status()
                return r.json()
        except (requests.Timeout, requests.ConnectionError):
            pass
        if tentativa == TENTATIVAS - 1:
            if r is not None:
                r.raise_for_status()
            raise requests.Timeout(f"Sem resposta da API após {TENTATIVAS} tentativas: {url}")
        espera = _espera(tentativa, r)
        logger.warning("API %s respondeu %s; nova tentativa em %.0f s.",
                       endpoint, r.status_code if r is not None else "timeout", espera)
        time.sleep(espera)


def _paginar(endpoint: str, params: dict) -> list[dict]:
    resultado = []
    pagina = params.get("pagina", 1)
    for _ in range(MAX_PAGINAS):
        params["pagina"] = pagina
        payload = _get(endpoint, params)
        resultado.extend(payload.get("resultado", []))
        if not payload.get("paginasRestantes"):
            break
        pagina += 1
    return resultado


def consultar_arps_vencendo(codigo_ug: str, data_vig_final_min: str, data_vig_final_max: str,
                             tamanho_pagina: int = 50) -> list[dict]:
    """Lista atas de uma UG gerenciadora cuja vigência final cai no período
    [data_vig_final_min, data_vig_final_max] (formato YYYY-MM-DD)."""
    return _paginar("1.2_consultarARP_FimVigencia", {
        "tamanhoPagina": tamanho_pagina,
        "codigoUnidadeGerenciadora": codigo_ug,
        "dataVigenciaFinalMin": data_vig_final_min,
        "dataVigenciaFinalMax": data_vig_final_max,
    })


def consultar_itens_arp(codigo_ug: str, data_vig_inicial_min: str, data_vig_inicial_max: str,
                         numero_compra: str = "", tamanho_pagina: int = 50) -> list[dict]:
    """Lista itens registrados de uma UG dentro de uma janela de vigência
    inicial (exigida pela API). O endpoint não aceita número da ata
    diretamente — passe numero_compra quando souber, ou filtre o
    resultado no chamador."""
    params = {
        "tamanhoPagina": tamanho_pagina,
        "codigoUnidadeGerenciadora": codigo_ug,
        "dataVigenciaInicialMin": data_vig_inicial_min,
        "dataVigenciaInicialMax": data_vig_inicial_max,
    }
    if numero_compra:
        params["numeroCompra"] = numero_compra
    return _paginar("2_consultarARPItem", params)


def consultar_saldo_item(numero_ata: str, unidade_gerenciadora: str, numero_item: str) -> list[dict]:
    """Saldo de empenho do item: quantidadeRegistrada, quantidadeEmpenhada,
    saldoEmpenho. Resultado vazio é normal — significa que ainda não houve
    empenho registrado para esse item."""
    return _paginar("4_consultarEmpenhosSaldoItem", {
        "tamanhoPagina": 50,
        "numeroAta": numero_ata,
        "unidadeGerenciadora": unidade_gerenciadora,
        "numeroItem": numero_item,
    })


def consultar_adesao_item(numero_ata: str, unidade_gerenciadora: str, numero_item: str) -> list[dict]:
    """Se o item aceita adesão (carona) e o saldo disponível para isso:
    aceitaAdesao, saldoAdesoes, quantidadeEmpenhada."""
    return _paginar("3_consultarUnidadesItem", {
        "tamanhoPagina": 50,
        "numeroAta": numero_ata,
        "unidadeGerenciadora": unidade_gerenciadora,
        "numeroItem": numero_item,
    })


def consultar_adesoes_feitas(numero_ata: str, unidade_gerenciadora: str, numero_item: str) -> list[dict]:
    """Adesões (caronas) já aprovadas para o item, por outras unidades:
    unidadeNaoParticipante, quantidadeAprovadaAdesao, dataAprovacaoAnalise."""
    return _paginar("5_consultarAdesoesItem", {
        "tamanhoPagina": 50,
        "numeroAta": numero_ata,
        "unidadeGerenciadora": unidade_gerenciadora,
        "numeroItem": numero_item,
    })
