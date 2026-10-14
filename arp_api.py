"""
Integração com a API pública (sem autenticação, sem certificado digital)
de Atas de Registro de Preço (ARP) do Portal de Dados Abertos de Compras
Públicas — dadosabertos.compras.gov.br/modulo-arp.

Referência: https://dadosabertos.compras.gov.br/swagger-ui/index.html
(tag "08 - ARP"). Todos os endpoints abaixo são de leitura pública.
"""
from __future__ import annotations
import logging
import requests

logger = logging.getLogger(__name__)

BASE = "https://dadosabertos.compras.gov.br/modulo-arp"
TIMEOUT = (5, 20)
MAX_PAGINAS = 20  # trava de segurança contra loop de paginação


def _get(endpoint: str, params: dict) -> dict:
    url = f"{BASE}/{endpoint}"
    r = requests.get(url, params=params, headers={"Accept": "application/json"}, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


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


def consultar_pregao_da_ata(codigo_ug: str, numero_ata: str, data_vigencia_inicial: str) -> dict:
    """Tenta descobrir o pregão/compra de origem de uma ata. O endpoint de
    itens (consultarARPItem) não aceita o número da ata como filtro, então
    usamos a própria data de vigência inicial da ata (retornada por
    consultar_arps_vencendo) como janela exata de busca — min=max — e
    filtramos o resultado pelo número da ata. Retorna {} se não achar."""
    itens = consultar_itens_arp(codigo_ug, data_vigencia_inicial, data_vigencia_inicial)
    for it in itens:
        if it.get("numeroAtaRegistroPreco") == numero_ata:
            return {
                "numeroCompra": it.get("numeroCompra", ""),
                "modalidade":   it.get("nomeModalidadeCompra", ""),
            }
    return {}


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
