"""
Mapeamento dos pregões (SRP) em que a nossa UASG é PARTICIPANTE (ou a
própria gerenciadora), com os itens em que temos quantidade registrada —
para responder rápido perguntas como "qual pregão tem material de
expediente pra gente?" sem refazer a varredura na API a cada busca.

A API pública de ARP não filtra por unidade participante: a única forma é
olhar, item a item, as "unidades do item" (3_consultarUnidadesItem) das
atas de cada UASG gerenciadora conhecida. Como isso custa uma chamada por
item (~4 s cada), o resultado fica guardado em duas abas da planilha do
SSAC e a atualização é incremental — só pregões ainda não verificados são
varridos de novo, a menos que se peça para remapear tudo.

Uso fora do app (a primeira varredura pode levar uns 15 min):
    python pregoes_participante.py            # só pregões novos
    python pregoes_participante.py --forcar   # remapeia tudo
"""
from __future__ import annotations

import logging
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta
from typing import Callable

from arp_api import consultar_adesao_item, consultar_arps_vencendo, consultar_itens_arp

logger = logging.getLogger(__name__)

ABA_ITENS = "SSAC_PREGOES_PARTICIPANTE"
ABA_VERIFICADOS = "SSAC_PREGOES_VERIFICADOS"

COLUNAS_ITENS = [
    "ID_COMPRA", "UASG_GERENCIADORA", "GERENCIADORA", "PREGAO", "OBJETO", "ATA",
    "VIGENCIA_FINAL", "ITEM", "DESCRICAO", "PDM", "TIPO", "FORNECEDOR", "CNPJ",
    "VALOR_UNIT", "PAPEL", "QTD_NOSSA", "ACEITA_ADESAO", "LINK_PNCP",
]
COLUNAS_VERIFICADOS = [
    "ID_COMPRA", "UASG_GERENCIADORA", "PREGAO", "OBJETO", "PARTICIPA", "ITENS_NOSSOS", "VERIFICADO_EM",
]

PARALELISMO = 16  # chamadas simultâneas à API (cada uma leva ~4 s)
ANOS_VIGENCIA = 3  # janela de vigência final considerada "vigente"


# ── Varredura na API (sem planilha) ─────────────────────────────────────────
def _atas_vigentes(codigo_ug: str) -> list[dict]:
    hoje = date.today()
    atas = consultar_arps_vencendo(
        codigo_ug, hoje.strftime("%Y-%m-%d"),
        (hoje + timedelta(days=365 * ANOS_VIGENCIA)).strftime("%Y-%m-%d"),
        tamanho_pagina=500,
    )
    return [a for a in atas if not a.get("ataExcluido")]


def _itens_das_atas(codigo_ug: str, atas: list[dict]) -> list[dict]:
    """Itens de todas as atas, numa única janela de vigência inicial, cada
    um com "_ata" apontando para a ata de origem."""
    vigs = [a["dataVigenciaInicial"] for a in atas if a.get("dataVigenciaInicial")]
    if not vigs:
        return []
    por_ata = {(a.get("numeroAtaRegistroPreco"), a.get("idCompra")): a for a in atas}
    itens = consultar_itens_arp(codigo_ug, min(vigs), max(vigs), tamanho_pagina=500)
    resultado = []
    for it in itens:
        ata = por_ata.get((it.get("numeroAtaRegistroPreco"), it.get("idCompra")))
        if ata and not it.get("itemExcluido"):
            resultado.append({**it, "_ata": ata})
    return resultado


def _linha_item(it: dict, unidade: dict) -> dict:
    ata = it["_ata"]
    return {
        "ID_COMPRA":         it.get("idCompra", ""),
        "UASG_GERENCIADORA": str(ata.get("codigoUnidadeGerenciadora", "")),
        "GERENCIADORA":      ata.get("nomeUnidadeGerenciadora", ""),
        "PREGAO":            f"{ata.get('numeroCompra', '')}/{ata.get('anoCompra', '')}",
        "OBJETO":            ata.get("objeto", "") or "",
        "ATA":               it.get("numeroAtaRegistroPreco", ""),
        "VIGENCIA_FINAL":    ata.get("dataVigenciaFinal", ""),
        "ITEM":              it.get("numeroItem", ""),
        "DESCRICAO":         it.get("descricaoItem", "") or "",
        "PDM":               it.get("nomePdm", "") or "",
        "TIPO":              it.get("tipoItem", "") or "",
        "FORNECEDOR":        it.get("nomeRazaoSocialFornecedor", "") or "",
        "CNPJ":              it.get("niFornecedor", "") or "",
        "VALOR_UNIT":        it.get("valorUnitario", ""),
        "PAPEL":             unidade.get("tipoUnidade", ""),
        "QTD_NOSSA":         unidade.get("quantidadeRegistrada", ""),
        "ACEITA_ADESAO":     "SIM" if unidade.get("aceitaAdesao") else "NÃO",
        "LINK_PNCP":         ata.get("linkAtaPNCP", "") or "",
    }


def _nossa_unidade(it: dict, ug_propria: str, tentativas: int = 3) -> dict | None:
    # a API dá timeout de vez em quando; sem retentar, um item derruba a UASG toda
    for n in range(tentativas):
        try:
            unidades = consultar_adesao_item(
                it.get("numeroAtaRegistroPreco", ""),
                str(it["_ata"].get("codigoUnidadeGerenciadora", "")),
                it.get("numeroItem", ""),
            )
            break
        except Exception:
            if n == tentativas - 1:
                raise
    return next((u for u in unidades if str(u.get("codigoUnidade", "")) == str(ug_propria)), None)


def mapear_uasg(codigo_ug: str, ug_propria: str, ja_verificados: set[str],
                ao_concluir_pregao: Callable[[dict, list[dict]], None],
                progresso: Callable[[str], None] | None = None) -> None:
    """Varre os pregões vigentes de uma UASG gerenciadora que ainda não estão
    em ja_verificados e, a cada pregão concluído, chama
    ao_concluir_pregao(verificado, itens_nossos) — assim o que já foi
    varrido fica salvo mesmo se a varredura for interrompida no meio."""
    atas = _atas_vigentes(codigo_ug)
    itens = _itens_das_atas(codigo_ug, atas)

    por_compra: dict[str, list[dict]] = {}
    for it in itens:
        por_compra.setdefault(it.get("idCompra", ""), []).append(it)

    pendentes = {c: its for c, its in por_compra.items() if c and c not in ja_verificados}
    for n, (id_compra, its) in enumerate(pendentes.items(), start=1):
        ata = its[0]["_ata"]
        pregao = f"{ata.get('numeroCompra', '')}/{ata.get('anoCompra', '')}"
        if progresso:
            progresso(f"UASG {codigo_ug} · pregão {pregao} ({n}/{len(pendentes)}, {len(its)} itens)")

        with ThreadPoolExecutor(PARALELISMO) as ex:
            unidades = list(ex.map(lambda it: _nossa_unidade(it, ug_propria), its))

        nossos = [_linha_item(it, u) for it, u in zip(its, unidades) if u]
        verificado = {
            "ID_COMPRA":         id_compra,
            "UASG_GERENCIADORA": str(codigo_ug),
            "PREGAO":            pregao,
            "OBJETO":            ata.get("objeto", "") or "",
            "PARTICIPA":         "SIM" if nossos else "NÃO",
            "ITENS_NOSSOS":      str(len(nossos)),
            "VERIFICADO_EM":     datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        ao_concluir_pregao(verificado, nossos)


# ── Persistência na planilha do SSAC ────────────────────────────────────────
def _ws(titulo: str, colunas: list[str]):
    from config import SHEET_ID_NC
    from sheets_nc import _conectar

    planilha = _conectar().open_by_key(SHEET_ID_NC)
    try:
        return planilha.worksheet(titulo)
    except Exception:
        ws = planilha.add_worksheet(title=titulo, rows=1000, cols=len(colunas))
        ws.update("A1", [colunas], value_input_option="RAW")
        logger.info("Aba %s criada.", titulo)
        return ws


def _ler(titulo: str, colunas: list[str]) -> list[dict]:
    rows = _ws(titulo, colunas).get_all_values()
    if len(rows) < 2:
        return []
    return [dict(zip(rows[0], r)) for r in rows[1:] if r and r[0].strip()]


def ler_itens_mapeados() -> list[dict]:
    return _ler(ABA_ITENS, COLUNAS_ITENS)


def ler_pregoes_verificados() -> list[dict]:
    return _ler(ABA_VERIFICADOS, COLUNAS_VERIFICADOS)


def atualizar_mapeamento(uasgs_gerenciadoras: list[str], ug_propria: str, forcar: bool = False,
                         progresso: Callable[[str], None] | None = None) -> dict:
    """Varre as UASGs gerenciadoras e grava nas abas do SSAC. Sem forcar,
    pula os pregões já verificados (atualização incremental)."""
    ws_itens = _ws(ABA_ITENS, COLUNAS_ITENS)
    ws_verif = _ws(ABA_VERIFICADOS, COLUNAS_VERIFICADOS)
    if forcar:
        for ws, cols in ((ws_itens, COLUNAS_ITENS), (ws_verif, COLUNAS_VERIFICADOS)):
            ws.clear()
            ws.update("A1", [cols], value_input_option="RAW")
        ja = set()
    else:
        ja = {v["ID_COMPRA"] for v in ler_pregoes_verificados()}

    resumo = {"pregoes_verificados": 0, "pregoes_participante": 0, "itens": 0, "erros": []}

    def _salvar(verificado: dict, nossos: list[dict]):
        if nossos:
            ws_itens.append_rows([[str(l.get(c, "")) for c in COLUNAS_ITENS] for l in nossos],
                                 value_input_option="RAW")
            resumo["pregoes_participante"] += 1
            resumo["itens"] += len(nossos)
        ws_verif.append_row([verificado[c] for c in COLUNAS_VERIFICADOS], value_input_option="RAW")
        resumo["pregoes_verificados"] += 1
        ja.add(verificado["ID_COMPRA"])

    for ug in dict.fromkeys(u.strip() for u in uasgs_gerenciadoras if u and u.strip()):
        try:
            mapear_uasg(ug, ug_propria, ja, _salvar, progresso)
        except Exception as e:
            logger.exception("Erro ao mapear UASG %s", ug)
            resumo["erros"].append(f"UASG {ug}: {e}")
    return resumo


def uasgs_gerenciadoras_cadastradas(ug_propria: str) -> list[str]:
    """A própria UASG + as cadastradas em "Buscar Item" e em "Atas"."""
    from ugs_monitoradas import ler_ugs_monitoradas
    from uasgs_busca import ler_uasgs_busca

    ugs = [ug_propria]
    ugs += [u["UASG"] for u in ler_uasgs_busca()]
    ugs += [u["UG"] for u in ler_ugs_monitoradas()]
    return list(dict.fromkeys(u.strip() for u in ugs if u and u.strip()))


# ── Busca em memória ────────────────────────────────────────────────────────
_STOPWORDS = {"de", "da", "do", "das", "dos", "e", "para", "com", "em", "a", "o"}


def _normalizar(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode()
    return s.lower()


def _termos(busca: str) -> list[str]:
    return [t for t in _normalizar(busca).split() if t not in _STOPWORDS]


def _bate(texto: str, termos: list[str]) -> bool:
    alvo = _normalizar(texto)
    return all(t in alvo for t in termos)


def buscar(itens: list[dict], busca: str) -> tuple[set[str], list[dict]]:
    """Retorna (pregões cujo OBJETO bate com a busca, itens cuja descrição
    ou PDM batem). Todas as palavras precisam aparecer, sem diferenciar
    maiúsculas/acentos, ignorando "de", "da" etc."""
    termos = _termos(busca)
    if not termos:
        return set(), []
    pregoes = {it["ID_COMPRA"] for it in itens if _bate(it.get("OBJETO", ""), termos)}
    batem = [it for it in itens if _bate(f"{it.get('DESCRICAO', '')} {it.get('PDM', '')}", termos)]
    return pregoes, batem


if __name__ == "__main__":
    import argparse

    from config import UG_PADRAO

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    ap = argparse.ArgumentParser(description="Mapeia os pregões em que a nossa UASG é participante.")
    ap.add_argument("--forcar", action="store_true", help="remapeia tudo do zero")
    ap.add_argument("--uasg", action="append", help="UASG gerenciadora (repetível); padrão: as cadastradas")
    args = ap.parse_args()

    alvo = args.uasg or uasgs_gerenciadoras_cadastradas(UG_PADRAO)
    print(atualizar_mapeamento(alvo, UG_PADRAO, forcar=args.forcar, progresso=print))
