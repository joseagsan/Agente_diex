# -*- coding: utf-8 -*-
"""
core.py - Núcleo de carga e cálculo para o Relatório Semanal do 10º GAC Sl

Implementa, de forma centralizada e configurável, todas as regras de negócio
validadas manualmente ao longo das rodadas anteriores:

  - Filtro por RESP = "10° GAC Sl" / "10º GAC Sl" (normaliza o símbolo de grau)
  - % Empenho = Empenhado / (Recebido - Recolhido)   [Recebido Líquido]
  - % Liquidação = (Liquidado + Pago) / (Total - Anulado)
  - Classificação de NC "residual" (saldo em tela pequeno) -> excluída da
    lista "NCs com Saldo Relevante"
  - Lista de NCs tratadas como duplicatas de lançamento, mantendo apenas 1
    representante por grupo

Este módulo não gera HTML: ele produz dicionários/listas "prontos para
template" que o script gerar_relatorio.py consome. Assim, se o layout do
relatório mudar, você mexe só no template; se a fonte de dados mudar
(nova coluna, nova regra), você mexe só aqui.
"""

from __future__ import annotations

import csv
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Configuração (ajuste aqui quando as regras do relatório mudarem)
# ---------------------------------------------------------------------------

class Config:
    # Valores aceitos para a coluna RESP que identificam o 10º GAC Sl.
    # Comparação é feita já normalizada (sem acento, minúsculo, º/° tratados
    # como iguais), então não é preciso listar todas as variações de acento.
    RESP_ACEITAS = ["10 gac sl"]

    # Uma NC é considerada "residual" (fora da lista de saldo relevante)
    # quando o saldo em tela é pequeno em valor absoluto OU pequeno em
    # proporção do recebido. Ajuste os limiares conforme orientação do Cmt.
    RESIDUAL_VALOR_ABS_MAX = 100.0     # R$
    RESIDUAL_PERCENTUAL_MAX = 0.5      # 0.5% do recebido

    # NCs tratadas como duplicata de lançamento: mantém-se apenas a primeira
    # da lista (o "representante"); as demais são removidas da lista de
    # saldo relevante (mas continuam somadas nos totais gerais, a menos que
    # você decida o contrário abaixo).
    GRUPOS_DUPLICATAS: list[list[str]] = [
        ["2026NC004983", "2026NC005625", "2026NC005626", "2026NC005630"],
    ]

    # Se True, o valor das NCs duplicadas (exceto o representante) é
    # descontado do total "em tela" apresentado no relatório.
    DESCONTAR_DUPLICATAS_DO_TOTAL = False


# ---------------------------------------------------------------------------
# Helpers genéricos
# ---------------------------------------------------------------------------

def _strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)
    )


def normaliza_resp(valor: str) -> str:
    """Normaliza a coluna RESP para comparação: sem acento, minúsculo,
    º/° tratados como espaço, espaços colapsados."""
    if valor is None:
        return ""
    v = _strip_accents(str(valor)).lower()
    v = v.replace("º", " ").replace("°", " ")
    v = re.sub(r"\s+", " ", v).strip()
    return v


def eh_10_gac(resp: str) -> bool:
    return normaliza_resp(resp) in Config.RESP_ACEITAS


def parse_valor_brl(valor) -> float:
    """Converte 'R$ 1.234,56', '1234.56', '', None, etc. em float."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    if s == "" or s in ("-", "—"):
        return 0.0
    s = s.replace("R$", "").strip()
    # formato brasileiro: milhar '.' decimal ','
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s) if s not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


def fmt_brl(valor: float) -> str:
    """Formata float como 'R$ 1.234,56'."""
    s = f"{valor:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def fmt_pct(valor: float, casas: int = 2) -> str:
    """valor já em fração (0.9685) -> '96,85%'."""
    s = f"{valor * 100:.{casas}f}".replace(".", ",")
    return f"{s}%"


def ler_csv_dicts(caminho: str | Path) -> list[dict]:
    """Lê um CSV exportado do Google Sheets em UTF-8 (com ou sem BOM) e
    devolve uma lista de dicts com as chaves já 'limpas' (sem quebras de
    linha internas, sem espaços extras)."""
    caminho = Path(caminho)
    with open(caminho, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        linhas = []
        for row in reader:
            limpo = {}
            for k, v in row.items():
                if k is None:
                    continue
                chave = re.sub(r"\s+", " ", k.strip())
                limpo[chave] = v.strip() if isinstance(v, str) else v
            linhas.append(limpo)
        return linhas


# ---------------------------------------------------------------------------
# SAC / Controle de Crédito (NC)
# ---------------------------------------------------------------------------

@dataclass
class NotaCredito:
    nc: str
    ug: str
    data_nc: str
    dias: str
    finalidade: str
    op: str
    recebido: float
    recolhido: float
    empenhado: float
    em_tela: float
    justificativa: str = ""
    situacao: str = ""

    @property
    def recebido_liquido(self) -> float:
        return self.recebido - self.recolhido

    @property
    def pct_empenho(self) -> float:
        """% Empenho = Empenhado / (Recebido - Recolhido) -- regra fixada
        pelo usuário: SEMPRE sobre o Recebido Líquido, nunca sobre o
        Recebido bruto."""
        base = self.recebido_liquido
        if base <= 0:
            return 0.0
        return self.empenhado / base

    @property
    def eh_residual(self) -> bool:
        if abs(self.em_tela) <= Config.RESIDUAL_VALOR_ABS_MAX:
            return True
        if self.recebido > 0:
            pct = abs(self.em_tela) / self.recebido * 100
            if pct <= Config.RESIDUAL_PERCENTUAL_MAX:
                return True
        return False


def _grupo_duplicata_de(nc: str) -> Optional[list[str]]:
    for grupo in Config.GRUPOS_DUPLICATAS:
        if nc in grupo:
            return grupo
    return None


def carregar_notas_credito(*caminhos_csv: str | Path) -> list[NotaCredito]:
    """Carrega um ou mais CSVs de 'Controle de Crédito' (ex.: UG 160482 e
    167482), filtra por RESP = 10º GAC Sl e devolve a lista de NCs.

    Aceita variações de nome de coluna (ÓRGÃO/ORGÃO, 'ENV\\nSECAO', etc.)
    porque ler_csv_dicts() já normaliza espaços/quebras de linha nas chaves.
    """
    notas: list[NotaCredito] = []
    for caminho in caminhos_csv:
        for row in ler_csv_dicts(caminho):
            resp = row.get("RESP", "")
            if not eh_10_gac(resp):
                continue
            nc = row.get("NC", "").strip()
            if not nc:
                continue
            recebido = parse_valor_brl(row.get("RECEBIDO"))
            recolhido = parse_valor_brl(row.get("RECOLHIDO"))
            empenhado = parse_valor_brl(row.get("EMPENHADO"))
            em_tela = parse_valor_brl(row.get("EM TELA"))
            notas.append(
                NotaCredito(
                    nc=nc,
                    ug=row.get("UG", "").strip(),
                    data_nc=row.get("DATA NC", "").strip(),
                    dias=_fmt_dias(row.get("DIAS", "")),
                    finalidade=row.get("FINALIDADE", "").strip(),
                    op=row.get("OP", "").strip(),
                    recebido=recebido,
                    recolhido=recolhido,
                    empenhado=empenhado,
                    em_tela=em_tela,
                    situacao=row.get("SITU", "").strip(),
                )
            )
    return notas


def _fmt_dias(valor) -> str:
    """Coluna DIAS da planilha (dias desde a data da NC); o export pode
    trazer '32' ou '32.0' -- normaliza para inteiro."""
    s = str(valor or "").strip()
    try:
        return str(int(float(s.replace(",", "."))))
    except ValueError:
        return s


def indicador_pct_empenho_global(notas: list[NotaCredito]) -> dict:
    """Consolida o indicador '% Empenho' no nível do universo de NCs do
    10º GAC Sl (não por NC individual): soma Empenhado e soma Recebido
    Líquido, depois divide -- é assim que o card de indicador tem sido
    montado no relatório consolidado."""
    total_recebido = sum(n.recebido for n in notas)
    total_recolhido = sum(n.recolhido for n in notas)
    total_empenhado = sum(n.empenhado for n in notas)
    recebido_liquido = total_recebido - total_recolhido
    pct = total_empenhado / recebido_liquido if recebido_liquido > 0 else 0.0
    return {
        "qtd_ncs": len(notas),
        "total_recebido": total_recebido,
        "total_recolhido": total_recolhido,
        "recebido_liquido": recebido_liquido,
        "total_empenhado": total_empenhado,
        "saldo_a_empenhar": recebido_liquido - total_empenhado,
        "pct_empenho": pct,
    }


def notas_com_saldo_relevante(notas: list[NotaCredito]) -> list[NotaCredito]:
    """Aplica as duas regras de exclusão da lista 'SAC - em tela':
    1) remove NCs residuais (saldo em tela desprezível)
    2) para cada grupo de duplicatas, mantém só o representante (1º do grupo)
    Ordena por NC.
    """
    relevantes = [n for n in notas if not n.eh_residual]

    vistos_de_grupo: set[str] = set()
    resultado: list[NotaCredito] = []
    for n in relevantes:
        grupo = _grupo_duplicata_de(n.nc)
        if grupo is None:
            resultado.append(n)
            continue
        chave_grupo = grupo[0]
        if chave_grupo in vistos_de_grupo:
            continue  # já incluímos o representante deste grupo
        # inclui apenas se esta NC for o representante (primeiro do grupo)
        # presente nos dados; senão, procura o representante entre 'notas'
        if n.nc == grupo[0]:
            resultado.append(n)
            vistos_de_grupo.add(chave_grupo)
        else:
            representante = next((x for x in notas if x.nc == grupo[0]), None)
            if representante is not None and chave_grupo not in vistos_de_grupo:
                resultado.append(representante)
                vistos_de_grupo.add(chave_grupo)

    resultado.sort(key=lambda n: n.nc)
    return resultado


# ---------------------------------------------------------------------------
# Empenhos / Liquidação (NE) -- Corrente_Consolidado e RP_Adaptado
# ---------------------------------------------------------------------------

@dataclass
class NotaEmpenho:
    ne: str
    ug: str
    emissao: str
    fav: str
    nome_fav: str
    finalidade: str
    a_liquidar: float
    em_liquidacao: float
    liquidado: float
    pago: float
    anulado: float
    situacao: str
    e_restos_a_pagar: bool = False

    @property
    def total_bruto(self) -> float:
        return self.a_liquidar + self.em_liquidacao + self.liquidado + self.pago + self.anulado

    @property
    def total_liquido(self) -> float:
        """Total - Anulado, base do % de liquidação."""
        return self.total_bruto - self.anulado

    @property
    def pct_liquidacao(self) -> float:
        base = self.total_liquido
        if base <= 0:
            return 0.0
        return (self.liquidado + self.pago) / base

    @property
    def esta_anulada(self) -> bool:
        return self.anulado > 0 and (self.a_liquidar + self.em_liquidacao + self.liquidado + self.pago) == 0

    @property
    def esta_pendente(self) -> bool:
        """Ainda tem saldo a liquidar ou em liquidação."""
        return (self.a_liquidar + self.em_liquidacao) > 0


def carregar_notas_empenho(caminho_csv: str | Path, e_restos_a_pagar: bool = False) -> list[NotaEmpenho]:
    """Carrega Corrente_Consolidado.csv OU RP_Adaptado.csv (mesmo layout de
    colunas), filtra por RESP = 10º GAC Sl."""
    notas: list[NotaEmpenho] = []
    for row in ler_csv_dicts(caminho_csv):
        resp = row.get("RESP", "")
        if not eh_10_gac(resp):
            continue
        ne = row.get("NE", "").strip()
        if not ne:
            continue
        notas.append(
            NotaEmpenho(
                ne=ne,
                ug=row.get("UG", "").strip(),
                emissao=row.get("EMISSAO", "").strip(),
                fav=row.get("FAV", "").strip(),
                nome_fav=row.get("NOME_FAV", "").strip(),
                finalidade=row.get("FINALIDADE", "").strip(),
                a_liquidar=parse_valor_brl(row.get("A LIQUIDAR")),
                em_liquidacao=parse_valor_brl(row.get("EM LIQUIDAÇÃO")),
                liquidado=parse_valor_brl(row.get("LIQUIDADO")),
                pago=parse_valor_brl(row.get("PAGO")),
                anulado=parse_valor_brl(row.get("ANULADO")),
                situacao=row.get("SITUAÇÃO", "").strip(),
                e_restos_a_pagar=e_restos_a_pagar,
            )
        )
    return notas


def indicador_pct_liquidacao_global(notas: list[NotaEmpenho]) -> dict:
    total_a_liquidar = sum(n.a_liquidar for n in notas)
    total_em_liquidacao = sum(n.em_liquidacao for n in notas)
    total_liquidado = sum(n.liquidado for n in notas)
    total_pago = sum(n.pago for n in notas)
    total_anulado = sum(n.anulado for n in notas)
    total_bruto = total_a_liquidar + total_em_liquidacao + total_liquidado + total_pago + total_anulado
    total_liquido = total_bruto - total_anulado
    pct = (total_liquidado + total_pago) / total_liquido if total_liquido > 0 else 0.0
    return {
        "qtd_nes": len(notas),
        "total_a_liquidar": total_a_liquidar,
        "total_em_liquidacao": total_em_liquidacao,
        "total_liquidado": total_liquidado,
        "total_pago": total_pago,
        "total_anulado": total_anulado,
        "total_liquido": total_liquido,
        "pct_liquidacao": pct,
    }


def nes_pendentes(notas: list[NotaEmpenho]) -> list[NotaEmpenho]:
    return sorted([n for n in notas if n.esta_pendente and not n.esta_anulada], key=lambda n: n.ne)


def nes_anuladas(notas: list[NotaEmpenho]) -> list[NotaEmpenho]:
    return sorted([n for n in notas if n.esta_anulada], key=lambda n: n.ne)


# ---------------------------------------------------------------------------
# Material Permanente
# ---------------------------------------------------------------------------

@dataclass
class ItemMaterial:
    material: str
    qtd: str
    op: str
    valor_empenhado: float
    ne: str
    nc: str
    chegou: bool
    situacao: str
    destino: str


def carregar_material_permanente(caminho_csv: str | Path) -> list[ItemMaterial]:
    itens = []
    for row in ler_csv_dicts(caminho_csv):
        material = row.get("MATERIAL", "").strip()
        if not material:
            continue
        chegou_raw = row.get("Chegou (Sim/Não)", "").strip().lower()
        itens.append(
            ItemMaterial(
                material=material,
                qtd=row.get("QTD", "").strip(),
                op=row.get("OP", "").strip(),
                valor_empenhado=parse_valor_brl(row.get("Valor Empenhado")),
                ne=row.get("NE", "").strip(),
                nc=row.get(" NC", row.get("NC", "")).strip(),
                chegou=chegou_raw.startswith("sim"),
                situacao=row.get("Situação Atual", "").strip(),
                destino=row.get("DESTINO", "").strip(),
            )
        )
    return itens
