"""
Gestão de Recursos do 10º GAC Sl — cruza as NCs (abas 160482/167482) com os
empenhos (abas "Empenhos 160/167") da planilha de controle de crédito da
1ª Bda Inf Sl e produz a visão consolidada por NC: quanto foi recebido,
recolhido, empenhado, quanto segue em tela, prazo e quais NEs consumiram
cada crédito.

Só regras de negócio — sem Streamlit — para poder ser testado isoladamente.
A página fica em pagina_gestao_recursos.py.

Regras (mesmas da planilha da Bda e de relatorio_gac/core.py):
  - Saldo NC   = Recebido − Recolhido
  - Em tela    = Saldo NC − Empenhado
  - % Execução = Empenhado / Saldo NC   (nunca sobre o recebido bruto)
  - Saldo em tela até RESIDUAL_VALOR_ABS_MAX ou RESIDUAL_PERCENTUAL_MAX do
    recebido é "residual" (sobra de centavos) e não entra nos alertas.
"""
from __future__ import annotations

import io
import re
from collections import defaultdict
from datetime import date, datetime
from typing import Optional

import pandas as pd

RESIDUAL_VALOR_ABS_MAX = 100.0   # R$
RESIDUAL_PERCENTUAL_MAX = 0.5    # % do recebido
DIAS_ALERTA_PRAZO = 7
DIAS_ATENCAO_PRAZO = 30
# NC sem prazo definido e com saldo em tela há mais que isso: crédito parado.
DIAS_CREDITO_PARADO = 30
# Diferença tolerada entre o EMPENHADO da NC e a soma das NEs vinculadas.
TOLERANCIA_CONCILIACAO = 1.0

STATUS_EXECUTADA = "✅ Executada"
STATUS_RESIDUAL  = "⚪ Residual"
STATUS_VENCIDA   = "🔴 Vencida"
STATUS_VENCE_7D  = "🟠 Vence em 7 dias"
STATUS_VENCE_30D = "🟡 Vence em 30 dias"
STATUS_NO_PRAZO  = "🔵 No prazo"
STATUS_SEM_PRAZO = "⚫ Sem prazo"
STATUS_RECOLHER  = "↩️ Recolhimento solicitado"

ORDEM_STATUS = [STATUS_VENCIDA, STATUS_VENCE_7D, STATUS_RECOLHER, STATUS_VENCE_30D,
                STATUS_SEM_PRAZO, STATUS_NO_PRAZO, STATUS_RESIDUAL, STATUS_EXECUTADA]

_RE_NC_BASE = re.compile(r"(\d{4})\s*NC\s*(\d{5,7})", re.IGNORECASE)


# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_moeda(valor) -> float:
    """'R$ 11.464,480' / '-R$ 10,00' / '' → float."""
    s = str(valor or "").replace("R$", "").replace(" ", "").replace("\xa0", "")
    if not s:
        return 0.0
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def fmt_moeda(v: float) -> str:
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def parse_data(s) -> Optional[date]:
    try:
        return datetime.strptime(str(s or "").strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def _bool(v) -> bool:
    return str(v or "").strip().upper() in ("TRUE", "VERDADEIRO", "SIM", "X", "1")


def chave_nc(s) -> str:
    """Normaliza o número da NC preservando o sufixo de lançamento
    ('2026NC400478 2' é outra linha que não '2026NC400478')."""
    return " ".join(str(s or "").upper().split())


def base_nc(s) -> str:
    """Só o número da NC, sem sufixo: '2026NC400478 2' → '2026NC400478'."""
    m = _RE_NC_BASE.search(str(s or ""))
    return f"{m.group(1)}NC{m.group(2)}" if m else chave_nc(s)


def eh_residual(em_tela: float, recebido: float) -> bool:
    if abs(em_tela) <= RESIDUAL_VALOR_ABS_MAX:
        return True
    return recebido > 0 and abs(em_tela) / recebido * 100 <= RESIDUAL_PERCENTUAL_MAX


# ── Consolidação ──────────────────────────────────────────────────────────────
def _status(nc: dict, hoje: date) -> str:
    if nc["em_tela"] <= 0.005:
        return STATUS_EXECUTADA
    if eh_residual(nc["em_tela"], nc["recebido"]):
        return STATUS_RESIDUAL
    d = nc["dias_prazo"]
    if d is not None and d < 0:
        return STATUS_VENCIDA
    if d is not None and d <= DIAS_ALERTA_PRAZO:
        return STATUS_VENCE_7D
    if nc["sol_recolhimento"]:
        return STATUS_RECOLHER
    if d is not None and d <= DIAS_ATENCAO_PRAZO:
        return STATUS_VENCE_30D
    if d is None:
        return STATUS_SEM_PRAZO
    return STATUS_NO_PRAZO


def vincular_empenhos(ncs_chaves: list[str], empenhos: list[dict]) -> tuple[dict[str, list[dict]], list[dict]]:
    """Distribui as NEs pelas NCs. Casa primeiro pelo número exato (com
    sufixo); se não houver, pelo número base quando ele identifica uma única
    NC. Retorna ({chave_nc: [NEs]}, [NEs sem NC correspondente])."""
    por_chave = set(ncs_chaves)
    por_base: dict[str, list[str]] = defaultdict(list)
    for k in ncs_chaves:
        por_base[base_nc(k)].append(k)

    vinculos: dict[str, list[dict]] = defaultdict(list)
    orfas = []
    for e in empenhos:
        k = chave_nc(e.get("NC", ""))
        if k not in por_chave:
            candidatas = por_base.get(base_nc(k), [])
            k = candidatas[0] if len(candidatas) == 1 else ""
        if k:
            vinculos[k].append(e)
        else:
            orfas.append(e)
    return vinculos, orfas


def consolidar(ncs: list[dict], empenhos: list[dict], hoje: Optional[date] = None) -> tuple[list[dict], list[dict]]:
    """Uma linha por NC com valores numéricos, prazo, status e NEs
    vinculadas. Retorna (ncs_consolidadas, empenhos_sem_nc)."""
    hoje = hoje or date.today()
    chaves = [chave_nc(nc.get("NC", "")) for nc in ncs]
    vinculos, orfas = vincular_empenhos(chaves, empenhos)

    resultado = []
    for nc, k in zip(ncs, chaves):
        recebido  = parse_moeda(nc.get("RECEBIDO"))
        recolhido = parse_moeda(nc.get("RECOLHIDO"))
        saldo_nc  = recebido - recolhido
        empenhado = parse_moeda(nc.get("EMPENHADO"))
        em_tela   = saldo_nc - empenhado
        prazo_dt  = parse_data(nc.get("PRAZO"))
        data_nc   = parse_data(nc.get("DATA NC"))
        nes       = vinculos.get(k, [])
        total_nes = sum(parse_moeda(e.get("VALOR NE/NC")) for e in nes)

        linha = {
            "ug":          str(nc.get("UG", "") or nc.get("_ABA_ORIGEM", "")).strip(),
            "nc":          k,
            "orgao":       str(nc.get("ORGÃO", "")).strip(),
            "data_nc":     data_nc,
            "acao":        str(nc.get("AÇÃO", "")).strip(),
            "nd":          str(nc.get("ND", "")).strip(),
            "ptres":       str(nc.get("PTRES", "")).strip(),
            "pi":          str(nc.get("PI", "")).strip(),
            "op":          str(nc.get("OP", "")).strip(),
            "finalidade":  str(nc.get("FINALIDADE", "")).strip(),
            "prazo":       str(nc.get("PRAZO", "")).strip(),
            "prazo_dt":    prazo_dt,
            "dias_prazo":  (prazo_dt - hoje).days if prazo_dt else None,
            "dias_nc":     (hoje - data_nc).days if data_nc else None,
            "recebido":    recebido,
            "recolhido":   recolhido,
            "saldo_nc":    saldo_nc,
            "empenhado":   empenhado,
            "em_tela":     em_tela,
            "pct_exec":    (empenhado / saldo_nc * 100) if saldo_nc > 0 else 0.0,
            "tipo_saldo":  str(nc.get("TIPO SALDO", "")).strip(),
            "sol_recolhimento": _bool(nc.get("SOL RECOL")),
            "sol_justificativa": _bool(nc.get("SOL. JUSTF")),
            "doc_recolhimento": str(nc.get("DOC SOL RCLH/ALT", "")).strip(),
            "nes":         nes,
            "qtd_nes":     len({e.get("NE", "") for e in nes}),
            "total_nes":   total_nes,
            "divergencia": empenhado - total_nes,
        }
        linha["status"] = _status(linha, hoje)
        resultado.append(linha)
    return resultado, orfas


# ── Indicadores ───────────────────────────────────────────────────────────────
def totais(consolidado: list[dict]) -> dict:
    t = {c: sum(n[c] for n in consolidado)
         for c in ("recebido", "recolhido", "saldo_nc", "empenhado", "em_tela")}
    t["qtd_ncs"] = len(consolidado)
    t["pct_exec"] = (t["empenhado"] / t["saldo_nc"] * 100) if t["saldo_nc"] > 0 else 0.0
    abertas = ncs_com_saldo(consolidado)
    t["em_tela_relevante"] = sum(n["em_tela"] for n in abertas)
    t["qtd_com_saldo"] = len(abertas)
    return t


def ncs_com_saldo(consolidado: list[dict]) -> list[dict]:
    """NCs com saldo em tela relevante (não executadas nem residuais),
    ordenadas por urgência e depois por valor."""
    abertas = [n for n in consolidado if n["status"] not in (STATUS_EXECUTADA, STATUS_RESIDUAL)]
    return sorted(abertas, key=lambda n: (ORDEM_STATUS.index(n["status"]),
                                          n["dias_prazo"] if n["dias_prazo"] is not None else 99999,
                                          -n["em_tela"]))


def contagem_status(consolidado: list[dict]) -> list[dict]:
    acc: dict[str, dict] = {}
    for n in consolidado:
        d = acc.setdefault(n["status"], {"status": n["status"], "qtd": 0, "em_tela": 0.0})
        d["qtd"] += 1
        d["em_tela"] += n["em_tela"]
    return sorted(acc.values(), key=lambda d: ORDEM_STATUS.index(d["status"]))


def agrupar(consolidado: list[dict], campo: str) -> pd.DataFrame:
    """Totais por uma dimensão (ug, op, nd, pi, orgao, acao)."""
    acc: dict[str, dict] = {}
    for n in consolidado:
        chave = n.get(campo) or "(vazio)"
        d = acc.setdefault(chave, {campo: chave, "NCs": 0, "recebido": 0.0, "recolhido": 0.0,
                                   "saldo_nc": 0.0, "empenhado": 0.0, "em_tela": 0.0})
        d["NCs"] += 1
        for c in ("recebido", "recolhido", "saldo_nc", "empenhado", "em_tela"):
            d[c] += n[c]
    df = pd.DataFrame(list(acc.values()))
    if df.empty:
        return df
    df["pct_exec"] = df.apply(lambda r: r["empenhado"] / r["saldo_nc"] * 100 if r["saldo_nc"] > 0 else 0.0, axis=1)
    return df.sort_values("em_tela", ascending=False).reset_index(drop=True)


def alertas(consolidado: list[dict], orfas: list[dict]) -> list[dict]:
    """Lista de pendências que pedem ação do SSAC/OD, mais urgente primeiro.
    Cada alerta: {nivel: 'erro'|'aviso'|'info', titulo, detalhe, nc}."""
    out = []
    for n in ncs_com_saldo(consolidado):
        valor = fmt_moeda(n["em_tela"])
        if n["status"] == STATUS_VENCIDA:
            out.append({"nivel": "erro", "nc": n["nc"],
                        "titulo": f"{n['nc']} — prazo vencido há {-n['dias_prazo']} dia(s)",
                        "detalhe": f"{valor} em tela · {n['finalidade']}. Empenhar ou solicitar recolhimento/prorrogação."})
        elif n["status"] == STATUS_VENCE_7D:
            out.append({"nivel": "erro", "nc": n["nc"],
                        "titulo": f"{n['nc']} — vence em {n['dias_prazo']} dia(s) ({n['prazo']})",
                        "detalhe": f"{valor} em tela · {n['finalidade']}."})
        elif n["status"] == STATUS_RECOLHER:
            doc = f" ({n['doc_recolhimento']})" if n["doc_recolhimento"] else ""
            out.append({"nivel": "aviso", "nc": n["nc"],
                        "titulo": f"{n['nc']} — recolhimento solicitado{doc}",
                        "detalhe": f"{valor} em tela aguardando recolhimento · {n['finalidade']}."})
        elif n["status"] == STATUS_SEM_PRAZO and (n["dias_nc"] or 0) > DIAS_CREDITO_PARADO:
            out.append({"nivel": "aviso", "nc": n["nc"],
                        "titulo": f"{n['nc']} — crédito parado há {n['dias_nc']} dia(s), sem prazo",
                        "detalhe": f"{valor} em tela · {n['finalidade']}."})
        if n["sol_justificativa"]:
            out.append({"nivel": "aviso", "nc": n["nc"],
                        "titulo": f"{n['nc']} — Bda solicitou justificativa do saldo",
                        "detalhe": f"{valor} em tela · {n['finalidade']}."})

    for n in consolidado:
        if abs(n["divergencia"]) > TOLERANCIA_CONCILIACAO and (n["empenhado"] or n["total_nes"]):
            out.append({"nivel": "info", "nc": n["nc"],
                        "titulo": f"{n['nc']} — empenhado não confere com as NEs",
                        "detalhe": f"NC: {fmt_moeda(n['empenhado'])} · NEs vinculadas: "
                                   f"{fmt_moeda(n['total_nes'])} · diferença {fmt_moeda(n['divergencia'])}."})
    if orfas:
        total = sum(parse_moeda(e.get("VALOR NE/NC")) for e in orfas)
        out.append({"nivel": "info", "nc": "",
                    "titulo": f"{len(orfas)} lançamento(s) de NE sem NC do GAC correspondente",
                    "detalhe": f"{fmt_moeda(total)} — NC citada na aba de empenhos não aparece nas abas de crédito com RESP do GAC."})
    return out


# ── Tabelas para exibição/exportação ─────────────────────────────────────────
def df_ncs(consolidado: list[dict]) -> pd.DataFrame:
    return pd.DataFrame([{
        "Status":      n["status"],
        "UG":          n["ug"],
        "NC":          n["nc"],
        "Órgão":       n["orgao"],
        "Data NC":     n["data_nc"].strftime("%d/%m/%Y") if n["data_nc"] else "",
        "Ação":        n["acao"],
        "ND":          n["nd"],
        "PTRES":       n["ptres"],
        "PI":          n["pi"],
        "Operação":    n["op"],
        "Finalidade":  n["finalidade"],
        "Prazo":       n["prazo"],
        "Dias p/ prazo": n["dias_prazo"],
        "Recebido":    round(n["recebido"], 2),
        "Recolhido":   round(n["recolhido"], 2),
        "Saldo NC":    round(n["saldo_nc"], 2),
        "Empenhado":   round(n["empenhado"], 2),
        "Em tela":     round(n["em_tela"], 2),
        "% Exec":      round(n["pct_exec"], 1),
        "NEs":         n["qtd_nes"],
        "Soma NEs":    round(n["total_nes"], 2),
        "Divergência": round(n["divergencia"], 2),
        "Doc recolhimento": n["doc_recolhimento"],
    } for n in consolidado])


def df_empenhos(consolidado: list[dict], orfas: list[dict]) -> pd.DataFrame:
    linhas = []
    for n in consolidado:
        for e in n["nes"]:
            linhas.append(_linha_empenho(e, n))
    for e in orfas:
        linhas.append(_linha_empenho(e, None))
    return pd.DataFrame(linhas)


def _linha_empenho(e: dict, nc: Optional[dict]) -> dict:
    return {
        "UG":         str(e.get("UG", "")).strip(),
        "NE":         str(e.get("NE", "")).strip(),
        "Data":       str(e.get("DATA EMP", "")).strip(),
        "NC":         nc["nc"] if nc else chave_nc(e.get("NC", "")) + " (sem vínculo)",
        "ND":         str(e.get("ND", "")).strip(),
        "PI":         str(e.get("PI", "")).strip(),
        "Operação":   str(e.get("OP", "")).strip(),
        "Favorecido": str(e.get("CNPJ/CPF", "")).strip(),
        "Finalidade": str(e.get("FINALIDADE", "")).strip(),
        "Valor":      round(parse_moeda(e.get("VALOR NE/NC")), 2),
        "Saldo NE":   round(parse_moeda(e.get("SALDO")), 2),
    }


def exportar_excel(consolidado: list[dict], orfas: list[dict]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df_ncs(ncs_com_saldo(consolidado)).to_excel(xw, sheet_name="Saldo em tela", index=False)
        df_ncs(consolidado).to_excel(xw, sheet_name="NCs", index=False)
        df_empenhos(consolidado, orfas).to_excel(xw, sheet_name="Empenhos", index=False)
        for campo, nome in (("op", "Por Operação"), ("nd", "Por ND"), ("pi", "Por PI"), ("orgao", "Por Órgão")):
            agrupar(consolidado, campo).to_excel(xw, sheet_name=nome, index=False)
        pd.DataFrame(alertas(consolidado, orfas)).to_excel(xw, sheet_name="Alertas", index=False)
    return buf.getvalue()
