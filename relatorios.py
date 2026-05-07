"""
Geração de KPIs, relatórios e exportações.
"""
import io
from datetime import date, datetime
from typing import Optional

import pandas as pd


def _parse_moeda(s) -> float:
    try:
        return float(str(s).replace("R$", "").replace(" ", "").replace(".", "").replace(",", "."))
    except (ValueError, TypeError):
        return 0.0


def _fmt_moeda(v: float) -> str:
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def _parse_data(s: str) -> Optional[date]:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except (ValueError, TypeError):
            pass
    return None


def kpis(ncs: list[dict], reqs: list[dict]) -> dict:
    """Calcula e retorna os indicadores principais."""
    saldo_total = sum(_parse_moeda(nc.get("SALDO NC", 0)) for nc in ncs)
    empenhado_total = sum(_parse_moeda(nc.get("EMPENHADO", 0)) for nc in ncs)
    em_tela_total = sum(_parse_moeda(nc.get("EM TELA", 0)) for nc in ncs)
    recebido_total = sum(_parse_moeda(nc.get("RECEBIDO", 0)) for nc in ncs)

    hoje = date.today()

    def _dias(prazo_str):
        d = _parse_data(prazo_str)
        return (d - hoje).days if d else None

    vencendo_7d = sum(
        1 for nc in ncs
        if nc.get("SITU") == "OK" and _dias(nc.get("PRAZO", "")) is not None
        and 0 <= _dias(nc.get("PRAZO", "")) <= 7
    )
    vencidas = sum(
        1 for nc in ncs
        if nc.get("SITU") == "OK" and _dias(nc.get("PRAZO", "")) is not None
        and _dias(nc.get("PRAZO", "")) < 0
    )

    reqs_pendentes = sum(1 for r in reqs if r.get("SITUAÇÃO") == "Pendente")
    valor_reqs_pendentes = sum(_parse_moeda(r.get("VALOR", 0)) for r in reqs if r.get("SITUAÇÃO") == "Pendente")

    return {
        "total_ncs": len(ncs),
        "ncs_ok": sum(1 for nc in ncs if nc.get("SITU") == "OK"),
        "ncs_em_tela": sum(1 for nc in ncs if nc.get("SITU") == "EM TELA"),
        "saldo_total": saldo_total,
        "empenhado_total": empenhado_total,
        "em_tela_total": em_tela_total,
        "recebido_total": recebido_total,
        "vencendo_7d": vencendo_7d,
        "vencidas": vencidas,
        "reqs_total": len(reqs),
        "reqs_pendentes": reqs_pendentes,
        "valor_reqs_pendentes": valor_reqs_pendentes,
    }


def df_ncs_enriquecido(ncs: list[dict]) -> pd.DataFrame:
    """DataFrame de NCs com colunas numéricas e datas parseadas."""
    df = pd.DataFrame(ncs)
    if df.empty:
        return df

    for col in ["RECEBIDO", "RECOLHIDO", "SALDO NC", "EMPENHADO", "EM TELA"]:
        if col in df.columns:
            df[f"_{col}_num"] = df[col].apply(_parse_moeda)

    for col in ["DATA NC", "PRAZO"]:
        if col in df.columns:
            df[f"_{col}_dt"] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")

    return df


def df_reqs_enriquecido(reqs: list[dict]) -> pd.DataFrame:
    """DataFrame de REQs com colunas numéricas e datas parseadas."""
    df = pd.DataFrame(reqs)
    if df.empty:
        return df

    if "VALOR" in df.columns:
        df["_VALOR_num"] = df["VALOR"].apply(_parse_moeda)
    if "DATA REQ" in df.columns:
        df["_DATA_REQ_dt"] = pd.to_datetime(df["DATA REQ"], format="%d/%m/%Y", errors="coerce")

    return df


def ncs_por_operacao(ncs: list[dict]) -> pd.DataFrame:
    rows = [
        {
            "Operação": nc.get("OP") or "N/A",
            "Qtd NCs": 1,
            "Recebido": _parse_moeda(nc.get("RECEBIDO", 0)),
            "Saldo": _parse_moeda(nc.get("SALDO NC", 0)),
            "Empenhado": _parse_moeda(nc.get("EMPENHADO", 0)),
        }
        for nc in ncs
    ]
    df = pd.DataFrame(rows).groupby("Operação").agg(
        {"Qtd NCs": "sum", "Recebido": "sum", "Saldo": "sum", "Empenhado": "sum"}
    ).reset_index().sort_values("Saldo", ascending=False)
    for col in ["Recebido", "Saldo", "Empenhado"]:
        df[col] = df[col].apply(_fmt_moeda)
    return df


def ncs_vencendo(ncs: list[dict], dias: int = 30) -> list[dict]:
    hoje = date.today()
    resultado = []
    for nc in ncs:
        d = _parse_data(nc.get("PRAZO", ""))
        if d:
            restantes = (d - hoje).days
            if restantes <= dias:
                resultado.append({**nc, "_DIAS_RESTANTES": restantes})
    return sorted(resultado, key=lambda x: x["_DIAS_RESTANTES"])


def exportar_excel(ncs: list[dict], reqs: list[dict]) -> bytes:
    """Gera planilha Excel com abas NCs, Requisições e Resumo."""
    output = io.BytesIO()
    ind = kpis(ncs, reqs)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if ncs:
            pd.DataFrame(ncs).to_excel(writer, sheet_name="NCs", index=False)
        if reqs:
            pd.DataFrame(reqs).to_excel(writer, sheet_name="Requisições", index=False)

        resumo_rows = [
            ("Data do Relatório", datetime.today().strftime("%d/%m/%Y %H:%M")),
            ("Total de NCs", ind["total_ncs"]),
            ("NCs OK", ind["ncs_ok"]),
            ("NCs EM TELA", ind["ncs_em_tela"]),
            ("Valor Total Recebido", _fmt_moeda(ind["recebido_total"])),
            ("Saldo Total Disponível", _fmt_moeda(ind["saldo_total"])),
            ("Total Empenhado", _fmt_moeda(ind["empenhado_total"])),
            ("Total EM TELA", _fmt_moeda(ind["em_tela_total"])),
            ("NCs Vencendo em 7 dias", ind["vencendo_7d"]),
            ("NCs Vencidas", ind["vencidas"]),
            ("Total de REQs", ind["reqs_total"]),
            ("REQs Pendentes", ind["reqs_pendentes"]),
            ("Valor REQs Pendentes", _fmt_moeda(ind["valor_reqs_pendentes"])),
        ]
        pd.DataFrame(resumo_rows, columns=["Indicador", "Valor"]).to_excel(
            writer, sheet_name="Resumo", index=False
        )

        if ncs:
            ncs_por_operacao(ncs).to_excel(writer, sheet_name="Por Operação", index=False)

    return output.getvalue()
