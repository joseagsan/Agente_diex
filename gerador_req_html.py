"""
Gera HTML estilizado de uma Requisição de Empenho (estilo NE document).
"""
from __future__ import annotations

_CSS = """
:root {
  --green:#449D44; --green-dark:#3b8c3b; --green-pale:#f0faf0;
  --border:#e0e0e0; --border2:#f5f5f5; --bg:#f5f5f5; --surface:#fff;
  --text:#1a1a2e; --sub:#555; --muted:#999;
  --shadow:0 2px 6px rgba(0,0,0,.06);
}
*{box-sizing:border-box;}
body{background:var(--bg);padding:24px 20px;
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
  font-size:13px;color:var(--text);line-height:1.5;}
.wrapper{max-width:1100px;margin:0 auto;background:var(--surface);
  border-radius:10px;box-shadow:0 2px 12px rgba(0,0,0,.08);overflow:hidden;}

/* Header */
.hdr{background:linear-gradient(135deg,#3b8c3b,#449D44);
  color:#fff;padding:20px 26px;position:relative;overflow:hidden;}
.hdr::after{content:"";position:absolute;right:-60px;top:-60px;
  width:200px;height:200px;border-radius:50%;background:rgba(255,255,255,.07);}
.hdr-tipo{font-size:10px;font-weight:600;letter-spacing:2px;
  text-transform:uppercase;opacity:.75;margin-bottom:4px;}
.hdr-num{font-size:22px;font-weight:800;margin:0 0 3px;}
.hdr-ug{font-size:12px;opacity:.85;}

/* Cards row */
.cards{display:flex;gap:14px;padding:16px;flex-wrap:wrap;}
.card{flex:1;min-width:260px;background:var(--surface);
  border-radius:8px;border:1px solid var(--border);
  border-left:3px solid var(--green);
  box-shadow:var(--shadow);overflow:hidden;}
.card-hdr{padding:9px 14px;font-size:10px;font-weight:700;
  text-transform:uppercase;letter-spacing:.6px;color:var(--sub);
  background:#f8f9fa;border-bottom:1px solid var(--border);
  display:flex;align-items:center;gap:6px;}
.field{padding:7px 14px;border-bottom:1px solid var(--border2);
  display:flex;align-items:baseline;gap:8px;}
.field:last-child{border-bottom:none;}
.lbl{font-size:10px;text-transform:uppercase;color:var(--muted);
  letter-spacing:.4px;font-weight:600;min-width:120px;flex-shrink:0;}
.val{font-size:12.5px;color:var(--text);font-weight:500;}
.val-green{color:var(--green);font-weight:700;font-size:14px;}

/* Obs blocks */
.obs-row{padding:10px 14px;background:var(--green-pale);
  border-top:1px solid #c6e6d2;margin:0 16px 4px;}
.obs-row:last-child{margin-bottom:16px;}
.obs-lbl{font-size:10px;text-transform:uppercase;color:#3b8c3b;
  font-weight:700;letter-spacing:.5px;margin-bottom:3px;}
.obs-val{font-size:12.5px;color:var(--text);line-height:1.55;}

/* Section */
.section{margin:0 16px 14px;border:1px solid var(--border);
  border-radius:8px;overflow:hidden;box-shadow:var(--shadow);}
.sec-hdr{padding:10px 18px;font-size:11px;font-weight:700;
  text-transform:uppercase;letter-spacing:.6px;color:#fff;
  display:flex;justify-content:space-between;align-items:center;}
.sec-hdr.itens{background:linear-gradient(90deg,#3b8c3b,#449D44);}
.sec-hdr.info{background:linear-gradient(90deg,#1e3a5f,#2c5282);}
.sec-hdr.ass{background:linear-gradient(90deg,#064e3b,#065f46);}
.sec-hdr-count{font-weight:400;font-size:11px;opacity:.85;}

/* Table */
table{width:100%;border-collapse:collapse;}
thead th{background:#f8f9fa;padding:8px 10px;font-size:9.5px;
  font-weight:700;text-transform:uppercase;letter-spacing:.4px;
  color:var(--sub);border-bottom:2px solid var(--border);
  text-align:left;white-space:nowrap;}
tbody td{padding:8px 10px;font-size:12px;border-bottom:1px solid var(--border2);}
tbody tr:last-child td{border-bottom:none;}
tbody tr:hover{background:var(--green-pale);}
.td-r{text-align:right;}
.td-val{text-align:right;font-weight:700;color:var(--green);}

/* Total bar */
.total-bar{display:flex;justify-content:flex-end;padding:12px 16px;
  background:#f8f9fa;border-top:2px solid var(--border);
  font-weight:800;font-size:15px;color:var(--green);}
.total-bar span{color:var(--sub);font-weight:400;font-size:12px;margin-right:8px;align-self:center;}
"""

_ICON_ID   = "&#128196;"   # 📄
_ICON_USER = "&#128100;"   # 👤
_ICON_CASH = "&#128178;"   # 💲
_ICON_LIST = "&#128203;"   # 📋


def _field(label: str, value: str) -> str:
    return f'<div class="field"><span class="lbl">{label}</span><span class="val">{value or "—"}</span></div>'


def _field_green(label: str, value: str) -> str:
    return f'<div class="field"><span class="lbl">{label}</span><span class="val val-green">{value or "—"}</span></div>'


def _obs(label: str, value: str) -> str:
    return f'<div class="obs-row"><div class="obs-lbl">{label}</div><div class="obs-val">{value or "—"}</div></div>'


def gerar_html_req(campos: dict, itens: list[dict]) -> str:
    """Gera HTML estilizado da REQ a partir dos campos e itens."""
    req_num   = campos.get("requisition_id", "—")
    ug        = campos.get("UG", "")
    om        = campos.get("OM", "")
    data      = campos.get("LOCAL_DATA", "")
    forn      = campos.get("FORNECEDOR_NOME", "")
    cnpj      = campos.get("FORNECEDOR_CNPJ", "")
    modalidade= campos.get("MODALIDADE", "")
    vigencia  = campos.get("VIGENCIA_DA_ATA", "")
    nc        = campos.get("DADOS_NC", "")
    pi        = campos.get("PI", "")
    nd        = campos.get("ND", "")
    ptres     = campos.get("PTRES", "")
    tipo      = campos.get("TIPO", "Ordinário")
    ne        = campos.get("NE", "")
    assunto   = campos.get("ASSUNTO", "")
    intro     = campos.get("INTRO_1", "")
    just      = campos.get("JUSTIFICATIVA", "")
    finalidade= campos.get("FINALIDADE", "")
    total     = campos.get("TOTAL", "")

    # ── Tabela de itens
    rows_html = ""
    for it in itens:
        rows_html += f"""
        <tr>
          <td>{it.get("ORD","")}</td>
          <td>{it.get("ITEM","")}</td>
          <td>{it.get("SI","")}</td>
          <td style="max-width:340px;white-space:normal">{it.get("DESCRICAO_ITEM","")}</td>
          <td class="td-r">{it.get("UND","")}</td>
          <td class="td-r">{it.get("QTD","")}</td>
          <td class="td-val">{it.get("VALOR_UNIT","")}</td>
          <td class="td-val">{it.get("VALOR_TOTAL","")}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>REQ {req_num}</title>
<style>{_CSS}</style>
</head>
<body>
<div class="wrapper">

  <div class="hdr">
    <div class="hdr-tipo">Requisição de Empenho</div>
    <div class="hdr-num">REQ {req_num}</div>
    <div class="hdr-ug">{ug} — {om} &nbsp;|&nbsp; {data}</div>
  </div>

  <div class="cards">
    <div class="card">
      <div class="card-hdr">{_ICON_ID} Identificação</div>
      {_field("Nº Requisição", req_num)}
      {_field("Tipo", tipo)}
      {_field("Nota de Crédito", nc)}
      {_field("NE Vinculada", ne)}
      {_field("PI", pi)}
      {_field("ND", nd)}
      {_field("PTRES", ptres)}
    </div>
    <div class="card">
      <div class="card-hdr">{_ICON_USER} Fornecedor / Pregão</div>
      {_field("Fornecedor", forn)}
      {_field("CNPJ", cnpj)}
      {_field("Modalidade / Pregão", modalidade)}
      {_field("Vigência da ATA", vigencia)}
    </div>
    <div class="card">
      <div class="card-hdr">{_ICON_CASH} Valor</div>
      {_field_green("Total Geral", total)}
      {_field("Unidade Gestora", ug)}
      {_field("OM", om)}
    </div>
  </div>

  {_obs("Assunto", assunto)}
  {_obs("Introdução", intro)}
  {_obs("Justificativa", just)}
  {_obs("Finalidade / Objeto", finalidade)}

  <div class="section">
    <div class="sec-hdr itens">
      <span>{_ICON_LIST} Lista de Itens</span>
      <span class="sec-hdr-count">{len(itens)} item(ns)</span>
    </div>
    <div style="overflow-x:auto">
      <table>
        <thead>
          <tr>
            <th>Ord</th><th>Item</th><th>SI</th><th>Descrição</th>
            <th class="td-r">Und</th><th class="td-r">Qtd</th>
            <th class="td-r">V. Unit.</th><th class="td-r">V. Total</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
    <div class="total-bar"><span>TOTAL GERAL</span>{total}</div>
  </div>

</div>
</body>
</html>"""
