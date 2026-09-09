"""
Gera HTML estilizado de minutas de despacho e do ateste de Nota Fiscal —
Fase 2 (pós-empenho) do requisitante, conforme fluxograma SPED 3.0
(itens 10.7 "Ateste da NF" e 10.9 "Minutas de despacho do Cmt/Fisc Adm/OD").

O texto gerado é um rascunho pronto para copiar/colar no SPED — a
assinatura e a tramitação continuam sendo feitas lá.
"""
from __future__ import annotations

from datetime import date

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
.wrapper{max-width:820px;margin:0 auto;background:var(--surface);
  border-radius:10px;box-shadow:0 2px 12px rgba(0,0,0,.08);overflow:hidden;}
.hdr{background:linear-gradient(135deg,#3b8c3b,#449D44);
  color:#fff;padding:20px 26px;position:relative;overflow:hidden;}
.hdr::after{content:"";position:absolute;right:-60px;top:-60px;
  width:200px;height:200px;border-radius:50%;background:rgba(255,255,255,.07);}
.hdr-tipo{font-size:10px;font-weight:600;letter-spacing:2px;
  text-transform:uppercase;opacity:.75;margin-bottom:4px;}
.hdr-num{font-size:20px;font-weight:800;margin:0 0 3px;}
.hdr-ug{font-size:12px;opacity:.85;}
.meta{display:flex;flex-wrap:wrap;gap:14px;padding:14px 26px;
  background:#f8f9fa;border-bottom:1px solid var(--border);font-size:11.5px;color:var(--sub);}
.meta b{color:var(--text);}
.body{padding:24px 30px;}
.body p{margin:0 0 14px;text-align:justify;white-space:pre-wrap;}
.sign{margin-top:36px;text-align:center;}
.sign .line{border-top:1px solid var(--text);width:280px;margin:0 auto 4px;}
.sign .cargo{font-size:11.5px;color:var(--sub);}
"""


def _wrap(titulo: str, req_num: str, meta_html: str, corpo_html: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{titulo} — REQ {req_num}</title>
<style>{_CSS}</style>
</head>
<body>
<div class="wrapper">
  <div class="hdr">
    <div class="hdr-tipo">Minuta de Despacho</div>
    <div class="hdr-num">{titulo}</div>
    <div class="hdr-ug">REQ {req_num}</div>
  </div>
  <div class="meta">{meta_html}</div>
  <div class="body">{corpo_html}</div>
</div>
</body>
</html>"""


def _meta(campos: dict, extras: list[tuple[str, str]] | None = None) -> str:
    itens = [
        ("OM", campos.get("OM", "")),
        ("UG", campos.get("UG", "")),
        ("Data", campos.get("DATA", date.today().strftime("%d/%m/%Y"))),
        ("NC", campos.get("NC", "")),
        ("NE", campos.get("NE", "")),
    ]
    if extras:
        itens += extras
    return "".join(f"<span><b>{lbl}:</b> {val or '—'}</span>" for lbl, val in itens)


TIPOS_DESPACHO = {
    "CMT":      "Despacho do Comandante",
    "FISC_ADM": "Despacho do Fiscal Administrativo",
    "OD":       "Despacho do Ordenador de Despesas",
}


def texto_padrao_despacho(tipo: str, campos: dict) -> str:
    """Sugestão inicial de texto — o usuário edita antes de gerar/copiar."""
    req = campos.get("REQ", "")
    empresa = campos.get("EMPRESA", "")
    valor = campos.get("VALOR", "")
    if tipo == "CMT":
        return (
            f"Aprovo o presente processo referente à Requisição nº {req}, "
            f"junto à empresa {empresa}, no valor de {valor}, e encaminho para "
            f"apreciação do Fiscal Administrativo da Brigada."
        )
    if tipo == "FISC_ADM":
        return (
            f"Em análise ao processo referente à Requisição nº {req}, verifico que "
            f"a documentação encontra-se em conformidade com a legislação vigente. "
            f"Encaminho ao Ordenador de Despesas para prosseguimento."
        )
    if tipo == "OD":
        return (
            f"Autorizo o prosseguimento do processo referente à Requisição nº {req}, "
            f"no valor de {valor}, junto à empresa {empresa}."
        )
    return ""


def gerar_html_despacho(tipo: str, campos: dict) -> str:
    """tipo: 'CMT' | 'FISC_ADM' | 'OD'. campos deve conter REQ, OM, UG, DATA,
    NC, NE, EMPRESA, VALOR, TEXTO, SIGNATARIO (opcional)."""
    titulo = TIPOS_DESPACHO.get(tipo, "Despacho")
    req_num = campos.get("REQ", "—")
    texto = (campos.get("TEXTO", "") or "").strip()
    signatario = campos.get("SIGNATARIO", "")

    corpo = f"<p>{texto or '(texto do despacho)'}</p>"
    if signatario or True:
        corpo += (
            '<div class="sign"><div class="line"></div>'
            f'<div class="cargo">{signatario or titulo}</div></div>'
        )
    return _wrap(titulo, req_num, _meta(campos), corpo)


def texto_padrao_ateste(campos: dict) -> str:
    """Sugestão inicial de texto do ateste — o usuário edita antes de gerar/copiar."""
    req_num = campos.get("REQ", "—")
    num_nf = campos.get("NUM_NF", "")
    data_nf = campos.get("DATA_NF", "")
    tipo_nf = campos.get("TIPO_NF", "Material")
    empresa = campos.get("EMPRESA", "")
    valor = campos.get("VALOR", "")
    verbo = "recebido" if tipo_nf == "Material" else "prestado a contento"
    return (
        f"Atesto, para os devidos fins, que o {'material' if tipo_nf == 'Material' else 'serviço'} "
        f"referente à Nota Fiscal nº {num_nf or '____'}, emitida por {empresa}, datada de "
        f"{data_nf or '__/__/____'}, no valor de {valor}, foi {verbo}, encontrando-se em "
        f"conformidade com o empenhado na Requisição nº {req_num}."
    )


def gerar_html_ateste(campos: dict) -> str:
    """campos: REQ, OM, UG, DATA, NC, NE, EMPRESA, NUM_NF, DATA_NF, TIPO_NF,
    VALOR, TEXTO (opcional, sobrescreve o texto padrão), SIGNATARIO."""
    req_num = campos.get("REQ", "—")
    num_nf = campos.get("NUM_NF", "")
    data_nf = campos.get("DATA_NF", "")
    tipo_nf = campos.get("TIPO_NF", "Material")
    signatario = campos.get("SIGNATARIO", "")

    texto = (campos.get("TEXTO", "") or "").strip() or texto_padrao_ateste(campos)

    corpo = f"<p>{texto}</p>"
    corpo += (
        '<div class="sign"><div class="line"></div>'
        f'<div class="cargo">{signatario or "Requisitante"}</div></div>'
    )
    meta = _meta(campos, extras=[("NF", num_nf), ("Data NF", data_nf), ("Tipo", tipo_nf)])
    return _wrap("Ateste da Nota Fiscal", req_num, meta, corpo)
