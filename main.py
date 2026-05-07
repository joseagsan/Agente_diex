"""
Agente DIEx — Gerador de Documentos Administrativos Militares
Uso:
    python main.py                          # processa todos os IDs da planilha
    python main.py --id REQ-001             # processa apenas o ID especificado
    python main.py --local dados.xlsx       # usa arquivo local em vez do Sheets
    python main.py --template outro.docx    # usa template alternativo
"""
import argparse
import logging
import os
import sys

from config import TEMPLATE_PATH, OUTPUT_DIR

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


def _carregar_dados(args) -> list[dict]:
    if args.local:
        try:
            import pandas as pd
        except ImportError:
            logger.error("pandas não instalado. Execute: pip install pandas openpyxl")
            sys.exit(1)
        ext = os.path.splitext(args.local)[1].lower()
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(args.local, dtype=str)
        elif ext == ".csv":
            df = pd.read_csv(args.local, dtype=str)
        else:
            logger.error("Formato não suportado: %s. Use .xlsx, .xls ou .csv.", ext)
            sys.exit(1)
        df = df.fillna("")
        logger.info("Arquivo local lido: %s (%d linha(s))", args.local, len(df))
        return df.to_dict("records")
    else:
        from sheets import ler_dados
        return ler_dados()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Gera documentos DIEx/Requisição em DOCX a partir de planilha.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--id",
        metavar="REQUISITION_ID",
        help="Processa apenas o ID especificado (ex.: REQ-001).",
    )
    parser.add_argument(
        "--local",
        metavar="ARQUIVO",
        help="Usa arquivo local (.xlsx/.csv) em vez do Google Sheets.",
    )
    parser.add_argument(
        "--template",
        default=TEMPLATE_PATH,
        metavar="ARQUIVO.docx",
        help=f"Template DOCX com placeholders (padrão: {TEMPLATE_PATH}).",
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        metavar="DIRETÓRIO",
        help=f"Diretório de saída dos documentos gerados (padrão: {OUTPUT_DIR}).",
    )
    args = parser.parse_args()

    logger.info("=" * 55)
    logger.info("  AGENTE DIEx — Gerador de Documentos")
    logger.info("=" * 55)

    # --- Leitura ---
    dados_brutos = _carregar_dados(args)
    if not dados_brutos:
        logger.error("Nenhum dado encontrado. Verifique a planilha ou arquivo local.")
        sys.exit(1)

    # --- Processamento ---
    from processor import processar
    requisicoes = processar(dados_brutos)

    if not requisicoes:
        logger.error("Nenhuma requisição válida encontrada (verifique a coluna requisition_id).")
        sys.exit(1)

    # --- Filtro por ID (opcional) ---
    if args.id:
        if args.id not in requisicoes:
            ids_disponiveis = ", ".join(sorted(requisicoes.keys()))
            logger.error(
                "ID '%s' não encontrado. IDs disponíveis: %s",
                args.id,
                ids_disponiveis,
            )
            sys.exit(1)
        requisicoes = {args.id: requisicoes[args.id]}

    logger.info("%d requisição(ões) a processar.", len(requisicoes))

    # --- Geração ---
    from gerador import gerar_documento

    gerados = 0
    erros = 0

    for req_id, dados_req in sorted(requisicoes.items()):
        campos = dados_req["campos"]
        itens = dados_req["itens"]
        nd = campos.get("ND", "SEM_ND")
        nome_arquivo = f"REQ_{req_id}_{nd}.docx"
        output_path = os.path.join(args.output_dir, nome_arquivo)

        logger.info(
            "Gerando %-30s | %2d item(ns) | ND: %s",
            nome_arquivo,
            len(itens),
            nd,
        )
        try:
            gerar_documento(args.template, output_path, campos, itens)
            gerados += 1
        except FileNotFoundError as exc:
            logger.error("Template não encontrado: %s", exc)
            logger.error(
                "Coloque o arquivo em '%s' ou use --template para indicar outro caminho.",
                args.template,
            )
            sys.exit(1)
        except Exception as exc:
            logger.error("Erro ao gerar '%s': %s", nome_arquivo, exc, exc_info=True)
            erros += 1

    logger.info("=" * 55)
    logger.info(
        "  Concluído: %d gerado(s), %d erro(s).",
        gerados,
        erros,
    )
    logger.info("  Saída: %s", os.path.abspath(args.output_dir))
    logger.info("=" * 55)

    if erros:
        sys.exit(1)


if __name__ == "__main__":
    main()
