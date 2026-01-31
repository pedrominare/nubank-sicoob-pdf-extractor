import argparse
from pathlib import Path

from src.extractor import (
    CsvWriter,
    NubankTransactionsExtractor,
    SicoobCardStatementExtractor,
    XlsxWriter,
    detect_bank,
)


def main():
    parser = argparse.ArgumentParser(
        description='Extrai transações do bloco "Nome do Titular" e gera CSV/XLSX.'
    )
    
    parser.add_argument("pdf", nargs="?", help="Caminho do arquivo PDF (fatura Nubank)")
    parser.add_argument(
        "-o",
        "--out",
        default=None,
        help="Caminho do arquivo de saída (default: mesmo nome do PDF com extensão do formato)",
    )
    
    parser.add_argument(
        "--format",
        choices=["csv", "xlsx"],
        default="csv",
        help="Formato de saída (default: csv)",
    )

    parser.add_argument(
        "--bank",
        choices=["auto", "nubank", "sicoob"],
        default="auto",
        help="Banco do PDF (default: auto)",
    )
    
    parser.add_argument(
        "--year",
        type=int,
        default=2026,
        help="Ano da fatura (default: 2026). Usado para resolver DEZ do ano anterior.",
    )

    parser.add_argument(
        "--owner",
        default=None,
        help="Nome do titular do cartao",
    )
    args = parser.parse_args()

    if not args.pdf:
        args.pdf = input("Informe o caminho do PDF: ").strip().strip('"')
        if not args.pdf:
            raise SystemExit("PDF não informado.")
    

    pdf_path = Path(args.pdf).expanduser()
    if not pdf_path.exists():
        raise SystemExit(f"Arquivo não encontrado: {pdf_path}")

    out_path = Path(args.out) if args.out else pdf_path.with_suffix(f".{args.format}")

    bank = args.bank if args.bank != "auto" else detect_bank(pdf_path)
    if bank == "sicoob":
        extractor = SicoobCardStatementExtractor(statement_year=args.year)
    else:
        extractor = NubankTransactionsExtractor(statement_year=args.year, holder_name=args.owner)
    result = extractor.extract(pdf_path)

    if args.format == "xlsx":
        XlsxWriter().write_report(result, out_path)
    else:
        CsvWriter().write_report(result, out_path)
    total = len(result.transacoes) + len(result.pagamentos_e_financiamentos)
    print(f"OK: {total} linhas salvas em {out_path}")

if __name__ == "__main__":
    main()
