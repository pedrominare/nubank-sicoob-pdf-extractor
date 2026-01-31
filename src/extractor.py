from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

try:
    import pdfplumber
except ModuleNotFoundError as e:  # pragma: no cover
    raise ModuleNotFoundError(
        'Dependência ausente: "pdfplumber". Instale com: pip install pdfplumber'
    ) from e


@dataclass(frozen=True)
class TransactionRow:
    data: str
    descricao: str
    valor: str


@dataclass(frozen=True)
class NubankExtractionResult:
    transacoes: list[TransactionRow]
    pagamentos_e_financiamentos: list[TransactionRow]


class NubankTransactionsExtractor:
    """
    Extrai o bloco "TRANSAÇÕES ..." do titular.
    Retorna linhas normalizadas com: Data (YYYY-MM-DD), Descrição, Valor (string "R$ ...").
    """

    MONTHS = {
        "JAN": 1,
        "FEV": 2,
        "MAR": 3,
        "ABR": 4,
        "MAI": 5,
        "JUN": 6,
        "JUL": 7,
        "AGO": 8,
        "SET": 9,
        "OUT": 10,
        "NOV": 11,
        "DEZ": 12,
    }

    DATE_RE = re.compile(r"^(?P<dd>\d{2})\s+(?P<mon>[A-Z]{3})\b")
    # captura "R$ 1.234,56" e também "−R$ 3,94" / "-R$ 3,94"
    VALUE_RE = re.compile(
        r"(?P<sign>[−-]?)\s*R\$\s*(?P<num>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})"
    )

    def __init__(self, statement_year: int = 2026, holder_name: str = "Nome do Titular"):
        self.statement_year = statement_year
        self.holder_name = holder_name

    def extract(self, pdf_path: Path) -> NubankExtractionResult:
        in_transactions = False
        in_holder_block = False
        in_payments = False

        current: TransactionRow | None = None
        transacoes: list[TransactionRow] = []
        pagamentos: list[TransactionRow] = []

        def flush():
            nonlocal current
            if current:
                if in_payments:
                    pagamentos.append(current)
                else:
                    transacoes.append(current)
                current = None

        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for raw_line in text.splitlines():
                    line = raw_line.strip()
                    if not line:
                        continue

                    # O cabeçalho "TRANSAÇÕES DE ..." se repete a cada página. Se a gente
                    # zerar o bloco do titular aqui, só extrai a primeira página.
                    if line.startswith("TRANSAÇÕES DE "):
                        in_transactions = True
                        flush()
                        continue

                    if not in_transactions:
                        continue

                    # Enquanto não entramos em "Pagamentos e Financiamentos", só começamos a
                    # capturar transações após a linha do titular. Depois que entramos em
                    # "Pagamentos e Financiamentos", capturamos independentemente do titular.
                    if not in_payments and line.startswith(self.holder_name):
                        in_holder_block = True
                        flush()
                        continue
                    if in_payments and line.startswith(self.holder_name):
                        # cabeçalho no topo da página: ignorar para não contaminar descrições
                        continue

                    # entrada do bloco "Pagamentos e Financiamentos"
                    if line.startswith("Pagamentos e Financiamentos"):
                        flush()
                        in_payments = True
                        in_holder_block = False
                        continue
                    if not in_payments and not in_holder_block:
                        continue

                    # ruídos comuns
                    if re.match(r"^\d+\s*de\s*\d+$", line):
                        continue
                    if "FATURA" in line and "EMISSÃO" in line:
                        continue

                    mdate = self.DATE_RE.match(line)
                    if mdate:
                        flush()
                        dd = int(mdate.group("dd"))
                        mon_abbr = mdate.group("mon")
                        if mon_abbr not in self.MONTHS:
                            continue

                        yyyy = self._parse_year(mon_abbr)
                        mm = self.MONTHS[mon_abbr]
                        d = date(yyyy, mm, dd).isoformat()

                        desc = self._clean_desc(line[mdate.end() :])
                        desc = self._strip_trailing_brl(desc)

                        vals = list(self.VALUE_RE.finditer(line))
                        valor = (
                            self._brl_to_str(vals[-1].group("sign"), vals[-1].group("num"))
                            if vals
                            else ""
                        )

                        current = TransactionRow(data=d, descricao=desc, valor=valor)
                        continue

                    # continuação da descrição e/ou valor
                    if current:
                        vals = list(self.VALUE_RE.finditer(line))
                        if vals:
                            current = TransactionRow(
                                data=current.data,
                                descricao=current.descricao,
                                valor=self._brl_to_str(
                                    vals[-1].group("sign"), vals[-1].group("num")
                                ),
                            )
                        else:
                            extra = self._clean_desc(line)
                            if extra:
                                current = TransactionRow(
                                    data=current.data,
                                    descricao=(current.descricao + " " + extra).strip(),
                                    valor=current.valor,
                                )

                # O PDF repete no topo da página o nome completo do titular e outros
                # dados. Sem este flush, essa linha pode ser concatenada na última
                # transação da página anterior.
                flush()

        flush()
        return NubankExtractionResult(
            transacoes=transacoes,
            pagamentos_e_financiamentos=pagamentos,
        )

    def _parse_year(self, mon_abbr: str) -> int:
        # fatura típica: "31 DEZ a 31 JAN" com vencimento/ano 2026 -> DEZ é 2025
        return self.statement_year - 1 if mon_abbr == "DEZ" else self.statement_year

    @staticmethod
    def _brl_to_str(sign: str, num: str) -> str:
        sign = "-" if sign in ("-", "−") else ""
        return f"{sign}R$ {num}"

    @staticmethod
    def _clean_desc(s: str) -> str:
        s = s.strip()
        # remove máscara do cartão "•••• 0539" etc
        s = re.sub(r"^•{4}\s+\d{4}\s+", "", s)
        return re.sub(r"\s+", " ", s).strip()

    @classmethod
    def _strip_trailing_brl(cls, desc: str) -> str:
        """
        Remove um sufixo do tipo "R$ 123,45" (com sinal opcional) do fim da string.
        Isso evita duplicar o valor na coluna Descrição.
        """
        matches = list(cls.VALUE_RE.finditer(desc))
        if not matches:
            return desc

        last = matches[-1]
        if last.end() == len(desc):
            return desc[: last.start()].rstrip()
        return desc

class SicoobCardStatementExtractor:
    """
    Extrai um PDF do Sicoob (Extrato de Cartão de Crédito).

    - transacoes: bloco "GASTOS DE <titular> ..."
    - pagamentos_e_financiamentos: bloco "MOVIMENTOS" (saldo anterior, pagamentos, encargos etc.)
    """

    DATE_DDMM_RE = re.compile(r"^(?P<dd>\d{2})/(?P<mm>\d{2})\b")
    DATE_DDMMYYYY_RE = re.compile(r"\b(?P<dd>\d{2})/(?P<mm>\d{2})/(?P<yyyy>\d{4})\b")
    # último número do tipo 1.234,56 ou -30,00 (sem "R$")
    LAST_VALUE_RE = re.compile(r"(?P<num>-?\d{1,3}(?:\.\d{3})*,\d{2})\s*$")

    def __init__(self, statement_year: int = 2026):
        self.statement_year = statement_year

    def extract(self, pdf_path: Path) -> NubankExtractionResult:
        in_movimentos = False
        in_gastos = False

        current: TransactionRow | None = None
        transacoes: list[TransactionRow] = []
        movimentos: list[TransactionRow] = []

        statement_month = 1
        statement_year = self.statement_year
        valid_months_gastos: set[int] = {1, 12}

        def flush():
            nonlocal current
            if current:
                if in_movimentos and not in_gastos:
                    movimentos.append(current)
                elif in_gastos:
                    transacoes.append(current)
                current = None

        with pdfplumber.open(str(pdf_path)) as pdf:
            # tenta inferir mês/ano da fatura a partir do cabeçalho (ex.: 31/01/2026 ...)
            for page in pdf.pages[:1]:
                header_text = page.extract_text() or ""
                m = self.DATE_DDMMYYYY_RE.search(header_text)
                if m:
                    statement_year = int(m.group("yyyy"))
                    statement_month = int(m.group("mm"))
                    prev_month = 12 if statement_month == 1 else (statement_month - 1)
                    valid_months_gastos = {statement_month, prev_month}
                    break

            for page in pdf.pages:
                text = page.extract_text() or ""
                for raw_line in text.splitlines():
                    line = raw_line.strip()
                    if not line:
                        continue

                    if line.startswith("MOVIMENTOS"):
                        flush()
                        in_movimentos = True
                        in_gastos = False
                        continue

                    if line.startswith("GASTOS DE "):
                        flush()
                        in_gastos = True
                        in_movimentos = False
                        continue

                    # fim do bloco de gastos
                    if in_gastos and (line.startswith("TOTAL ") or line.startswith("DEMONSTRATIVO")):
                        flush()
                        in_gastos = False
                        continue

                    if not in_movimentos and not in_gastos:
                        continue

                    # ignora cabeçalhos/ruído
                    if line == "SICOOB" or "EXTRATO DE CARTÃO DE CRÉDITO" in line:
                        continue
                    if line.startswith("Cliente:") or line.startswith("Fatura de "):
                        continue

                    # linha "SALDO ANTERIOR ..." (sem data)
                    if in_movimentos and line.startswith("- SALDO ANTERIOR"):
                        flush()
                        val = self._parse_last_value_as_brl(line)
                        desc = line.lstrip("-").strip()
                        desc = self._strip_trailing_value(desc)
                        current = TransactionRow(
                            data="",
                            descricao=f"SALDO ANTERIOR: {desc.replace('SALDO ANTERIOR', '').strip()}",
                            valor=val or "",
                        )
                        flush()
                        continue

                    # tratar quebra de linha que começa com "01/02" (parcela) etc
                    # No Sicoob, compras parceladas aparecem como "01/02", "02/03" (parcela),
                    # e quando a descrição quebra, isso pode vir no começo da linha, parecendo
                    # uma data DD/MM. Para evitar criar uma transação falsa, usamos o mês de
                    # referência da fatura:
                    # - em "GASTOS", as datas reais tendem a ficar entre {mês da fatura, mês anterior}
                    # - se vier DD/MM com MM fora desse conjunto e existir um lançamento aberto,
                    #   tratamos como continuação.
                    if current and in_gastos:
                        mstart = self.DATE_DDMM_RE.match(line)
                        if mstart:
                            mm0 = int(mstart.group("mm"))
                            if mm0 not in valid_months_gastos:
                                current = self._merge_continuation(current, line)
                                continue

                    mdate = self.DATE_DDMM_RE.match(line)
                    if mdate:
                        flush()
                        dd = int(mdate.group("dd"))
                        mm = int(mdate.group("mm"))
                        yyyy = statement_year - 1 if mm > statement_month else statement_year
                        d = date(yyyy, mm, dd).isoformat()

                        rest = line[mdate.end() :].strip()
                        valor = (
                            self._parse_last_value_as_brl(rest)
                            or self._parse_last_value_as_brl(line)
                            or ""
                        )
                        desc = self._strip_trailing_value(rest)
                        current = TransactionRow(data=d, descricao=desc, valor=valor)
                        continue

                    # continuação (quebras de linha comuns no PDF)
                    if current:
                        current = self._merge_continuation(current, line)

                # evita concatenar cabeçalho da página seguinte
                flush()

        flush()
        return NubankExtractionResult(
            transacoes=transacoes,
            pagamentos_e_financiamentos=movimentos,
        )

    def _parse_last_value_as_brl(self, s: str) -> str | None:
        m = self.LAST_VALUE_RE.search(s.strip())
        if not m:
            return None
        num = m.group("num")
        sign = "-" if num.startswith("-") else ""
        num_clean = num[1:] if sign else num
        return f"{sign}R$ {num_clean}"

    def _strip_trailing_value(self, s: str) -> str:
        m = self.LAST_VALUE_RE.search(s.strip())
        if not m:
            return s.strip()
        return s[: m.start()].rstrip()

    # helper único para merge de continuação
    def _merge_continuation(self, current: TransactionRow, line: str) -> TransactionRow:
        """
        Junta uma linha de continuação no lançamento atual.

        Regras:
        - Se a linha termina com um valor (ex.: "CANED 69,15"), anexa a parte textual
          na descrição e preenche o valor (somente se ainda estiver vazio).
        - Caso contrário, concatena a linha inteira na descrição.
        """
        # evita concatenar linhas de cabeçalho entre páginas
        if line.startswith("Cliente:") or line.startswith("Conta Cartão:"):
            return current

        mval = self.LAST_VALUE_RE.search(line.strip())
        if mval:
            before = line[: mval.start()].strip()
            updated = current
            if before:
                updated = TransactionRow(
                    data=updated.data,
                    descricao=(updated.descricao + " " + before).strip(),
                    valor=updated.valor,
                )
            if not updated.valor:
                updated = TransactionRow(
                    data=updated.data,
                    descricao=updated.descricao,
                    valor=self._parse_last_value_as_brl(line) or "",
                )
            return updated

        return TransactionRow(
            data=current.data,
            descricao=(current.descricao + " " + line).strip(),
            valor=current.valor,
        )


def detect_bank(pdf_path: Path) -> str:
    """
    Identifica o banco do PDF: 'nubank' ou 'sicoob'.
    """
    with pdfplumber.open(str(pdf_path)) as pdf:
        first = (pdf.pages[0].extract_text() or "").upper()
    if "SICOOB" in first and "EXTRATO DE CARTÃO DE CRÉDITO" in first:
        return "sicoob"
    if "NU PAGAMENTOS" in first or "RESUMO DA FATURA ATUAL" in first:
        return "nubank"
    # fallback
    return "nubank"

class CsvWriter:
    def write_report(self, result: NubankExtractionResult, out_csv_path: Path) -> None:
        out_csv_path.parent.mkdir(parents=True, exist_ok=True)
        with out_csv_path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Data", "Descrição", "Valor"])

            for r in result.transacoes:
                w.writerow([r.data, r.descricao, r.valor])

            for r in result.pagamentos_e_financiamentos:
                # mantém o mesmo formato de 3 colunas, mas marca a origem
                w.writerow([r.data, f"[Pagamentos e Financiamentos] {r.descricao}", r.valor])


class XlsxWriter:
    def write_report(self, result: NubankExtractionResult, out_xlsx_path: Path) -> None:
        try:
            from openpyxl import Workbook
        except ModuleNotFoundError as e:  # pragma: no cover
            raise ModuleNotFoundError(
                'Dependência ausente: "openpyxl". Instale com: pip install openpyxl'
            ) from e

        out_xlsx_path.parent.mkdir(parents=True, exist_ok=True)

        # Observação: Excel/locale podem variar, mas este formato costuma funcionar bem:
        # - moeda "R$"
        # - negativos em parênteses (e vermelhos)
        # - traço para zero
        brl_accounting_format = (
            '_("R$"* #,##0.00_);[Red]_("R$"* (#,##0.00);_("R$"* "-"??_);_(@_)'
        )

        wb = Workbook()
        ws_trans = wb.active
        ws_trans.title = "Transações"

        ws_trans.append(["Data", "Descrição", "Valor"])
        for r in result.transacoes:
            self._append_row_with_accounting(ws_trans, r, brl_accounting_format)

        ws_pay = wb.create_sheet("Pagamentos e Financiamentos")
        ws_pay.append(["Data", "Descrição", "Valor"])
        for r in result.pagamentos_e_financiamentos:
            self._append_row_with_accounting(ws_pay, r, brl_accounting_format)

        wb.save(str(out_xlsx_path))

    @staticmethod
    def _append_row_with_accounting(ws, r: TransactionRow, brl_accounting_format: str) -> None:
        """
        Escreve (Data, Descrição, Valor) onde Valor vira número e recebe formato Contabilidade.
        """
        ws.append([r.data, r.descricao, None])
        cell = ws.cell(row=ws.max_row, column=3)
        cell.value = XlsxWriter._parse_brl_to_number(r.valor)
        cell.number_format = brl_accounting_format

    @staticmethod
    def _parse_brl_to_number(valor: str):
        """
        Converte strings como "R$ 1.234,56" / "-R$ 3,94" em número (float) para Excel.
        Se estiver vazio/indefinido, retorna None.
        """
        if not valor:
            return None

        s = valor.strip()
        sign = -1 if s.startswith("-") or s.startswith("−") else 1
        s = s.lstrip("−-").strip()
        s = s.replace("R$", "").strip()

        # remove separador de milhar "." e troca decimal "," -> "."
        s = s.replace(".", "").replace(",", ".")
        try:
            return sign * float(s)
        except ValueError:
            return None
