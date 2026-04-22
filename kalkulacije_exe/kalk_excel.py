"""kalk_excel.py — pretvori CSV v XLSX.

Uporaba:
    kalk_excel.exe <input.csv> <output.xlsx>

Exit codes:
    0 = uspeh
    1 = napaka (detajli na stderr)
"""
import csv
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ENCODINGS = ("utf-8-sig", "utf-8", "cp1250", "cp1252")


def read_csv(path: Path) -> list[list[str]]:
    """Preberi CSV z auto-detect encoding + delimiter."""
    raw = path.read_bytes()
    text = None
    used_enc = None
    for enc in ENCODINGS:
        try:
            text = raw.decode(enc)
            used_enc = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"Ne morem dekodirati {path} (poskusil: {ENCODINGS})")

    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ";"

    reader = csv.reader(text.splitlines(), dialect=dialect)
    rows = [row for row in reader if row]
    print(f"[info] encoding={used_enc} delimiter={dialect.delimiter!r} vrstic={len(rows)}", file=sys.stderr)
    return rows


def try_numeric(value: str):
    """Ce je value stevilka (int/float), vrni numeric; drugace originalen str."""
    v = value.strip()
    if not v:
        return v
    try:
        if "." not in v and "," not in v and "e" not in v.lower():
            return int(v)
    except ValueError:
        pass
    try:
        return float(v.replace(",", "."))
    except ValueError:
        return value


def write_xlsx(rows: list[list[str]], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Podatki"

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, cell in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=try_numeric(cell) if r_idx > 1 else cell)

    if rows:
        header_font = Font(bold=True)
        for c_idx in range(1, len(rows[0]) + 1):
            ws.cell(row=1, column=c_idx).font = header_font
        ws.freeze_panes = "A2"

        for c_idx in range(1, len(rows[0]) + 1):
            max_len = max((len(str(r[c_idx - 1])) for r in rows if c_idx - 1 < len(r)), default=10)
            ws.column_dimensions[get_column_letter(c_idx)].width = min(max(max_len + 2, 10), 50)

    wb.save(out_path)


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print("Uporaba: kalk_excel.exe <input.csv> <output.xlsx>", file=sys.stderr)
        return 1

    in_path = Path(argv[1])
    out_path = Path(argv[2])

    if not in_path.is_file():
        print(f"[err] vhodna datoteka ne obstaja: {in_path}", file=sys.stderr)
        return 1

    try:
        rows = read_csv(in_path)
    except Exception as exc:
        print(f"[err] branje CSV: {exc}", file=sys.stderr)
        return 1

    try:
        write_xlsx(rows, out_path)
    except PermissionError:
        print(f"[err] izhodna datoteka je zaklenjena (Excel odprt?): {out_path}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"[err] pisanje XLSX: {exc}", file=sys.stderr)
        return 1

    print(f"[ok] {in_path} -> {out_path} ({len(rows)} vrstic)", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
