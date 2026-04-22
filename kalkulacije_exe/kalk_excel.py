"""kalk_excel.py — pretvori popis del CSV v formatiran XLSX (sheet BLIST).

Vhod: CSV (;-separated, windows-1250 / utf-8) s strukturo popisa del:
  R0: header (Z.st.;Sifra;Opis;Opomba;EM;Kolicina;Cena;Znesek)
  R1: naslov projekta (samo stolpec C)
  R2+: sekcije (3-stolpcni: A=hierarhija, B=st., C=naziv)
       ali postavke (6-stolpcni: A=hierarhija, B=st., C=opis, D=opomba, E=EM, F=kolicina)

Izhod: XLSX:
  - sheet "BLIST", freeze A3
  - R1: naslov "Ponudbene postavke"
  - R2: header (E=Em, F/G/H = stevilski formati)
  - R3: grand total (zelen fill BEF0BE, C=naslov iz CSV, H=SUM L1)
  - R4+: sekcije in postavke
    - L1: fill C6D7F5 + bold, H=SUM podrejenih
    - L2: fill F0F0F0 + bold, H=SUM podrejenih
    - L3+: bold, H=SUM podrejenih
    - postavke: brez fill/bold, H=ROUND(F*G,2), G prazno (uporabnik izpolni)

Uporaba: kalk_excel.exe <input.csv> <output.xlsx>
Exit codes: 0=OK, 1=napaka (stderr)
"""
import csv
import sys
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ENCODINGS = ("utf-8-sig", "utf-8", "cp1250", "cp1252")

FILL_TOTAL = PatternFill(start_color="FFBEF0BE", end_color="FFBEF0BE", fill_type="solid")
FILL_L1 = PatternFill(start_color="FFC6D7F5", end_color="FFC6D7F5", fill_type="solid")
FILL_L2 = PatternFill(start_color="FFF0F0F0", end_color="FFF0F0F0", fill_type="solid")

FMT_KOLICINA = "#,##0.00##"
FMT_EUR = '#,##0.00\\ "€"'

COL_WIDTHS = {"A": 23.57, "B": 11.43, "C": 57.14, "D": 21.43,
              "E": 6.0, "F": 9.29, "G": 12.86, "H": 14.29}

WRAP = Alignment(wrap_text=True, vertical="top")


@dataclass
class Row:
    kind: str  # "section" | "item"
    level: int  # samo za sekcije, pri postavkah = level leaf-sekcije + 1 (virtualno)
    a: str
    b: str
    c: str
    d: str = ""
    e: str = ""
    f: float | None = None
    xlsx_row: int = 0
    children_rows: list[int] = field(default_factory=list)  # xlsx row indices


def decode(raw: bytes) -> tuple[str, str]:
    for enc in ENCODINGS:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Ne morem dekodirati (poskusil: {ENCODINGS})")


def parse_num(s: str) -> float | None:
    s = s.strip()
    if not s:
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def parse_csv(path: Path) -> tuple[str, list[Row]]:
    raw = path.read_bytes()
    text, enc = decode(raw)
    reader = list(csv.reader(text.splitlines(), delimiter=";", quotechar='"'))
    _ = enc  # encoding odkrit — brez izpisa

    if len(reader) < 2:
        raise ValueError("CSV ima premalo vrstic (vsaj header + naslov)")

    # R1 = naslov projekta (iskani v C stolpcu)
    title = reader[1][2].strip() if len(reader[1]) >= 3 else ""

    rows: list[Row] = []
    for raw_row in reader[2:]:
        if not raw_row or not raw_row[0].strip():
            continue
        a = raw_row[0].strip()
        # Nivo = stevilo pik v A (npr. "1.1.1." ima 3 pike = nivo 3)
        level = a.count(".")
        if len(raw_row) == 3:
            # sekcija
            rows.append(Row(kind="section", level=level, a=a,
                           b=raw_row[1].strip(), c=raw_row[2].strip()))
        elif len(raw_row) >= 5:
            # postavka (6 stolpcev ali vec; E=EM, F=kolicina)
            b = raw_row[1].strip()
            c = raw_row[2].strip()
            d = raw_row[3].strip() if len(raw_row) > 3 else ""
            e = raw_row[4].strip() if len(raw_row) > 4 else ""
            f = parse_num(raw_row[5]) if len(raw_row) > 5 else None
            rows.append(Row(kind="item", level=level + 1, a=a,
                           b=b, c=c, d=d, e=e, f=f))
        # else: ignore malformed

    return title, rows


def assign_xlsx_rows_and_children(rows: list[Row]) -> list[Row]:
    """Dodeli xlsx_row (zacne pri 4) in sestavi children_rows za sekcije.

    Otrok sekcije L: najbljizji vrstici pod njo do naslednje sekcije L ali nizjega.
    - Ce so otroci sekcije nivoja L+1 (subsekcije), list = njihovi xlsx_row.
    - Ce so otroci postavke (leaf sekcija), list = njihovi xlsx_row.
    """
    for i, r in enumerate(rows):
        r.xlsx_row = 4 + i

    # Za vsako sekcijo najdi otroke (direct children).
    for i, r in enumerate(rows):
        if r.kind != "section":
            continue
        children: list[int] = []
        next_child_level: int | None = None
        for j in range(i + 1, len(rows)):
            nxt = rows[j]
            # prekinemo ko pridemo do sekcije na istem ali nizjem nivoju
            if nxt.kind == "section" and nxt.level <= r.level:
                break
            if nxt.kind == "section" and nxt.level == r.level + 1:
                children.append(nxt.xlsx_row)
                if next_child_level is None:
                    next_child_level = nxt.level
            elif nxt.kind == "item" and nxt.level == r.level + 1:
                # item neposredno pod leaf sekcijo
                children.append(nxt.xlsx_row)
                if next_child_level is None:
                    next_child_level = nxt.level
        r.children_rows = children

    return rows


def sum_formula(rows_xlsx: list[int]) -> str:
    if not rows_xlsx:
        return ""
    # Contiguous range ali list?
    rows_xlsx = sorted(rows_xlsx)
    if len(rows_xlsx) > 1 and all(rows_xlsx[i] + 1 == rows_xlsx[i + 1] for i in range(len(rows_xlsx) - 1)):
        return f"=SUM(H{rows_xlsx[0]}:H{rows_xlsx[-1]})"
    # posamezne reference, z vejicami
    refs = ",".join(f"H{r}" for r in rows_xlsx)
    return f"=SUM({refs},)"


def l1_xlsx_rows(rows: list[Row]) -> list[int]:
    return [r.xlsx_row for r in rows if r.kind == "section" and r.level == 1]


def write_xlsx(title: str, rows: list[Row], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "BLIST"

    # Summary row zgoraj (sekcija nad podrobnostmi), ne spodaj.
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False

    # Column widths
    for letter, w in COL_WIDTHS.items():
        ws.column_dimensions[letter].width = w

    # R1: glavni naslov
    ws.cell(row=1, column=1, value="Ponudbene postavke").alignment = WRAP

    # R2: header
    headers = [" Z. št.", "Šifra", "Opis", "Opomba", "Em", "Količina", "Cena", "Znesek"]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c_idx, value=h)
        cell.alignment = WRAP
    ws.cell(row=2, column=6).number_format = FMT_KOLICINA
    ws.cell(row=2, column=7).number_format = FMT_EUR
    ws.cell(row=2, column=8).number_format = FMT_EUR

    # R3: grand total
    total_children = l1_xlsx_rows(rows)
    bold = Font(bold=True)
    for c_idx in range(1, 9):
        cell = ws.cell(row=3, column=c_idx)
        cell.fill = FILL_TOTAL
        cell.font = bold
        cell.alignment = WRAP
    ws.cell(row=3, column=3, value=title)
    ws.cell(row=3, column=6).number_format = FMT_KOLICINA
    ws.cell(row=3, column=7).number_format = FMT_EUR
    ws.cell(row=3, column=8).number_format = FMT_EUR
    if total_children:
        ws.cell(row=3, column=8, value=sum_formula(total_children))

    # R4+: sekcije in postavke
    for r in rows:
        xr = r.xlsx_row

        # napolni osnovne vrednosti
        ws.cell(row=xr, column=1, value=r.a)
        ws.cell(row=xr, column=2, value=r.b)
        ws.cell(row=xr, column=3, value=r.c)
        if r.kind == "item":
            ws.cell(row=xr, column=4, value=r.d)
            ws.cell(row=xr, column=5, value=r.e)
            if r.f is not None:
                ws.cell(row=xr, column=6, value=r.f)
            # G (cena) ostane prazen — uporabnik izpolni
            ws.cell(row=xr, column=8, value=f"=ROUND($F{xr}*G{xr},2)")
        else:  # section
            formula = sum_formula(r.children_rows)
            if formula:
                ws.cell(row=xr, column=8, value=formula)

        # formati + alignment + fill + bold
        is_bold = r.kind == "section"
        if r.kind == "section" and r.level == 1:
            fill = FILL_L1
        elif r.kind == "section" and r.level == 2:
            fill = FILL_L2
        else:
            fill = None

        for c_idx in range(1, 9):
            cell = ws.cell(row=xr, column=c_idx)
            cell.alignment = WRAP
            if is_bold:
                cell.font = bold
            if fill is not None:
                cell.fill = fill
        ws.cell(row=xr, column=6).number_format = FMT_KOLICINA
        ws.cell(row=xr, column=7).number_format = FMT_EUR
        ws.cell(row=xr, column=8).number_format = FMT_EUR

        # Row outline level (za gumbe "1 2 3 4 5 ..." na levi strani):
        # L1 section = 1, L2 = 2, ..., postavka = parent_level + 1 (ze v r.level).
        ws.row_dimensions[xr].outline_level = r.level

    ws.freeze_panes = "A3"
    wb.save(out_path)


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        return 1

    in_path = Path(argv[1])
    out_path = Path(argv[2])

    if not in_path.is_file():
        return 1

    try:
        title, rows = parse_csv(in_path)
    except Exception:
        return 1

    assign_xlsx_rows_and_children(rows)

    try:
        write_xlsx(title, rows, out_path)
    except Exception:
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
