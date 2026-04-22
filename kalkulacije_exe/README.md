# kalkulacije_exe

Sourci za **Kalkulacije** (Magic dodatki, EXE) — seja 902.

## Aktivni EXE-ji

### `kalk_excel.exe` — CSV -> XLSX pretvornik

- **Stack**: Python 3 + openpyxl + PyInstaller (onefile, self-contained)
- **Source**: [kalk_excel.py](kalk_excel.py)
- **Build**: [build.ps1](build.ps1) -> `c:/ai_exe_dll/kalkulacije_exe/kalk_excel.exe`
- **Uporaba**: `kalk_excel.exe <input.csv> <output.xlsx>`
- **Features**:
  - Auto-detect encoding (UTF-8 BOM / UTF-8 / cp1250 / cp1252)
  - Auto-detect delimiter (`,` `;` TAB `|`), fallback `;`
  - Prva vrstica = header (bold + frozen pane v Excelu)
  - Numericne celice v podatkih se shranijo kot stevila (int/float), decimal tako `.` kot `,`
  - Auto-sirina stolpcev (10-50 znakov)
- **Exit codes**: 0 = OK, 1 = napaka (detajli na stderr)

## Poti

- **Source**: `c:/Projekti/magic_exe_dll/kalkulacije_exe/` (pod git-om)
- **Build output**: `c:/ai_exe_dll/kalkulacije_exe/` (ni pod git-om)

## Build setup (enkrat)

```powershell
python -m pip install --user openpyxl pyinstaller
```

Potem kadarkoli:

```powershell
cd c:\Projekti\magic_exe_dll\kalkulacije_exe
.\build.ps1
```
