# magic_exe_dll

Sourci za Magic dodatke (EXE + DLL) — seja Claude 902.

## Layout

- **`c:/Projekti/magic_exe_dll/`** — ta repo, samo sourci (source tree).
- **`c:/ai_exe_dll/`** — buildani artefakti (EXE/DLL + NAVODILA). **Ni** pod git-om tega repoja.

**En program = en podfolder.** Pred vsakim novim programom se uporabnik in Claude dogovorita za ime (glej pravilo 1a v [.ai/902/CLAUDE.md](../.ai/902/CLAUDE.md)).

## Register programov (aktivni)

| Podfolder | EXE | Namen | Stack | Status | Zacetek |
|---|---|---|---|---|---|
| [kalkulacije_exe](kalkulacije_exe/) | `kalk_excel.exe` | Pretvori "popis del" CSV (Slovenska gradbena ponudba, cp1250, `;`-locen, 26-stolpcni) v formatiran XLSX (sheet BLIST, hierarhicne sekcije z color-coded nivoji, outline grupe, SUM/ROUND formule, hidden Z stolpec z alfa kodo) | Python 3 + openpyxl + PyInstaller (onefile, ~20 MB, tih) | v produkciji | 2026-04-22 |

### Kako brati register

- **Podfolder** — source pot znotraj tega repoja; klik vodi do `README.md` podfoldra
- **EXE** — ime build outputa; dejanska lokacija je `c:/ai_exe_dll/<podfolder>/<exe>`
- **Namen** — kaj program dela, kaj so vhodi/izhodi
- **Stack** — jezik in ključne knjižnice
- **Status** — `v pripravi (specifikacija)` / `v razvoju` / `v produkciji` / `arhiv`
- **Zacetek** — datum kreacije podfoldra

### Kako dodati nov program

1. Dogovori se z uporabnikom za ime podfoldra (pravilo 1a).
2. `c:/Projekti/magic_exe_dll/<ime>/` — source + lasten README.md + build skripta
3. `c:/ai_exe_dll/<ime>/` — build output target (dodaj v `.gitignore`? Ne — cel c:/ai_exe_dll/ je ze izven git-a)
4. Dopisi vrstico v tabelo zgoraj.
5. Posodobi [c:/ai_exe_dll/INDEX.txt](file:///C:/ai_exe_dll/INDEX.txt) (flat register na build strani).

## Pravila

Glej [.ai/902/CLAUDE.md](../.ai/902/CLAUDE.md) za popolna pravila seje 902 (folder layout, SB_WEB off-limits, git/GitHub vzorec, Codex komunikacija, …).
