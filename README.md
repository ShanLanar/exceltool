# exceltool

Schlanke **Tkinter-GUI-Toolbox** (Hauptdatei `exceltool.py`) mit fünf Werkzeugen
rund um Tabellen- und B2B-Datenformate.

## Zweck

| Tab | Funktion |
|-----|----------|
| 📂 **CSV/XLS/XLSX → Excel** | Mehrere Dateien zusammenführen – jede Datei/jedes Blatt wird ein eigener Reiter |
| 🔓 **Blattschutz entfernen** | Sheet-/Arbeitsmappenschutz aus `.xlsx`/`.xls` entfernen (Stapellauf) |
| 📄 **OpenTrans / ORDERS05 / EDIFACT** | Bestell-Dokumente parsen und einheitlich nach Excel exportieren (Header + Positionen) |
| ✂ **CSV-Splitter** | Große CSVs in gleich große Blöcke teilen (jeweils mit Kopfzeile) |
| 🔄 **XML → CSV** | `<PARAMETER DISPLAYNAME="…">`-Fragmente als CSV exportieren |

## Installation

```bash
python -m venv .venv
# Linux/macOS:  source .venv/bin/activate
# Windows:      .venv\Scripts\activate
pip install -e ".[dev]"
```

Installiert die Laufzeit-Abhängigkeiten (`pandas`, `openpyxl`, `xlrd`) sowie
`pytest`. `tkinter` gehört zur Python-Standardbibliothek; unter Linux ggf.
`sudo apt-get install python3-tk` nachinstallieren.

## Start

```bash
python exceltool.py
```

### Windows – Doppelklick-Launcher

Für Windows liegen Batch-Dateien bei (kein manuelles venv/CLI nötig):

| Datei | Zweck |
|-------|-------|
| `start.bat` | legt beim ersten Lauf `.venv` an, installiert aus `pyproject.toml` und startet das Programm (Neuinstallation nur, wenn sich `pyproject.toml` geändert hat) |
| `update.bat` | `git pull` |
| `update-and-run.bat` | `git pull` und anschließend `start.bat` |

## Entwicklung / Tests

```bash
pytest                       # bzw.  python test_exceltool.py
```

Die Tests laufen **headless**: `tkinter` wird nur gestubbt, wenn nicht vorhanden,
und der GUI-Aufbau steckt komplett in `main()` (läuft beim Import nicht an).
Reine Logik (Parser, Konverter, Splitter, Blattschutz) ist damit ohne Display
testbar. CI prüft per `py_compile` die Syntax bei jedem Push/PR.
