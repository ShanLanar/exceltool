# exceltool

Eine schlanke **Tkinter-GUI-Toolbox** (eine Datei: `exceltool.py`) mit fünf
Werkzeugen rund um Tabellen- und B2B-Datenformate:

| Tab | Funktion |
|-----|----------|
| 📂 **CSV/XLS/XLSX → Excel** | Mehrere Dateien zusammenführen – jede Datei/jedes Blatt wird ein eigener Reiter |
| 🔓 **Blattschutz entfernen** | Sheet-/Arbeitsmappenschutz aus `.xlsx`/`.xls` entfernen (Stapellauf) |
| 📄 **OpenTrans / ORDERS05 / EDIFACT** | Bestell-Dokumente parsen und einheitlich nach Excel exportieren (Header + Positionen) |
| ✂ **CSV-Splitter** | Große CSVs in gleich große Blöcke teilen (jeweils mit Kopfzeile) |
| 🔄 **XML → CSV** | `<PARAMETER DISPLAYNAME="…">`-Fragmente als CSV exportieren |

## Installation & Start

```bash
pip install -r requirements.txt      # pandas, openpyxl, xlrd
python exceltool.py
```

`tkinter` gehört zur Python-Standardbibliothek; unter Linux ggf. nachinstallieren:

```bash
sudo apt-get install python3-tk
```

Alternativ als Paket (stellt den Befehl `exceltool` bereit):

```bash
pip install .
exceltool
```

## Eigenschaften

- **Nicht-blockierendes UI:** Schwere Aufgaben laufen in Hintergrund-Threads;
  die Oberfläche bleibt bedienbar. UI-Updates werden thread-sicher über eine
  Queue in den Main-Thread zurückgereicht (Tkinter ist nicht thread-safe).
- **Namespace-sicheres Entschützen:** Der Blattschutz wird rein textuell aus der
  OOXML-Struktur entfernt – kein XML-Neuaufbau, dadurch keine beschädigten Dateien.
- **Robuste Eingaben:** Encoding-Fallback (`utf-8-sig` → `utf-8` → `iso-8859-1`
  → `cp1252`), automatische Trennzeichen-Erkennung.
- **Sicherer Export:** Formel-/CSV-Injection wird beim Excel-Export entschärft.

## Entwicklung

```bash
pip install -e ".[dev]"     # + pytest, ruff
pytest                      # bzw.  python test_exceltool.py
ruff check .                # Lint
```

Die Tests laufen **headless**: `tkinter` wird nur gestubbt, wenn nicht vorhanden,
und der GUI-Aufbau steckt komplett in `main()` (läuft beim Import nicht an).
Reine Logik (Parser, Konverter, Splitter, Blattschutz) ist damit ohne Display
testbar. CI (GitHub Actions) prüft Lint + Tests auf Python 3.11 und 3.12.

## Code-Konventionen

- GUI-Aufbau ausschließlich in `main()`; das Modul bleibt import-/testbar.
- Reine Logik von der GUI trennen (z. B. `convert_files_to_workbook`).
- Aus Hintergrund-Threads Widgets **nur** über `_post_ui(...)` ansprechen.
