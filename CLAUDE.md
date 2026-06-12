# exceltool

Tkinter-GUI mit 5 Werkzeugen: CSV/XLS/XLSX → Excel, Blattschutz entfernen,
OpenTrans/ORDERS05/EDIFACT-Parser, CSV-Splitter, XML → CSV.

Aufgeteilt in zwei Module:
- `core.py` – GUI-freie Logik (Parser, Konverter, Splitter, Blattschutz). Keine
  tkinter-Abhängigkeit, direkt importier-/testbar.
- `exceltool.py` – GUI (tkinter) + Threading; importiert die Logik aus `core`.

## Starten
```
pip install -e ".[dev]"               # pandas, openpyxl, xlrd, pytest
python exceltool.py                   # tkinter nötig (Linux: python3-tk)
```

## Tests
```
pytest                                # bzw. python test_core.py / test_exceltool.py
```
`test_core.py` testet die reine Logik (braucht kein tkinter). `test_exceltool.py`
prüft GUI-/Threading-Schicht; `tkinter` wird dort gestubbt, falls nicht vorhanden.
Der GUI-Aufbau steckt in `main()` und läuft beim Import nicht an.

## Code-Struktur (Konventionen)
- Neue **GUI-freie** Logik nach `core.py` (headless testbar); GUI-Code/Threading
  nach `exceltool.py`.
- GUI-Aufbau ausschließlich in `main()`; Module bleiben import-/testbar.
- Schwere Tasks laufen in Daemon-Threads; UI-Updates **nur** über `_post_ui(...)`
  (Queue + `_pump_ui_queue`), da Tkinter nicht thread-safe ist.
- Blattschutz-Entfernung bearbeitet OOXML-XML rein textuell (kein ElementTree-
  Reparse → kein Namespace-Mangling).

## Git-Workflow (Wunsch des Owners)
**Zum Abschluss jeder Aufgabe** den Arbeits-Branch per Fast-Forward nach `main`
mergen und `main` pushen, damit ein einfaches `git pull` auf `main` genügt:
```
git switch main
git merge --ff-only <arbeitsbranch>
git push origin main
```
(Stehende Freigabe des Owners, nach `main` zu mergen.)
