# exceltool

Einzeldatei-Tkinter-GUI (`exceltool.py`) mit 5 Werkzeugen:
CSV/XLS/XLSX → Excel, Blattschutz entfernen, OpenTrans/ORDERS05/EDIFACT-Parser,
CSV-Splitter, XML → CSV.

## Starten
```
pip install -r requirements.txt      # pandas, openpyxl, xlrd
python exceltool.py                   # tkinter nötig (Linux: python3-tk)
```

## Tests
```
python test_exceltool.py
```
Headless lauffähig: `tkinter` wird im Test gestubbt, der GUI-Aufbau steckt in
`main()` und läuft beim Import nicht an. Reine Logik (Parser, Konverter,
Blattschutz, Splitter) ist ohne Display testbar.

## Code-Struktur (Konventionen)
- GUI-Aufbau ausschließlich in `main()`; Modul bleibt import-/testbar.
- Schwere Tasks laufen in Daemon-Threads; UI-Updates **nur** über `_post_ui(...)`
  (Queue + `_pump_ui_queue`), da Tkinter nicht thread-safe ist.
- Reine Logik von der GUI getrennt halten (z. B. `convert_files_to_workbook`),
  damit sie headless getestet werden kann; neue Funktionen ebenso bauen.
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
