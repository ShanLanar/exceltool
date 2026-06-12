"""
Daten-Tool (GUI) für:
  1) CSV/XLS/XLSX → Excel (Mehrfachauswahl, jede Datei/Sheet als eigener Reiter)
  2) Blattschutz entfernen (Mehrfachauswahl, .xls wird intern nach .xlsx konvertiert)
  3) Parser für OpenTrans / ORDERS05 / EDIFACT mit einheitlichem Excel-Export (Header + Positionen)

Hinweise:
- Benötigt: pandas, openpyxl, xlrd
- Einheitliches Exportformat:
    * Sheet "Header": Schlüssel/Wert-Paare, inkl. Adressen & Ansprechpartner
    * Sheet "Positionen": Position, Artikelnummer, Beschreibung, Menge, Einheit,
                          Preis, Nettowert, Lieferdatum, Incoterm, Steuer

BUGFIX: Syntaxfehler in analyze_text() – unterminated string literals repariert
        (echte Newlines in Strings ersetzt durch \\n Escape-Sequenzen)

Verbesserungen (siehe Commit-History):
  * Blattschutz-Entfernung ist jetzt namespace-sicher (kein Neuschreiben des
    kompletten XML mehr -> keine beschädigten Dateien)
  * GUI-Aufbau in main() gekapselt -> Modul ist ohne Display importier-/testbar
  * diverse Bugfixes (Encoding-Fallback, leere Zielmappe, Exception-Typ)
"""

from __future__ import annotations

import csv
import logging
import os
import queue
import threading
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox, ttk
from typing import Any

import pandas as pd

from core import (
    convert_files_to_workbook,
    entferne_schutz_on_file,
    escape_excel_formula,
    format_excel,
    parse_edifact,
    parse_opentrans_xml,
    parse_orders05_xml,
    parse_parameter_xml,
    safe_str,
    split_csv_file,
)

logger = logging.getLogger("exceltool")


# =========================================================
# Thread-Infrastruktur (nicht-blockierendes UI)
# =========================================================

# Wird in main() gesetzt; hier vordefiniert, damit Hintergrund-Threads
# gefahrlos prüfen können, ob das Fenster (noch) existiert.
root = None

_ui_queue: queue.Queue = queue.Queue()
_task_running = False
_pump_after_id = None


def _post_ui(fn) -> None:
    """Reiht eine UI-Aktion ein; sie wird im Main-Thread abgearbeitet."""
    _ui_queue.put(fn)


def _pump_ui_queue() -> None:
    """Arbeitet eingereihte UI-Aktionen im Main-Thread ab (periodisch)."""
    global _pump_after_id
    try:
        while True:
            fn = _ui_queue.get_nowait()
            try:
                fn()
            except Exception:
                logger.exception("UI-Update-Fehler")
    except queue.Empty:
        pass
    if root is not None:
        try:
            _pump_after_id = root.after(80, _pump_ui_queue)
        except tk.TclError:
            pass  # Fenster wird gerade geschlossen


def _shutdown() -> None:
    """Sauberes Beenden: laufenden Pump-Timer abbrechen, dann Fenster schließen."""
    global root, _pump_after_id
    if root is None:
        return
    if _pump_after_id is not None:
        try:
            root.after_cancel(_pump_after_id)
        except Exception:
            pass
        _pump_after_id = None
    win, root = root, None
    try:
        win.destroy()
    except Exception:
        pass


def _run_in_background(work) -> bool:
    """
    Führt work() in einem Daemon-Thread aus (nicht-blockierendes UI).
    work() darf Tk-Widgets ausschließlich über _post_ui(...) ansprechen.
    Parallele Tasks werden verhindert (gemeinsame Fortschrittsanzeige/Status).
    Gibt False zurück, wenn bereits ein Task läuft.
    """
    global _task_running
    if _task_running:
        messagebox.showinfo("Bitte warten", "Es läuft bereits eine Verarbeitung.")
        return False
    _task_running = True

    def runner():
        global _task_running
        try:
            work()
        except Exception as e:
            logger.exception("Hintergrund-Task fehlgeschlagen")
            _post_ui(lambda e=e: messagebox.showerror("Fehler", f"Unerwarteter Fehler: {e}"))
        finally:
            _task_running = False

    threading.Thread(target=runner, daemon=True).start()
    return True


def export_to_excel(header_dict: dict[str, Any], items_list: list[dict[str, Any]],
                    doc_id: str | None = None) -> None:
    """Exportiert Header und Positionen in eine Excel-Datei."""
    default_name = f"{doc_id or 'export'}.xlsx"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=default_name,
        title="Export speichern unter",
        filetypes=[("Excel-Datei", "*.xlsx")]
    )
    if not save_path:
        return
    try:
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            header_rows = [{"Feld": k, "Wert": escape_excel_formula(safe_str(v))}
                           for k, v in header_dict.items()]
            pd.DataFrame(header_rows).to_excel(writer, sheet_name="Header", index=False)
            all_keys = [
                "Position", "Artikelnummer", "Beschreibung", "Menge", "Einheit",
                "Preis", "Nettowert", "Lieferdatum", "Incoterm", "Steuer", "Währung"
            ]
            norm_items = [{k: escape_excel_formula(safe_str(it.get(k))) for k in all_keys}
                          for it in items_list]
            pd.DataFrame(norm_items).to_excel(writer, sheet_name="Positionen", index=False)
        format_excel(save_path)
        messagebox.showinfo("Erfolg", f"Export erfolgreich: {save_path}")
    except Exception as e:
        messagebox.showerror("Fehler", f"Export fehlgeschlagen: {e}")


# =========================================================
# CSV/XLS/XLSX → Excel (Konverter)
# =========================================================

selected_files_convert: list[str] = []


def select_files_convert(listbox: tk.Listbox, status_label: ttk.Label,
                         btn_convert: ttk.Button) -> None:
    global selected_files_convert
    files = filedialog.askopenfilenames(
        title="Dateien auswählen",
        filetypes=[("CSV/XLS/XLSX", "*.csv *.xls *.xlsx")]
    )
    if not files:
        return
    selected_files_convert = list(files)
    listbox.delete(0, tk.END)
    for f in selected_files_convert:
        listbox.insert(tk.END, f)
    btn_convert.config(state="normal")
    status_label.config(text=f"{len(selected_files_convert)} Datei(en) ausgewählt.")


def csv_xls_to_excel(progress_bar: ttk.Progressbar, status_label: ttk.Label) -> None:
    if not selected_files_convert:
        messagebox.showerror("Fehler", "Keine Dateien ausgewählt!")
        return
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        title="Zieldatei speichern unter",
        filetypes=[("Excel-Datei", "*.xlsx")]
    )
    if not save_path:
        return
    files = list(selected_files_convert)
    status_label.config(text="Verarbeite Dateien…")
    progress_bar["maximum"] = len(files)
    progress_bar["value"] = 0

    def work():
        def progress(i):
            _post_ui(lambda i=i: progress_bar.config(value=i))

        def on_file_error(file, e):
            _post_ui(lambda file=file, e=e:
                     messagebox.showerror("Fehler", f"Fehler bei {file}: {e}"))

        written = convert_files_to_workbook(files, save_path, progress, on_file_error)
        if written == 0:
            _post_ui(lambda: status_label.config(text="Keine Daten konvertiert."))
            _post_ui(lambda: messagebox.showwarning(
                "Hinweis", "Es konnten keine Daten konvertiert werden.\n"
                           "Bitte Eingabedateien prüfen."))
        else:
            _post_ui(lambda: status_label.config(text=f"Fertig! Datei gespeichert: {save_path}"))
            _post_ui(lambda: messagebox.showinfo("Erfolg", f"Excel-Datei erstellt: {save_path}"))

    _run_in_background(work)


# =========================================================
# Blattschutz entfernen (Stapellauf)
# =========================================================

selected_files_protect: list[str] = []


def select_files_protect(listbox: tk.Listbox, status_label: ttk.Label,
                         btn_protect: ttk.Button) -> None:
    """Excel-Dateien wählen und Button-Zustand steuern."""
    global selected_files_protect
    files = filedialog.askopenfilenames(
        title="Excel-Dateien auswählen (.xls/.xlsx)",
        filetypes=[("Excel-Dateien", "*.xls *.xlsx")]
    )
    selected_files_protect = list(files) if files else []
    listbox.delete(0, tk.END)
    for f in selected_files_protect:
        listbox.insert(tk.END, f)
    btn_protect.config(state=("normal" if selected_files_protect else "disabled"))
    status_label.config(
        text=f"{len(selected_files_protect)} Excel-Datei(en) für Entschützen ausgewählt."
    )


def remove_protection_batch(progress_bar: ttk.Progressbar, status_label: ttk.Label) -> None:
    if not selected_files_protect:
        return
    files = list(selected_files_protect)
    status_label.config(text="Entferne Blattschutz…")
    progress_bar["maximum"] = len(files)
    progress_bar["value"] = 0

    def work():
        results = []
        for i, file in enumerate(files, start=1):
            ok, msg = entferne_schutz_on_file(file)
            results.append((file, ok, msg))
            _post_ui(lambda i=i: progress_bar.config(value=i))
        success = [r for r in results if r[1]]
        fail = [r for r in results if not r[1]]
        info_text = ""
        if success:
            info_text += "Erfolgreich entschützt:\n" + "\n".join(
                [f"- {r[2]}" for r in success]) + "\n\n"
        if fail:
            info_text += "Fehlgeschlagen:\n" + "\n".join(
                [f"- {r[0]}: {r[2]}" for r in fail])
        _post_ui(lambda: status_label.config(text="Entschützen abgeschlossen."))
        _post_ui(lambda info_text=info_text:
                 messagebox.showinfo("Ergebnis", info_text or "Keine Ergebnisse."))

    _run_in_background(work)


# =========================================================
# Parser: OpenTrans / ORDERS05 / EDIFACT
# =========================================================

last_parsed_header: dict[str, Any] = {}
last_parsed_items: list[dict[str, Any]] = []
last_doc_id: str | None = None


# -------------------------------
# Analyse-Handler
# -------------------------------

def analyze_text(parser_input: tk.Text, parser_output: tk.Text,
                 btn_export_items: ttk.Button) -> None:
    """
    BUGFIX: Alle parser_output.insert()-Aufrufe nutzen jetzt korrekte \\n-Escape-Sequenzen
    statt echter Newlines im String-Literal (das war der SyntaxError).
    """
    global last_parsed_header, last_parsed_items, last_doc_id
    raw = parser_input.get("1.0", tk.END).strip()
    if not raw:
        messagebox.showerror("Fehler", "Bitte OpenTrans-, ORDERS05- oder EDIFACT-Text einfügen.")
        return
    last_parsed_header = {}
    last_parsed_items = []
    last_doc_id = None

    if raw.startswith("<"):
        try:
            root_xml = ET.fromstring(raw)
            root_name = root_xml.tag.split('}')[-1] if root_xml.tag.startswith("{") else root_xml.tag
        except Exception:
            root_name = ""
        if ("ORDERS05" in raw) or (root_name.upper() in ["ORDERS05", "IDOC"]):
            header, items, doc_id = parse_orders05_xml(raw)
        else:
            header, items, doc_id = parse_opentrans_xml(raw)
    else:
        header, items, doc_id = parse_edifact(raw)

    last_parsed_header = header
    last_parsed_items = items
    last_doc_id = doc_id

    parser_output.config(state="normal")
    parser_output.delete("1.0", tk.END)

    # *** BUGFIX: \n als Escape-Sequenz, nicht als echter Zeilenumbruch im Literal ***
    parser_output.insert(tk.END, "Kopf-Informationen:\n")
    for k, v in header.items():
        parser_output.insert(tk.END, f"  {k}: {safe_str(v)}\n")
    parser_output.insert(tk.END, "\nPositionen:\n")

    if items:
        cols = [
            "Position", "Artikelnummer", "Beschreibung", "Menge", "Einheit",
            "Preis", "Nettowert", "Lieferdatum", "Incoterm", "Steuer", "Währung"
        ]
        for it in items:
            line = " | ".join([safe_str(it.get(c)) for c in cols])
            parser_output.insert(tk.END, f"  - {line}\n")
    else:
        parser_output.insert(tk.END, "  (keine Positionen erkannt)\n")

    parser_output.config(state="disabled")
    btn_export_items.config(state=("normal" if items else "disabled"))


def export_items_to_excel() -> None:
    export_to_excel(last_parsed_header, last_parsed_items, last_doc_id)


def copy_summary_to_clipboard(parser_output: tk.Text) -> None:
    txt = parser_output.get("1.0", tk.END)
    root.clipboard_clear()
    root.clipboard_append(txt)
    messagebox.showinfo("Kopiert", "Zusammenfassung in die Zwischenablage kopiert.")


# =========================================================
# Tab 4: CSV-Splitter (split_csv_2000)
# =========================================================

selected_files_split: list[str] = []


def select_files_split(listbox: tk.Listbox, status_label: ttk.Label,
                       btn_split: ttk.Button) -> None:
    """CSV-Dateien für den Splitter auswählen."""
    global selected_files_split
    files = filedialog.askopenfilenames(
        title="CSV-Dateien auswählen",
        filetypes=[("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")]
    )
    selected_files_split = list(files) if files else []
    listbox.delete(0, tk.END)
    for f in selected_files_split:
        listbox.insert(tk.END, f)
    btn_split.config(state=("normal" if selected_files_split else "disabled"))
    status_label.config(text=f"{len(selected_files_split)} CSV-Datei(en) ausgewählt.")


def run_split(chunk_size_var: tk.IntVar, progress_bar: ttk.Progressbar,
              status_label: ttk.Label, log_widget: tk.Text) -> None:
    """Startet den Splitter für alle ausgewählten Dateien."""
    if not selected_files_split:
        return

    chunk_size = chunk_size_var.get()
    if chunk_size < 1:
        messagebox.showerror("Fehler", "Zeilenzahl pro Teil muss ≥ 1 sein.")
        return

    files = list(selected_files_split)

    def log(msg):
        def _do():
            log_widget.config(state="normal")
            log_widget.insert("end", msg + "\n")
            log_widget.see("end")
            log_widget.config(state="disabled")
        _post_ui(_do)

    progress_bar.config(value=0, maximum=len(files))
    status_label.config(text="Splitte CSV-Dateien…")

    def work():
        total_parts = 0
        for i, file_path in enumerate(files, 1):
            log(f"\n→ {os.path.basename(file_path)}")
            try:
                parts = split_csv_file(file_path, chunk_size, log)
                total_parts += parts
            except Exception as e:
                log(f"  ✖ Fehler: {e}")
            _post_ui(lambda i=i: progress_bar.config(value=i))
        _post_ui(lambda total_parts=total_parts:
                 status_label.config(text=f"Fertig: {total_parts} Teildatei(en) erzeugt."))
        _post_ui(lambda total_parts=total_parts: messagebox.showinfo(
            "CSV-Splitter",
            f"{total_parts} Teildatei(en) erzeugt.\nAusgabe im jeweiligen Quellordner."))

    _run_in_background(work)


def xml_to_csv_convert(xml_input: tk.Text, log_widget: tk.Text,
                       sep_var: tk.StringVar) -> None:
    """Liest XML aus dem Eingabefeld und speichert als CSV."""
    xml_text = xml_input.get("1.0", tk.END).strip()
    if not xml_text:
        messagebox.showwarning("Hinweis", "Bitte XML in das Eingabefeld einfügen.")
        return

    def log(msg):
        log_widget.config(state="normal")
        log_widget.insert("end", msg + "\n")
        log_widget.see("end")
        log_widget.config(state="disabled")

    try:
        header, values = parse_parameter_xml(xml_text)
    except ValueError as e:
        messagebox.showerror("Fehler", str(e))
        return

    sep = sep_var.get() or ";"
    save_path = filedialog.asksaveasfilename(
        title="CSV speichern unter",
        defaultextension=".csv",
        initialfile="partner_export.csv",
        filetypes=[("CSV-Datei", "*.csv")]
    )
    if not save_path:
        return

    try:
        with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=sep)
            writer.writerow(header)
            writer.writerow(values)
        log(f"✔ {len(header)} Felder exportiert → {save_path}")
        log(f"  Felder: {', '.join(header[:8])}{'…' if len(header) > 8 else ''}")
        messagebox.showinfo("Erfolg", f"CSV gespeichert:\n{save_path}\n\n{len(header)} Felder.")
    except Exception as e:
        messagebox.showerror("Fehler", f"Speichern fehlgeschlagen: {e}")


def paste_from_clipboard(xml_input: tk.Text) -> None:
    """Holt Inhalt aus Zwischenablage und setzt ihn ins Eingabefeld."""
    try:
        content = root.clipboard_get()
        xml_input.delete("1.0", tk.END)
        xml_input.insert("1.0", content)
    except tk.TclError:
        messagebox.showwarning("Zwischenablage", "Zwischenablage ist leer oder enthält keinen Text.")


def preview_xml(xml_input: tk.Text, preview_widget: tk.Text) -> None:
    """Zeigt eine Tabellenvorschau der geparsten Felder."""
    xml_text = xml_input.get("1.0", tk.END).strip()
    if not xml_text:
        return
    try:
        header, values = parse_parameter_xml(xml_text)
    except ValueError as e:
        preview_widget.config(state="normal")
        preview_widget.delete("1.0", tk.END)
        preview_widget.insert("1.0", f"Fehler: {e}")
        preview_widget.config(state="disabled")
        return

    preview_widget.config(state="normal")
    preview_widget.delete("1.0", tk.END)
    preview_widget.insert("end", f"{len(header)} Felder erkannt:\n\n")
    col_w = 28
    for h, v in zip(header, values, strict=False):
        line = f"  {h[:col_w]:<{col_w}}  {v[:60]}\n"
        preview_widget.insert("end", line)
    preview_widget.config(state="disabled")


# =========================================================
# GUI Aufbau
# =========================================================

def _scrolled(parent, factory, pack_kw, horizontal=False):
    """
    Bettet ein Listbox-/Text-Widget mit Scrollbar(s) in einen eigenen Rahmen ein.
    ``factory(frame)`` erzeugt das Widget; ``pack_kw`` platziert den Rahmen im
    übergeordneten Container (wie zuvor das Widget selbst).
    """
    frame = ttk.Frame(parent)
    frame.pack(**pack_kw)
    widget = factory(frame)
    vsb = ttk.Scrollbar(frame, orient="vertical", command=widget.yview)
    widget.configure(yscrollcommand=vsb.set)
    widget.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    if horizontal:
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=widget.xview)
        widget.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    return widget


def main():
    global root
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    root = tk.Tk()
    root.title("Excel-Toolbox: CSV→Excel | Blattschutz | Parser | CSV-Splitter | XML→CSV")
    root.geometry("1000x740")
    root.minsize(820, 600)

    tab_control = ttk.Notebook(root)
    tab_csv     = ttk.Frame(tab_control)
    tab_protect = ttk.Frame(tab_control)
    tab_parser  = ttk.Frame(tab_control)
    tab_split   = ttk.Frame(tab_control)
    tab_xml2csv = ttk.Frame(tab_control)

    tab_control.add(tab_csv,     text="📂 CSV/XLS/XLSX → Excel")
    tab_control.add(tab_protect, text="🔓 Blattschutz entfernen")
    tab_control.add(tab_parser,  text="📄 OpenTrans/ORDERS05/EDIFACT")
    tab_control.add(tab_split,   text="✂ CSV-Splitter")
    tab_control.add(tab_xml2csv, text="🔄 XML → CSV")
    tab_control.pack(expand=1, fill="both")

    # --- Tab 1: Konverter ---
    frame_csv_top = ttk.Frame(tab_csv)
    frame_csv_top.pack(fill="x", padx=10, pady=10)
    ttk.Label(frame_csv_top,
              text="Wähle CSV/XLS/XLSX-Dateien und füge alle als Reiter in eine Excel-Datei zusammen."
              ).pack(side="left")

    convert_listbox = _scrolled(
        tab_csv, lambda p: tk.Listbox(p, height=12),
        dict(fill="both", expand=True, padx=10, pady=5))

    frame_csv_bottom = ttk.Frame(tab_csv)
    frame_csv_bottom.pack(fill="x", padx=10, pady=10)

    status_label_global = ttk.Label(root, text="", foreground="blue")
    progress_bar_global = ttk.Progressbar(root, length=400, mode="determinate")

    btn_run_convert = ttk.Button(
        frame_csv_bottom, text="Zusammenführen & Speichern…",
        command=lambda: csv_xls_to_excel(progress_bar_global, status_label_global),
        state="disabled"
    )
    btn_select_convert = ttk.Button(
        frame_csv_bottom, text="Dateien auswählen…",
        command=lambda: select_files_convert(convert_listbox, status_label_global, btn_run_convert)
    )
    btn_select_convert.pack(side="left")
    btn_run_convert.pack(side="left", padx=10)

    # --- Tab 2: Schutz entfernen ---
    frame_protect_top = ttk.Frame(tab_protect)
    frame_protect_top.pack(fill="x", padx=10, pady=10)
    ttk.Label(frame_protect_top,
              text="Wähle Excel-Dateien (.xls/.xlsx), um Blattschutz/Arbeitsmappenschutz zu entfernen."
              ).pack(side="left")

    protect_listbox = _scrolled(
        tab_protect, lambda p: tk.Listbox(p, height=12),
        dict(fill="both", expand=True, padx=10, pady=5))

    frame_protect_bottom = ttk.Frame(tab_protect)
    frame_protect_bottom.pack(fill="x", padx=10, pady=10)

    btn_run_protect = ttk.Button(
        frame_protect_bottom, text="Blattschutz entfernen (Stapellauf)",
        command=lambda: remove_protection_batch(progress_bar_global, status_label_global),
        state="disabled"
    )
    btn_select_protect = ttk.Button(
        frame_protect_bottom, text="Excel-Dateien auswählen…",
        command=lambda: select_files_protect(protect_listbox, status_label_global, btn_run_protect)
    )
    btn_select_protect.pack(side="left")
    btn_run_protect.pack(side="left", padx=10)

    # --- Tab 3: Parser ---
    frame_parser_top = ttk.Frame(tab_parser)
    frame_parser_top.pack(fill="x", padx=10, pady=10)
    ttk.Label(frame_parser_top,
              text="OpenTrans-XML, SAP IDoc ORDERS05-XML oder EDIFACT-Rohtext einfügen und analysieren."
              ).pack(side="left")

    parser_input = _scrolled(
        tab_parser, lambda p: tk.Text(p, height=16),
        dict(fill="both", expand=True, padx=10, pady=5))

    frame_parser_actions = ttk.Frame(tab_parser)
    frame_parser_actions.pack(fill="x", padx=10, pady=5)

    btn_export_items = ttk.Button(
        frame_parser_actions, text="Export (Header + Positionen) → Excel…",
        command=export_items_to_excel, state="disabled"
    )
    btn_analyze = ttk.Button(
        frame_parser_actions, text="Analysieren",
        command=lambda: analyze_text(parser_input, parser_output, btn_export_items)
    )
    btn_copy = ttk.Button(
        frame_parser_actions, text="Zusammenfassung kopieren",
        command=lambda: copy_summary_to_clipboard(parser_output)
    )
    btn_analyze.pack(side="left")
    btn_copy.pack(side="left", padx=10)
    btn_export_items.pack(side="left", padx=10)

    parser_output = _scrolled(
        tab_parser, lambda p: tk.Text(p, height=18, state="disabled"),
        dict(fill="both", expand=True, padx=10, pady=5))

    # --- Tab 4: CSV-Splitter ---
    ttk.Label(tab_split,
              text="Teilt große CSV-Dateien in gleichmäßige Blöcke auf.\n"
                   "Jeder Block enthält die Kopfzeile. Ausgabe im Quellordner der Originaldatei."
              ).pack(anchor="w", padx=10, pady=(10, 4))

    frm_split_opts = ttk.LabelFrame(tab_split, text="Einstellungen")
    frm_split_opts.pack(fill="x", padx=10, pady=4)
    ttk.Label(frm_split_opts, text="Zeilen pro Datei:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    split_chunk_var = tk.IntVar(value=2000)
    ttk.Spinbox(frm_split_opts, from_=100, to=100000, increment=500,
                textvariable=split_chunk_var, width=10).grid(row=0, column=1, sticky="w", padx=6, pady=6)
    ttk.Label(frm_split_opts, text="(Standard: 2000 – passend für viele Shop-Importe)").grid(
        row=0, column=2, sticky="w", padx=6)

    split_listbox = _scrolled(
        tab_split, lambda p: tk.Listbox(p, height=7),
        dict(fill="both", expand=False, padx=10, pady=4))

    frm_split_log = ttk.LabelFrame(tab_split, text="Protokoll")
    frm_split_log.pack(fill="both", expand=True, padx=10, pady=4)
    split_log = _scrolled(
        frm_split_log, lambda p: tk.Text(p, height=8, state="disabled"),
        dict(fill="both", expand=True, padx=5, pady=5))

    frm_split_btn = ttk.Frame(tab_split)
    frm_split_btn.pack(fill="x", padx=10, pady=6)

    btn_run_split = ttk.Button(
        frm_split_btn, text="Splitten",
        command=lambda: run_split(split_chunk_var, progress_bar_global, status_label_global, split_log),
        state="disabled"
    )
    btn_select_split = ttk.Button(
        frm_split_btn, text="CSV-Dateien auswählen…",
        command=lambda: select_files_split(split_listbox, status_label_global, btn_run_split)
    )
    btn_select_split.pack(side="left")
    btn_run_split.pack(side="left", padx=10)

    # --- Tab 5: XML → CSV ---
    ttk.Label(tab_xml2csv,
              text="Liest ein XML-Fragment mit <PARAMETER DISPLAYNAME=\"...\">Wert</PARAMETER>\n"
                   "und exportiert es als CSV. Typisch für ETC-Partner-Daten aus der Zwischenablage."
              ).pack(anchor="w", padx=10, pady=(10, 4))

    frm_xml_opts = ttk.LabelFrame(tab_xml2csv, text="Einstellungen")
    frm_xml_opts.pack(fill="x", padx=10, pady=4)
    ttk.Label(frm_xml_opts, text="Trennzeichen:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    xml_sep_var = tk.StringVar(value=";")
    sep_box = ttk.Combobox(frm_xml_opts, values=[";", ",", "\t", "|"],
                           textvariable=xml_sep_var, width=6, state="normal")
    sep_box.grid(row=0, column=1, sticky="w", padx=6, pady=6)
    ttk.Label(frm_xml_opts, text="(Standard: Semikolon)").grid(row=0, column=2, sticky="w", padx=6)

    frm_xml_input = ttk.LabelFrame(tab_xml2csv, text="XML-Eingabe (einfügen oder Zwischenablage)")
    frm_xml_input.pack(fill="both", expand=True, padx=10, pady=4)
    xml_input = _scrolled(
        frm_xml_input, lambda p: tk.Text(p, height=10, wrap="none"),
        dict(fill="both", expand=True, padx=5, pady=5), horizontal=True)

    frm_xml_preview = ttk.LabelFrame(tab_xml2csv, text="Vorschau erkannter Felder")
    frm_xml_preview.pack(fill="both", expand=True, padx=10, pady=4)
    xml_preview = _scrolled(
        frm_xml_preview, lambda p: tk.Text(p, height=7, state="disabled",
                                           font=("Courier New", 9)),
        dict(fill="both", expand=True, padx=5, pady=5), horizontal=True)

    frm_xml_btn = ttk.Frame(tab_xml2csv)
    frm_xml_btn.pack(fill="x", padx=10, pady=6)
    ttk.Button(frm_xml_btn, text="📋 Aus Zwischenablage",
               command=lambda: paste_from_clipboard(xml_input)).pack(side="left")
    ttk.Button(frm_xml_btn, text="🔍 Vorschau",
               command=lambda: preview_xml(xml_input, xml_preview)).pack(side="left", padx=8)
    ttk.Button(frm_xml_btn, text="💾 Als CSV speichern…",
               command=lambda: xml_to_csv_convert(xml_input, xml_preview, xml_sep_var)).pack(side="left")
    ttk.Button(frm_xml_btn, text="🗑 Eingabe leeren",
               command=lambda: xml_input.delete("1.0", tk.END)).pack(side="right")

    # --- Globaler Status ---
    status_frame = ttk.Frame(root)
    status_frame.pack(fill="x", padx=10, pady=5)
    status_label_global.pack(side="left")
    progress_bar_global.pack(side="right")

    # Sauberes Beenden (Pump-Timer abbrechen) beim Schließen des Fensters
    root.protocol("WM_DELETE_WINDOW", _shutdown)
    # Hintergrund-Tasks melden UI-Updates über diese Pumpe zurück
    _pump_ui_queue()
    root.mainloop()


# =========================================================
# Start
# =========================================================
if __name__ == "__main__":
    main()
