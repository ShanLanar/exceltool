"""
Tests für die GUI-/Threading-Schicht von exceltool.py.

Die reine Logik wird in test_core.py geprüft (ohne tkinter). Hier geht es nur um
Dinge, die das GUI-Modul betreffen: dass es sauber importiert und dass der
Hintergrund-Runner samt Reentrancy-Flag funktioniert.

tkinter wird nur gestubbt, wenn es nicht installiert ist (Headless-/CI-Umgebung).
Der GUI-Aufbau liegt in main() und wird beim bloßen Import NICHT ausgeführt.
"""

import importlib.util
import os
import sys
import time
import types

if importlib.util.find_spec("tkinter") is None:  # pragma: no cover - umgebungsabhängig
    _stub = types.ModuleType("tkinter")
    for _sub in ("filedialog", "messagebox", "ttk"):
        _m = types.ModuleType(f"tkinter.{_sub}")
        setattr(_stub, _sub, _m)
        sys.modules[f"tkinter.{_sub}"] = _m
    sys.modules["tkinter"] = _stub

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import core  # noqa: E402
import exceltool as et  # noqa: E402, I001  (Import bewusst nach dem tkinter-Stub)


def check(cond, msg):
    """Wie ein assert, mit Klartext-Ausgabe; pytest-tauglich (wirft bei Fehler)."""
    if cond:
        print(f"  ok   - {msg}")
    else:
        print(f"  FAIL - {msg}")
        raise AssertionError(msg)


def test_module_imports_and_reexports():
    print("test_module_imports_and_reexports")
    check(callable(et.main), "main() ist aufrufbar")
    # exceltool re-exportiert die reine Logik aus core (gleiche Funktionsobjekte)
    check(et.parse_edifact is core.parse_edifact, "parse_edifact aus core re-exportiert")
    check(et.convert_files_to_workbook is core.convert_files_to_workbook,
          "convert_files_to_workbook aus core re-exportiert")
    check(callable(et._run_in_background), "_run_in_background vorhanden")


def test_run_in_background_smoke():
    print("test_run_in_background_smoke")
    import threading as _t
    done = _t.Event()
    result = {}

    def work():
        result["ran"] = True
        done.set()

    started = et._run_in_background(work)
    check(started, "Task gestartet")
    check(done.wait(timeout=5), "work() im Hintergrund-Thread ausgeführt")
    check(result.get("ran") is True, "Seiteneffekt sichtbar")
    # Flag muss zurückgesetzt werden (kurz pollen, da finally im Thread läuft)
    for _ in range(50):
        if not et._task_running:
            break
        time.sleep(0.02)
    check(not et._task_running, "_task_running nach Abschluss zurückgesetzt")


def main():
    failures = []
    for t in (test_module_imports_and_reexports, test_run_in_background_smoke):
        try:
            t()
        except AssertionError as e:
            failures.append(f"{t.__name__}: {e}")
        except Exception as e:  # noqa: BLE001
            failures.append(f"{t.__name__}: unerwartet {e!r}")
    print("\n" + ("=" * 50))
    if failures:
        print(f"FEHLGESCHLAGEN: {len(failures)} Test(s)")
        for m in failures:
            print(f"  - {m}")
        raise SystemExit(1)
    print("ALLE TESTS BESTANDEN")


if __name__ == "__main__":
    main()
