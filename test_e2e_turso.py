"""
End-to-End-Test: Legt in jeder Datenbank einen Testeintrag an (wie ein User),
prüft ob er lokal vorhanden ist UND ob er in Turso angekommen ist.
Am Ende werden alle Testeinträge wieder gelöscht.

Aufruf:
    python test_e2e_turso.py
"""
import sys, os, sqlite3, time
sys.path.insert(0, os.path.dirname(__file__))
os.chdir(os.path.dirname(__file__))

from database.turso_sync import _rows_from_turso, TABLE_MAP

TURSO_TEST_TAG = "__E2E_TEST__"  # wird in Text-Felder geschrieben, damit wir Testzeilen sicher erkennen

results: list[tuple[str, str, bool, str]] = []   # (db, tabelle, ok, fehler)

# Kleine Hilfsfunktion: Turso-Tabellenname aus (db_dateiname, tabelle)
def turso_tname(db_filename: str, table: str) -> str | None:
    return TABLE_MAP.get((db_filename, table))

# Warte max n Sekunden bis Turso eine Zeile mit dem Tag enthält
def warte_auf_turso(turso_table: str, feld: str, tag: str, timeout: int = 15) -> dict | None:
    for _ in range(timeout * 2):
        rows = _rows_from_turso(turso_table)
        for r in rows:
            if str(r.get(feld, "") or "").find(tag) >= 0:
                return r
        time.sleep(0.5)
    return None

def ok(db, table):
    results.append((db, table, True, ""))
    print(f"  [PASS]  {db:30s}  {table}")

def fail(db, table, err):
    results.append((db, table, False, str(err)))
    print(f"  [FAIL]  {db:30s}  {table}")
    print(f"           -> {err}")

# ══════════════════════════════════════════════════════════════════════════════
#  1. psa.db  →  psa_verstoss
# ══════════════════════════════════════════════════════════════════════════════
def test_psa():
    from functions.psa_db import psa_speichern, psa_loeschen, _DB_PFAD
    jetzt = time.strftime("%Y-%m-%dT%H:%M:%S")
    row_id = None
    try:
        row_id = psa_speichern({
            "mitarbeiter": TURSO_TEST_TAG,
            "datum": jetzt[:10],
            "psa_typ": "Helm",
            "bemerkung": TURSO_TEST_TAG,
            "aufgenommen_von": "E2E-Test",
        })
        assert row_id, "keine ID zurückgegeben"
        # lokal prüfen
        conn = sqlite3.connect(str(_DB_PFAD)); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM psa_verstoss WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        # Turso prüfen
        turso = warte_auf_turso("psa__psa_verstoss", "mitarbeiter", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("psa.db", "psa_verstoss")
    except Exception as e:
        fail("psa.db", "psa_verstoss", e)
    finally:
        if row_id:
            try: psa_loeschen(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  2. verspaetungen.db  →  verspaetungen
# ══════════════════════════════════════════════════════════════════════════════
def test_verspaetung():
    from functions.verspaetung_db import verspaetung_speichern, verspaetung_loeschen, _DB_PFAD
    row_id = None
    try:
        row_id = verspaetung_speichern({
            "mitarbeiter": TURSO_TEST_TAG,
            "datum": time.strftime("%Y-%m-%d"),
            "dienst": "",
            "dienstbeginn": "06:00",
            "dienstantritt": "06:05",
            "verspaetung_min": 5,
            "begruendung": TURSO_TEST_TAG,
            "aufgenommen_von": "E2E-Test",
        })
        assert row_id, "keine ID"
        conn = sqlite3.connect(str(_DB_PFAD)); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM verspaetungen WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        turso = warte_auf_turso("vers__verspaetungen", "mitarbeiter", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("verspaetungen.db", "verspaetungen")
    except Exception as e:
        fail("verspaetungen.db", "verspaetungen", e)
    finally:
        if row_id:
            try: verspaetung_loeschen(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  3. call_transcription.db  →  call_logs
# ══════════════════════════════════════════════════════════════════════════════
def test_call_transcription():
    from functions.call_transcription_db import speichern, loeschen
    row_id = None
    try:
        row_id = speichern({
            "datum": time.strftime("%Y-%m-%d"),
            "uhrzeit": time.strftime("%H:%M"),
            "anrufer": TURSO_TEST_TAG,
            "betreff": TURSO_TEST_TAG,
            "notiz": "Automatischer E2E-Test",
            "richtung": "Eingehend",
            "erledigt": 0,
        })
        assert row_id, "keine ID"
        from functions.call_transcription_db import _DB_PATH
        conn = sqlite3.connect(_DB_PATH); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM call_logs WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        turso = warte_auf_turso("call__call_logs", "anrufer", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("call_transcription.db", "call_logs")
    except Exception as e:
        fail("call_transcription.db", "call_logs", e)
    finally:
        if row_id:
            try: loeschen(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  4. telefonnummern.db  →  telefonnummern
# ══════════════════════════════════════════════════════════════════════════════
def test_telefonnummern():
    from functions.telefonnummern_db import eintrag_speichern, eintrag_loeschen, _DB_PFAD
    row_id = None
    try:
        row_id = eintrag_speichern({
            "bezeichnung": TURSO_TEST_TAG,
            "nummer": "0000-E2E",
            "kategorie": "Test",
            "quelle": "E2E",
            "sheet": "E2E",
        })
        assert row_id, "keine ID"
        conn = sqlite3.connect(str(_DB_PFAD)); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM telefonnummern WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        turso = warte_auf_turso("tel__telefonnummern", "bezeichnung", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("telefonnummern.db", "telefonnummern")
    except Exception as e:
        fail("telefonnummern.db", "telefonnummern", e)
    finally:
        if row_id:
            try: eintrag_loeschen(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  5. patienten_station.db  →  patienten
# ══════════════════════════════════════════════════════════════════════════════
def test_patienten():
    import importlib.util, sys as _sys
    spec = importlib.util.spec_from_file_location("dienstliches", "gui/dienstliches.py")
    m = importlib.util.module_from_spec(spec)
    # GUI-Importe temporär abfangen
    import unittest.mock as _mock
    with _mock.patch.dict(_sys.modules, {
        "PySide6": _mock.MagicMock(), "PySide6.QtWidgets": _mock.MagicMock(),
        "PySide6.QtCore": _mock.MagicMock(), "PySide6.QtGui": _mock.MagicMock(),
    }):
        spec.loader.exec_module(m)

    row_id = None
    try:
        daten = {
            "erstellt_am": time.strftime("%Y-%m-%d %H:%M:%S"),
            "datum": time.strftime("%d.%m.%Y"),
            "uhrzeit": time.strftime("%H:%M"),
            "behandlungsdauer": 0,
            "patient_name": TURSO_TEST_TAG,
            "patient_alter": 0,
            "geschlecht": "",
            "diagnose": TURSO_TEST_TAG,
            "massnahmen": "",
            "symptome": "",
            "beschwerde_art": "",
            "patient_typ": "",
            "patient_abteilung": "",
            "hergang_was": "",
            "hergang_wie": "",
            "unfall_ort": "",
            "abcde_a":"","abcde_b":"","abcde_c":"","abcde_d":"","abcde_e":"",
            "monitoring_bz":"","monitoring_rr":"","monitoring_spo2":"","monitoring_hf":"",
            "vorerkrankungen":"","medikamente_patient":"",
            "medikamente_gegeben":0,"medikamente_gegeben_was":"",
            "arbeitsunfall":0,"arbeitsunfall_details":"",
            "drk_ma1":"","drk_ma2":"","weitergeleitet":"","bemerkung":"",
            "_medikamente":[],
        }
        row_id = m.patient_speichern(daten, [])
        assert row_id, "keine ID"
        conn = sqlite3.connect(m._PATIENTEN_DB_PFAD); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM patienten WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        turso = warte_auf_turso("pat__patienten", "patient_name", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("patienten_station.db", "patienten")
    except Exception as e:
        fail("patienten_station.db", "patienten", e)
    finally:
        if row_id:
            try: m.patient_loeschen(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  6. einsaetze.db  →  einsaetze
# ══════════════════════════════════════════════════════════════════════════════
def test_einsaetze():
    import importlib.util, sys as _sys, unittest.mock as _mock
    spec = importlib.util.spec_from_file_location("dienstliches2", "gui/dienstliches.py")
    m = importlib.util.module_from_spec(spec)
    with _mock.patch.dict(_sys.modules, {
        "PySide6": _mock.MagicMock(), "PySide6.QtWidgets": _mock.MagicMock(),
        "PySide6.QtCore": _mock.MagicMock(), "PySide6.QtGui": _mock.MagicMock(),
    }):
        spec.loader.exec_module(m)

    row_id = None
    try:
        row_id = m.einsatz_speichern({
            "datum": time.strftime("%d.%m.%Y"),
            "uhrzeit": time.strftime("%H:%M"),
            "einsatzdauer": 0,
            "einsatzstichwort": TURSO_TEST_TAG,
            "einsatzort": TURSO_TEST_TAG,
            "einsatznr_drk": "E2E-TEST",
            "drk_ma1": "", "drk_ma2": "",
            "angenommen": True,
            "grund_abgelehnt": "",
            "bemerkung": "",
        })
        assert row_id, "keine ID"
        conn = sqlite3.connect(m._EINSATZ_DB_PFAD); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM einsaetze WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        turso = warte_auf_turso("einsaetze__einsaetze", "einsatzstichwort", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("einsaetze.db", "einsaetze")
    except Exception as e:
        fail("einsaetze.db", "einsaetze", e)
    finally:
        if row_id:
            try: m.einsatz_loeschen(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  7. mitarbeiter.db  →  mitarbeiter
# ══════════════════════════════════════════════════════════════════════════════
def test_mitarbeiter():
    from functions.mitarbeiter_functions import mitarbeiter_erstellen, mitarbeiter_loeschen
    from config import MITARBEITER_DB_PATH as _MA_DB_PATH
    from database.models import Mitarbeiter
    m = Mitarbeiter(
        id=None,
        vorname=TURSO_TEST_TAG,
        nachname="E2E",
        personalnummer="E2E-999",
        funktion="stamm",
        position="Test",
        abteilung="Test",
        email="",
        telefon="",
        eintrittsdatum=None,
        status="aktiv",
    )
    row_id = None
    try:
        m = mitarbeiter_erstellen(m)
        row_id = m.id
        assert row_id, "keine ID"
        conn = sqlite3.connect(_MA_DB_PATH); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM mitarbeiter WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        turso = warte_auf_turso("ma__mitarbeiter", "vorname", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("mitarbeiter.db", "mitarbeiter")
    except Exception as e:
        fail("mitarbeiter.db", "mitarbeiter", e)
    finally:
        if row_id:
            try: mitarbeiter_loeschen(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  8. nesk3.db  →  fahrzeuge
# ══════════════════════════════════════════════════════════════════════════════
def test_fahrzeug():
    from functions.fahrzeug_functions import erstelle_fahrzeug, loesche_fahrzeug
    from config import DB_PATH as _NESK3_DB_PATH
    row_id = None
    try:
        row_id = erstelle_fahrzeug(
            kennzeichen="E2E-TEST",
            typ="PKW",
            marke="Test",
            modell=TURSO_TEST_TAG,
            baujahr=2026,
            notizen=TURSO_TEST_TAG,
        )
        assert row_id, "keine ID"
        conn = sqlite3.connect(_NESK3_DB_PATH); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM fahrzeuge WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        turso = warte_auf_turso("nesk3__fahrzeuge", "modell", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("nesk3.db", "fahrzeuge")
    except Exception as e:
        fail("nesk3.db", "fahrzeuge", e)
    finally:
        if row_id:
            try: loesche_fahrzeug(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  9. nesk3.db  →  uebergabe_protokolle
# ══════════════════════════════════════════════════════════════════════════════
def test_uebergabe():
    from functions.uebergabe_functions import erstelle_protokoll, loesche_protokoll
    from config import DB_PATH as _NESK3_DB_PATH
    row_id = None
    try:
        row_id = erstelle_protokoll(
            datum=time.strftime("%Y-%m-%d"),
            schicht_typ="tagdienst",
            beginn_zeit="06:00",
            ende_zeit="14:00",
            patienten_anzahl=0,
            personal=TURSO_TEST_TAG,
            ereignisse=TURSO_TEST_TAG,
            massnahmen="",
            uebergabe_notiz="",
            ersteller=TURSO_TEST_TAG,
        )
        assert row_id, "keine ID"
        conn = sqlite3.connect(_NESK3_DB_PATH); conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM uebergabe_protokolle WHERE id=?", (row_id,)).fetchone()
        conn.close()
        assert row, "lokaler Eintrag fehlt"
        turso = warte_auf_turso("nesk3__uebergabe_protokolle", "personal", TURSO_TEST_TAG)
        assert turso, "Turso-Eintrag nicht gefunden"
        ok("nesk3.db", "uebergabe_protokolle")
    except Exception as e:
        fail("nesk3.db", "uebergabe_protokolle", e)
    finally:
        if row_id:
            try: loesche_protokoll(row_id)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
#  Stellungnahmen hat keinen einfachen Speicherpfad ohne Word-Datei → überspringen
#  (würde echte Word-Erstellung erfordern)
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print()
    print("=== E2E Turso Integration-Test ===")
    print(f"    Tag: '{TURSO_TEST_TAG}'")
    print()

    tests = [
        ("psa.db           → psa_verstoss",       test_psa),
        ("verspaetungen.db → verspaetungen",       test_verspaetung),
        ("call_log.db      → call_logs",           test_call_transcription),
        ("telefonnummern   → telefonnummern",      test_telefonnummern),
        ("patienten_st.    → patienten",           test_patienten),
        ("einsaetze.db     → einsaetze",           test_einsaetze),
        ("mitarbeiter.db   → mitarbeiter",         test_mitarbeiter),
        ("nesk3.db         → fahrzeuge",           test_fahrzeug),
        ("nesk3.db         → uebergabe_protokolle",test_uebergabe),
    ]

    for label, fn in tests:
        print(f"\n── {label}")
        fn()

    print()
    print("══════════════════════════════════════")
    passed = sum(1 for _, _, ok_, _ in results if ok_)
    total  = len(results)
    for db, table, ok_, err in results:
        icon = "✅" if ok_ else "❌"
        print(f"  {icon}  {db:28s}  {table}")
        if err:
            print(f"     -> {err}")
    print()
    print(f"  {passed}/{total} Datenbanken OK")
    sys.exit(0 if passed == total else 1)
