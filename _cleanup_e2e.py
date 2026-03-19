"""
Bereinigt alle E2E-Test-Eintraege aus lokalen DBs und aus Turso.
"""
import sys, os, sqlite3
sys.path.insert(0, os.path.dirname(__file__))
os.chdir(os.path.dirname(__file__))

from config import BASE_DIR, DB_PATH, MITARBEITER_DB_PATH
from pathlib import Path

TAG = "__E2E_TEST__"
DB_SQL = Path(BASE_DIR) / "database SQL"

# (db_pfad, tabelle, suchfeld, lösch-funktion oder None)
checks = [
    (str(DB_SQL / "psa.db"),              "psa_verstoss",         "mitarbeiter"),
    (str(DB_SQL / "verspaetungen.db"),    "verspaetungen",        "mitarbeiter"),
    (str(DB_SQL / "call_transcription.db"),"call_logs",           "anrufer"),
    (str(DB_SQL / "telefonnummern.db"),   "telefonnummern",       "bezeichnung"),
    (str(DB_SQL / "patienten_station.db"),"patienten",            "patient_name"),
    (str(DB_SQL / "einsaetze.db"),        "einsaetze",            "einsatzstichwort"),
    (MITARBEITER_DB_PATH,                 "mitarbeiter",          "vorname"),
    (DB_PATH,                             "fahrzeuge",            "modell"),
    (DB_PATH,                             "uebergabe_protokolle", "personal"),
]

# Turso-Tabellennamen für push_delete
from database.turso_sync import TABLE_MAP
def turso_name(db_pfad, table):
    db_file = os.path.basename(db_pfad)
    return TABLE_MAP.get((db_file, table))

print()
print("=== E2E Cleanup ===")
total_deleted = 0

for db, table, field in checks:
    db_file = os.path.basename(db)
    try:
        conn = sqlite3.connect(db, timeout=5)
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            f"SELECT id, {field} FROM {table} WHERE {field} LIKE ?",
            (f"%{TAG}%",)
        ).fetchall()

        if not rows:
            conn.close()
            continue

        for r in rows:
            rid = r["id"]
            val = r[field]
            # Lokal löschen
            conn.execute(f"DELETE FROM {table} WHERE id = ?", (rid,))
            conn.commit()
            print(f"  [LOCAL DEL]  {db_file:30s}  {table:30s}  id={rid}")

            # Turso löschen
            tname = turso_name(db, table)
            if tname:
                try:
                    from database.turso_sync import push_delete
                    push_delete(db, table, rid)
                    print(f"  [TURSO DEL]  {tname}")
                except Exception as e:
                    print(f"  [TURSO ERR]  {tname}: {e}")
            total_deleted += 1

        conn.close()
    except Exception as e:
        print(f"  [ERR]  {db_file}  {table}: {e}")

if total_deleted == 0:
    print("  Keine Resteintraege gefunden - alles sauber.")
else:
    print(f"\n  {total_deleted} Eintraege bereinigt.")

# Turso direkt prüfen
print()
print("=== Turso Prüfung ===")
from database.turso_sync import _rows_from_turso
turso_tables = [v for v in TABLE_MAP.values()]
turso_rest = 0
for tname in turso_tables:
    try:
        rows = _rows_from_turso(tname)
        for r in rows:
            if TAG in str(list(r.values())):
                print(f"  [TURSO NOCH DA]  {tname}  id={r.get('id')}")
                # direkt löschen
                from database.turso_sync import _turso_execute_batch
                _turso_execute_batch([{"sql": f'DELETE FROM "{tname}" WHERE id = ?',
                                       "args": [{"type": "text", "value": str(r.get("id"))}]}])
                print(f"  [TURSO DEL]  {tname}  id={r.get('id')}")
                turso_rest += 1
    except Exception:
        pass

if turso_rest == 0:
    print("  Keine Resteintraege in Turso.")

print()
print("Fertig.")
