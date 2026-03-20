import sys, os
sys.path.insert(0, os.path.dirname(__file__))
os.chdir(os.path.dirname(__file__))

from functions.fahrzeug_functions import erstelle_fahrzeug, loesche_fahrzeug
from database import turso_sync

orig = turso_sync.push_row

def verbose_push(db, table, row):
    print(f"push_row: db={os.path.basename(str(db))}, table={table}, id={row.get('id')}")
    try:
        result = orig(db, table, row)
        print(f"  -> OK")
        return result
    except Exception as e:
        print(f"  -> FEHLER: {e}")
        raise

turso_sync.push_row = verbose_push

fid = erstelle_fahrzeug(kennzeichen="E2E-DBG", typ="PKW", marke="Test", modell="E2E_DEBUG_MODEL", baujahr=2026)
print("fid=", fid)
import time; time.sleep(4)

from database.turso_sync import _rows_from_turso
rows = _rows_from_turso("nesk3__fahrzeuge")
found = [r for r in rows if "E2E-DBG" in str(r.values())]
print("Turso gefunden:", found)
if fid:
    loesche_fahrzeug(fid)
    print("Testeintrag gelöscht.")
