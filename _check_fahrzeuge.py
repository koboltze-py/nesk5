import sqlite3, pathlib

TAG = "__E2E_TEST__"

for p in pathlib.Path(".").rglob("nesk3.db"):
    con = sqlite3.connect(str(p))
    con.row_factory = sqlite3.Row
    tables = [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
    if "fahrzeuge" in tables:
        print(f"DB: {p}")
        rows = con.execute(
            "SELECT * FROM fahrzeuge WHERE modell LIKE ? OR kennzeichen LIKE ?",
            (f"%{TAG}%", f"%{TAG}%")
        ).fetchall()
        print(f"  E2E-Eintraege in fahrzeuge: {len(rows)}")
        for r in rows:
            print("  ", dict(r))

        # Auch alle Zeilen zeigen (ohne Filter)
        alle = con.execute("SELECT id, modell, kennzeichen FROM fahrzeuge ORDER BY id DESC LIMIT 10").fetchall()
        print(f"  Letzte 10 Eintraege:")
        for r in alle:
            print("  ", dict(r))
    con.close()
