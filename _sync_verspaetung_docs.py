"""
Synchronisiert Protokoll-Ordner mit der Verspätungs-DB:
1. Für jeden DB-Eintrag ohne/ungültigem Pfad → Word-Dokument erstellen + DB aktualisieren
2. Dateien im Protokoll-Ordner, die NICHT in der DB referenziert sind → löschen
"""
import sqlite3
import os

# DB direkt öffnen (ohne turso_sync-Push)
_DB_PFAD = r"database SQL\verspaetungen.db"

from functions.verspaetung_functions import erstelle_verspaetungs_dokument, PROTOKOLL_DIR

con = sqlite3.connect(_DB_PFAD, timeout=5)
con.row_factory = sqlite3.Row

rows = con.execute("SELECT * FROM verspaetungen ORDER BY id").fetchall()

referenzierte = set()

print("=== Schritt 1: Dokumente erstellen / reparieren ===")
for r in rows:
    pfad = r["dokument_pfad"] or ""
    if pfad and os.path.isfile(pfad):
        referenzierte.add(os.path.normpath(pfad))
        print(f"  [OK]    ID {r['id']:3} {r['datum']} {r['mitarbeiter']}")
    else:
        daten = dict(r)
        try:
            neuer_pfad = erstelle_verspaetungs_dokument(daten)
            con.execute(
                "UPDATE verspaetungen SET dokument_pfad=? WHERE id=?",
                (neuer_pfad, r["id"])
            )
            referenzierte.add(os.path.normpath(neuer_pfad))
            tag = "ERSTELLT" if not pfad else "NEU"
            print(f"  [{tag}] ID {r['id']:3} {r['datum']} {r['mitarbeiter']} → {os.path.basename(neuer_pfad)}")
        except Exception as exc:
            print(f"  [FEHLER] ID {r['id']:3} {r['datum']} {r['mitarbeiter']}: {exc}")

con.commit()

print("\n=== Schritt 2: Verwaiste Dateien löschen ===")
geloescht = 0
for datei in sorted(PROTOKOLL_DIR.iterdir()):
    if datei.is_file() and datei.suffix.lower() == ".docx":
        if os.path.normpath(str(datei)) not in referenzierte:
            print(f"  [LÖSCHEN] {datei.name}")
            datei.unlink()
            geloescht += 1

if geloescht == 0:
    print("  Keine verwaisten Dateien gefunden.")

con.close()
print(f"\nFertig. {geloescht} Datei(en) gelöscht.")
