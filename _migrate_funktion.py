"""Einmalige Migration: funktion-Spalte CHECK-Constraint erweitern."""
import sqlite3, os, sys

db = os.path.join(os.path.dirname(__file__), "database SQL", "mitarbeiter.db")
if not os.path.exists(db):
    print("mitarbeiter.db nicht gefunden – nichts zu tun.")
    sys.exit(0)

conn = sqlite3.connect(db)
conn.execute("PRAGMA journal_mode = WAL")

row = conn.execute(
    "SELECT sql FROM sqlite_master WHERE type='table' AND name='mitarbeiter'"
).fetchone()
current_sql = row[0] if row else ""
print("Aktuell:", current_sql[:150])

if "'stamm','dispo'" not in current_sql and "'Schichtleiter'" in current_sql:
    print("Migration bereits durchgeführt – nichts zu tun.")
    conn.close()
    sys.exit(0)

print("Starte Migration …")
conn.executescript("""
PRAGMA foreign_keys = OFF;
BEGIN;
CREATE TABLE IF NOT EXISTS mitarbeiter_new (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    vorname         TEXT NOT NULL,
    nachname        TEXT NOT NULL,
    personalnummer  TEXT DEFAULT '',
    funktion        TEXT DEFAULT 'Schichtleiter'
                    CHECK (funktion IN ('Schichtleiter','Dispo','Betreuer')),
    position        TEXT DEFAULT '',
    abteilung       TEXT DEFAULT '',
    email           TEXT DEFAULT '',
    telefon         TEXT DEFAULT '',
    eintrittsdatum  TEXT,
    status          TEXT DEFAULT 'aktiv',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime')),
    geaendert_am    TEXT DEFAULT (datetime('now','localtime'))
);
INSERT INTO mitarbeiter_new
    SELECT id, vorname, nachname,
        COALESCE(personalnummer,''),
        CASE funktion
            WHEN 'stamm'  THEN 'Schichtleiter'
            WHEN 'dispo'  THEN 'Dispo'
            ELSE COALESCE(funktion,'Schichtleiter')
        END,
        COALESCE(position,''), COALESCE(abteilung,''),
        COALESCE(email,''), COALESCE(telefon,''),
        eintrittsdatum, COALESCE(status,'aktiv'),
        erstellt_am, geaendert_am
    FROM mitarbeiter;
DROP TABLE mitarbeiter;
ALTER TABLE mitarbeiter_new RENAME TO mitarbeiter;
COMMIT;
PRAGMA foreign_keys = ON;
""")
conn.close()
print("Migration erfolgreich abgeschlossen.")
