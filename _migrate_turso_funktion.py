"""Einmalige Migration: CHECK-Constraint in Turso erweitern + Werte migrieren."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from database.turso_sync import _turso_execute_batch

print("Starte Turso-Schema-Migration …")

_turso_execute_batch([
    # Temporäre Tabelle ohne CHECK-Constraint erstellen
    {"sql": """
        CREATE TABLE IF NOT EXISTS ma__mitarbeiter_new (
            id              INTEGER PRIMARY KEY,
            vorname         TEXT NOT NULL,
            nachname        TEXT NOT NULL,
            personalnummer  TEXT DEFAULT '',
            funktion        TEXT DEFAULT 'Schichtleiter',
            position        TEXT DEFAULT '',
            abteilung       TEXT DEFAULT '',
            email           TEXT DEFAULT '',
            telefon         TEXT DEFAULT '',
            eintrittsdatum  TEXT,
            status          TEXT DEFAULT 'aktiv',
            erstellt_am     TEXT,
            geaendert_am    TEXT
        )
    """},
    # Daten mit gemappten Werten hinüberkopieren
    {"sql": """
        INSERT OR REPLACE INTO ma__mitarbeiter_new
        SELECT id, vorname, nachname, personalnummer,
            CASE funktion
                WHEN 'stamm' THEN 'Schichtleiter'
                WHEN 'dispo' THEN 'Dispo'
                ELSE COALESCE(funktion, 'Schichtleiter')
            END,
            position, abteilung, email, telefon,
            eintrittsdatum, status, erstellt_am, geaendert_am
        FROM ma__mitarbeiter
    """},
    {"sql": "DROP TABLE ma__mitarbeiter"},
    {"sql": "ALTER TABLE ma__mitarbeiter_new RENAME TO ma__mitarbeiter"},
])

print("Turso-Migration abgeschlossen.")
