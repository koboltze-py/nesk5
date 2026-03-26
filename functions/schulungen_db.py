"""
Schulungen-Datenbank
Verwaltung von Mitarbeiter-Schulungseinträgen (Schulungsart, Datum, Gültig bis, Status).
"""
import sqlite3
from pathlib import Path
from datetime import datetime
from config import BASE_DIR as _BASE_DIR

_DB_PFAD = Path(_BASE_DIR) / "database SQL" / "schulungen.db"

SCHULUNGSARTEN = [
    "Erste Hilfe",
    "Brandschutz",
    "Hygieneunterweisung",
    "Arbeitssicherheit",
    "Datenschutz",
    "Notfallsanitäter-Fortbildung",
    "Rettungssanitäter-Fortbildung",
    "Sonstiges",
]

STATUS_OPTIONEN = ["bestanden", "ausstehend", "abgebrochen", "abgelaufen"]


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(_DB_PFAD, timeout=5)
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous  = NORMAL")
    conn.execute("PRAGMA busy_timeout  = 5000")
    return conn


def _init_db():
    _DB_PFAD.parent.mkdir(parents=True, exist_ok=True)
    with _connect() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS schulungen (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            erstellt_am      TEXT NOT NULL,
            mitarbeiter      TEXT NOT NULL,
            schulungsart     TEXT NOT NULL,
            datum            TEXT NOT NULL,
            gueltig_bis      TEXT,
            status           TEXT NOT NULL DEFAULT 'bestanden',
            bemerkung        TEXT,
            aufgenommen_von  TEXT
        )
        """)
        conn.commit()


def schulung_speichern(daten: dict) -> int:
    _init_db()
    now = datetime.now().isoformat(timespec="seconds")
    with _connect() as conn:
        cur = conn.execute(
            """
            INSERT INTO schulungen
              (erstellt_am, mitarbeiter, schulungsart, datum, gueltig_bis, status, bemerkung, aufgenommen_von)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                now,
                daten.get("mitarbeiter", ""),
                daten.get("schulungsart", ""),
                daten.get("datum", ""),
                daten.get("gueltig_bis", ""),
                daten.get("status", "bestanden"),
                daten.get("bemerkung", ""),
                daten.get("aufgenommen_von", ""),
            ),
        )
        conn.commit()
        return cur.lastrowid


def schulung_aktualisieren(schulung_id: int, daten: dict) -> None:
    _init_db()
    with _connect() as conn:
        conn.execute(
            """
            UPDATE schulungen SET
              mitarbeiter=?, schulungsart=?, datum=?, gueltig_bis=?,
              status=?, bemerkung=?, aufgenommen_von=?
            WHERE id=?
            """,
            (
                daten.get("mitarbeiter", ""),
                daten.get("schulungsart", ""),
                daten.get("datum", ""),
                daten.get("gueltig_bis", ""),
                daten.get("status", "bestanden"),
                daten.get("bemerkung", ""),
                daten.get("aufgenommen_von", ""),
                schulung_id,
            ),
        )
        conn.commit()


def schulung_loeschen(schulung_id: int) -> None:
    _init_db()
    with _connect() as conn:
        conn.execute("DELETE FROM schulungen WHERE id=?", (schulung_id,))
        conn.commit()


def lade_schulungen(jahr: int | None = None, mitarbeiter: str | None = None) -> list[dict]:
    _init_db()
    sql = "SELECT * FROM schulungen WHERE 1=1"
    params: list = []
    if jahr:
        sql += " AND substr(datum, 1, 4) = ?"
        params.append(str(jahr))
    if mitarbeiter:
        sql += " AND mitarbeiter LIKE ?"
        params.append(f"%{mitarbeiter}%")
    sql += " ORDER BY datum DESC"
    with _connect() as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows]


def lade_jahre() -> list[int]:
    _init_db()
    with _connect() as conn:
        rows = conn.execute(
            "SELECT DISTINCT substr(datum,1,4) AS j FROM schulungen ORDER BY j DESC"
        ).fetchall()
    return [int(r[0]) for r in rows if r[0] and r[0].isdigit()]
