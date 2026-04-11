"""
SQLite Datenbankverbindung
Verwaltet Verbindungen zur Nesk3 SQLite-Datenbank.
WAL-Modus für sicheres gleichzeitiges Lesen von mehreren PCs.
"""
import sqlite3
from contextlib import contextmanager
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import DB_PATH, MITARBEITER_DB_PATH


def _row_factory(cursor: sqlite3.Cursor, row: tuple) -> dict:
    """Gibt jede Zeile als dict zurück (Spaltenname → Wert)."""
    cols = [c[0] for c in cursor.description]
    return dict(zip(cols, row))


def get_connection() -> sqlite3.Connection:
    """Gibt eine neue SQLite-Verbindung zurück (WAL-Modus, dict-Zeilen)."""
    conn = sqlite3.connect(DB_PATH, timeout=3, check_same_thread=False)
    conn.row_factory = _row_factory
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 5000")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn


def test_connection() -> tuple[bool, str]:
    """
    Testet die Datenbankverbindung.
    Gibt (True, Info) oder (False, Fehlermeldung) zurück.
    """
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT sqlite_version()")
        row = cur.fetchone()
        version = row["sqlite_version()"] if row else "?"
        conn.close()
        return True, f"SQLite {version}  |  {DB_PATH}"
    except Exception as e:
        return False, str(e)


@contextmanager
def db_cursor(commit: bool = False):
    """
    Kontextmanager für DB-Cursor mit automatischem Commit/Rollback.
    Cursor liefert Zeilen als dict.

    Verwendung:
        with db_cursor(commit=True) as cur:
            cur.execute("INSERT ...")
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        yield cur
        if commit:
            conn.commit()
    except Exception:
        if conn:
            conn.rollback()
        raise
    finally:
        if conn:
            conn.close()


# ── Mitarbeiter-Datenbank (database SQL/mitarbeiter.db) ───────────────────────

_MA_SCHEMA = """
CREATE TABLE IF NOT EXISTS positionen (
    id    INTEGER PRIMARY KEY AUTOINCREMENT,
    name  TEXT UNIQUE NOT NULL
);
INSERT OR IGNORE INTO positionen(name) VALUES
    ('Notfallsanitäter'),('Rettungssanitäter'),('Sanitätshelfer'),
    ('Arzt'),('Verwaltung'),('Führungskraft');

CREATE TABLE IF NOT EXISTS abteilungen (
    id    INTEGER PRIMARY KEY AUTOINCREMENT,
    name  TEXT UNIQUE NOT NULL
);
INSERT OR IGNORE INTO abteilungen(name) VALUES
    ('Erste-Hilfe-Station'),('Sanitätsdienst'),('Verwaltung');

CREATE TABLE IF NOT EXISTS mitarbeiter (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    vorname         TEXT NOT NULL,
    nachname        TEXT NOT NULL,
    personalnummer  TEXT DEFAULT '',
    funktion        TEXT DEFAULT 'Schichtleiter' CHECK (funktion IN ('Schichtleiter','Dispo','Betreuer','stamm','dispo')),
    position        TEXT DEFAULT '',
    abteilung       TEXT DEFAULT '',
    email           TEXT DEFAULT '',
    telefon         TEXT DEFAULT '',
    eintrittsdatum  TEXT,
    status          TEXT DEFAULT 'aktiv',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime')),
    geaendert_am    TEXT DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_ma_nachname ON mitarbeiter(nachname);
CREATE INDEX IF NOT EXISTS idx_ma_funktion ON mitarbeiter(funktion);
"""


def get_ma_connection() -> sqlite3.Connection:
    """Gibt eine Verbindung zur Mitarbeiter-Datenbank zurück."""
    conn = sqlite3.connect(MITARBEITER_DB_PATH, timeout=3, check_same_thread=False)
    conn.row_factory = _row_factory
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 5000")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn


def init_mitarbeiter_db() -> None:
    """
    Erstellt die Mitarbeiter-Datenbank und Tabellen, falls noch nicht vorhanden.
    Wird beim App-Start aufgerufen.
    """
    conn = get_ma_connection()
    conn.executescript(_MA_SCHEMA)
    conn.commit()
    # Migration: CHECK-Constraint für funktion erweitern (SQLite: Tabelle neu erstellen)
    try:
        cols = [r[1] for r in conn.execute("PRAGMA table_info(mitarbeiter)").fetchall()]
        if cols:
            tbl_sql = conn.execute(
                "SELECT sql FROM sqlite_master WHERE type='table' AND name='mitarbeiter'"
            ).fetchone()
            if tbl_sql and "funktion IN ('stamm','dispo')" in (tbl_sql.get('sql') or tbl_sql[0] if isinstance(tbl_sql, tuple) else ''):
                conn.executescript("""
                    PRAGMA foreign_keys = OFF;
                    BEGIN;
                    CREATE TABLE IF NOT EXISTS mitarbeiter_new (
                        id              INTEGER PRIMARY KEY AUTOINCREMENT,
                        vorname         TEXT NOT NULL,
                        nachname        TEXT NOT NULL,
                        personalnummer  TEXT DEFAULT '',
                        funktion        TEXT DEFAULT 'Schichtleiter'
                                        CHECK (funktion IN ('Schichtleiter','Dispo','Betreuer','stamm','dispo')),
                        position        TEXT DEFAULT '',
                        abteilung       TEXT DEFAULT '',
                        email           TEXT DEFAULT '',
                        telefon         TEXT DEFAULT '',
                        eintrittsdatum  TEXT,
                        status          TEXT DEFAULT 'aktiv',
                        erstellt_am     TEXT DEFAULT (datetime('now','localtime')),
                        geaendert_am    TEXT DEFAULT (datetime('now','localtime'))
                    );
                    INSERT INTO mitarbeiter_new SELECT id,vorname,nachname,
                        COALESCE(personalnummer,''),
                        CASE funktion WHEN 'stamm' THEN 'Schichtleiter' WHEN 'dispo' THEN 'Dispo' ELSE funktion END,
                        COALESCE(position,''),COALESCE(abteilung,''),
                        COALESCE(email,''),COALESCE(telefon,''),
                        eintrittsdatum,COALESCE(status,'aktiv'),
                        erstellt_am,geaendert_am FROM mitarbeiter;
                    DROP TABLE mitarbeiter;
                    ALTER TABLE mitarbeiter_new RENAME TO mitarbeiter;
                    COMMIT;
                    PRAGMA foreign_keys = ON;
                """)
    except Exception:
        pass
    conn.close()


@contextmanager
def ma_db_cursor(commit: bool = False):
    """
    Kontextmanager für die Mitarbeiter-Datenbank (mitarbeiter.db).
    Verwendung identisch zu db_cursor().
    """
    conn = None
    try:
        conn = get_ma_connection()
        cur = conn.cursor()
        yield cur
        if commit:
            conn.commit()
    except Exception:
        if conn:
            conn.rollback()
        raise
    finally:
        if conn:
            conn.close()
