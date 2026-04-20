"""SQLite-CRUD für persönliche Dashboard-Notizen."""
import sqlite3
from datetime import datetime, timedelta


def _db_path() -> str:
    from config import NOTIZEN_DB_PATH
    return NOTIZEN_DB_PATH


def _get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _init_db():
    with _get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS notizen (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                titel       TEXT    NOT NULL,
                text        TEXT    NOT NULL DEFAULT '',
                datum       TEXT    NOT NULL,
                erstellt_am TEXT    NOT NULL,
                status      TEXT    NOT NULL DEFAULT 'offen'
            )
        """)
        conn.commit()


def speichern(titel: str, text: str = "", datum: str = "") -> int:
    """Neue Notiz anlegen. datum im Format dd.MM.yyyy."""
    _init_db()
    if not datum:
        datum = datetime.today().strftime("%d.%m.%Y")
    jetzt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO notizen (titel, text, datum, erstellt_am, status) "
            "VALUES (?, ?, ?, ?, 'offen')",
            (titel, text, datum, jetzt),
        )
        conn.commit()
        return cur.lastrowid


def als_gelesen(nid: int):
    """Status auf 'gelesen' setzen."""
    _init_db()
    with _get_conn() as conn:
        conn.execute("UPDATE notizen SET status='gelesen' WHERE id=?", (nid,))
        conn.commit()


def als_erledigt(nid: int):
    """Status auf 'erledigt' setzen (Notiz verschwindet nach 5 Tagen nicht mehr)."""
    _init_db()
    with _get_conn() as conn:
        conn.execute("UPDATE notizen SET status='erledigt' WHERE id=?", (nid,))
        conn.commit()


def loeschen(nid: int):
    """Notiz dauerhaft löschen."""
    _init_db()
    with _get_conn() as conn:
        conn.execute("DELETE FROM notizen WHERE id=?", (nid,))
        conn.commit()


def lade_aktive() -> list[dict]:
    """
    Notizen der letzten 5 Tage (nach erstellt_am).
    Erledigte werden ebenfalls angezeigt (als durchgestrichen / grau).
    """
    _init_db()
    grenze = (datetime.now() - timedelta(days=5)).strftime("%Y-%m-%d %H:%M:%S")
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM notizen WHERE erstellt_am >= ? ORDER BY erstellt_am DESC",
            (grenze,),
        ).fetchall()
    return [dict(r) for r in rows]


def lade_alle() -> list[dict]:
    """Alle Notizen, neueste zuerst."""
    _init_db()
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM notizen ORDER BY erstellt_am DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def lade_fuer_datum(datum_de: str) -> list[dict]:
    """Alle Notizen für ein bestimmtes Datum (Format dd.MM.yyyy)."""
    _init_db()
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM notizen WHERE datum=? ORDER BY erstellt_am DESC",
            (datum_de,),
        ).fetchall()
    return [dict(r) for r in rows]
