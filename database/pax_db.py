"""
PAX- und Einsatz-Datenbank-Funktionen
Speichert und liest tägliche Passagierzahlen und Einsatzzahlen.

WICHTIG: get_connection() liefert Zeilen als dict (Row Factory).
         Daher werden alle Spaltenwerte per Name (row['spaltenname']) abgerufen,
         NICHT per Index (row[0]).
"""
from .connection import get_connection


def speichere_tages_pax(datum: str, pax_zahl: int) -> None:
    """
    Speichert oder überschreibt die PAX-Zahl für ein Datum (YYYY-MM-DD).
    Bei erneutem Aufruf für das gleiche Datum wird der Wert aktualisiert.
    """
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO tages_pax (datum, pax_zahl)
        VALUES (?, ?)
        ON CONFLICT(datum) DO UPDATE SET
            pax_zahl   = excluded.pax_zahl,
            erfasst_am = datetime('now','localtime')
        """,
        (datum, pax_zahl),
    )
    conn.commit()


def lade_tages_pax(datum: str) -> int:
    """Gibt die gespeicherte PAX-Zahl für ein Datum zurück, oder 0."""
    conn = get_connection()
    row = conn.execute(
        "SELECT pax_zahl FROM tages_pax WHERE datum = ?", (datum,)
    ).fetchone()
    return int(row["pax_zahl"]) if row else 0


def lade_jahres_pax(jahr: int) -> int:
    """Gibt die Summe aller PAX-Zahlen für ein Jahr zurück."""
    conn = get_connection()
    row = conn.execute(
        "SELECT COALESCE(SUM(pax_zahl), 0) AS total FROM tages_pax"
        " WHERE strftime('%Y', datum) = ?",
        (str(jahr),),
    ).fetchone()
    return int(row["total"]) if row else 0


# ---------------------------------------------------------------------------
# Einsatz-Funktionen
# ---------------------------------------------------------------------------

def speichere_tages_einsaetze(datum: str, einsaetze_zahl: int) -> None:
    """
    Speichert oder ueberschreibt die Einsatz-Zahl fuer ein Datum (YYYY-MM-DD).
    """
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO tages_einsaetze (datum, einsaetze_zahl)
        VALUES (?, ?)
        ON CONFLICT(datum) DO UPDATE SET
            einsaetze_zahl = excluded.einsaetze_zahl,
            erfasst_am     = datetime('now','localtime')
        """,
        (datum, einsaetze_zahl),
    )
    conn.commit()


def lade_tages_einsaetze(datum: str) -> int:
    """Gibt die gespeicherte Einsatz-Zahl fuer ein Datum zurueck, oder 0."""
    conn = get_connection()
    row = conn.execute(
        "SELECT einsaetze_zahl FROM tages_einsaetze WHERE datum = ?", (datum,)
    ).fetchone()
    return int(row["einsaetze_zahl"]) if row else 0


def lade_jahres_einsaetze(jahr: int) -> int:
    """Gibt die Summe aller Einsatz-Zahlen fuer ein Jahr zurueck."""
    conn = get_connection()
    row = conn.execute(
        "SELECT COALESCE(SUM(einsaetze_zahl), 0) AS total FROM tages_einsaetze"
        " WHERE strftime('%Y', datum) = ?",
        (str(jahr),),
    ).fetchone()
    return int(row["total"]) if row else 0


def lade_alle_eintraege(jahr: int) -> list[dict]:
    """
    Gibt alle Eintraege (PAX + Einsaetze) fuer ein Jahr als Liste von Dicts zurueck.
    Jeder Eintrag: {'datum': 'YYYY-MM-DD', 'pax_zahl': int, 'einsaetze_zahl': int}
    Sortiert absteigend nach Datum.
    Enthaelt alle Daten aus beiden Tabellen (UNION der Daten).
    """
    conn = get_connection()
    rows = conn.execute(
        """
        SELECT
            dates.datum,
            COALESCE(p.pax_zahl, 0)       AS pax_zahl,
            COALESCE(e.einsaetze_zahl, 0) AS einsaetze_zahl
        FROM (
            SELECT datum FROM tages_pax        WHERE strftime('%Y', datum) = ?
            UNION
            SELECT datum FROM tages_einsaetze  WHERE strftime('%Y', datum) = ?
        ) AS dates
        LEFT JOIN tages_pax           p ON p.datum = dates.datum
        LEFT JOIN tages_einsaetze     e ON e.datum = dates.datum
        ORDER BY dates.datum DESC
        """,
        (str(jahr), str(jahr)),
    ).fetchall()
    return [{"datum": r["datum"], "pax_zahl": int(r["pax_zahl"]),
             "einsaetze_zahl": int(r["einsaetze_zahl"])} for r in rows]


def loesche_eintrag(datum: str) -> None:
    """Loescht PAX- und Einsatz-Eintrag fuer ein bestimmtes Datum."""
    conn = get_connection()
    conn.execute("DELETE FROM tages_pax WHERE datum = ?", (datum,))
    conn.execute("DELETE FROM tages_einsaetze WHERE datum = ?", (datum,))
    conn.commit()