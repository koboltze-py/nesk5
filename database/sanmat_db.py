"""
Sanitätsmaterial-Datenbank für Nesk3
Adaptiert aus DRK Sanmat – verwendet SANMAT_DB_PATH aus config.py
"""

import sqlite3
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import SANMAT_DB_PATH

_CREATE_TABLES_SQL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS artikel (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    artikelnr       TEXT    DEFAULT '',
    bezeichnung     TEXT    NOT NULL,
    kategorie       TEXT    DEFAULT '',
    einheit         TEXT    DEFAULT 'Stück',
    packungsinhalt  TEXT    DEFAULT '',
    hersteller      TEXT    DEFAULT 'meetB',
    pzn             TEXT    DEFAULT '',
    aktiv           INTEGER DEFAULT 1,
    bemerkung       TEXT    DEFAULT '',
    erstellt_am     TEXT    DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS bestand (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    artikel_id      INTEGER NOT NULL REFERENCES artikel(id) ON DELETE CASCADE,
    menge           INTEGER NOT NULL DEFAULT 0,
    min_menge       INTEGER DEFAULT 0,
    lagerort        TEXT    DEFAULT '',
    bemerkung       TEXT    DEFAULT '',
    erstellt_am     TEXT    DEFAULT (datetime('now','localtime')),
    geaendert_am    TEXT    DEFAULT (datetime('now','localtime')),
    UNIQUE(artikel_id)
);

CREATE TABLE IF NOT EXISTS buchungen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    artikel_id      INTEGER REFERENCES artikel(id),
    artikel_name    TEXT    NOT NULL,
    menge           INTEGER NOT NULL,
    typ             TEXT    NOT NULL,
    von             TEXT    DEFAULT '',
    bemerkung       TEXT    DEFAULT '',
    datum           TEXT    NOT NULL,
    erstellt_am     TEXT    DEFAULT (datetime('now','localtime'))
);

CREATE INDEX IF NOT EXISTS idx_bestand_artikel   ON bestand(artikel_id);
CREATE INDEX IF NOT EXISTS idx_buchungen_datum   ON buchungen(datum);
CREATE INDEX IF NOT EXISTS idx_buchungen_typ     ON buchungen(typ);
CREATE INDEX IF NOT EXISTS idx_artikel_kategorie ON artikel(kategorie);
"""

# Vordefinierte Artikel (meetB-Lieferscheine)
_INITIAL_ARTIKEL = [
    ("600400",          "Sofortkältekompresse 15 x 21 cm",                        "Wundversorgung",       "Pck.",  "10 Stück",   ""),
    ("1541 P",          "Steri-Strip 6 x 75 mm",                                  "Wundversorgung",       "Pck.",  "36 Streifen","PZN 03328876"),
    ("9422041",         "Nitril Handschuh Peha-Soft guard XL (lange Stulpe)",      "Schutzausrüstung",     "Pck.",  "100 Stück",  "PZN 03539202"),
    ("9422031",         "Nitril Handschuh Peha-Soft guard L (lange Stulpe)",       "Schutzausrüstung",     "Pck.",  "100 Stück",  "PZN 03539194"),
    ("9422021",         "Nitril Handschuh Peha-Soft guard M (lange Stulpe)",       "Schutzausrüstung",     "Pck.",  "100 Stück",  "PZN 03539188"),
    ("9422011",         "Nitril Handschuh Peha-Soft guard S (lange Stulpe)",       "Schutzausrüstung",     "Pck.",  "100 Stück",  "PZN 03539171"),
    ("920313",          "Alphacheck professional Blutzuckerteststreifen",           "Diagnostik",           "Pck.",  "50 Stück",   "PZN 10329014"),
    ("35408",           "BaSick Bag Spuckbeutel 1500 ml",                          "Patientenversorgung",  "Pck.",  "25 Stück",   "PZN 13818825"),
    ("00-323DS-T060",   "Descosept Sensitive Wipes (20x22 cm)",                    "Desinfektion",         "Pck.",  "60 Tücher",  ""),
    ("00-323DS-OSEB120","Descosept Sensitive Wipes XL Standbeutel (17,5x36 cm)",   "Desinfektion",         "Pck.",  "120 Blatt",  ""),
    ("12562",           "Haft-Fixierbinde 4 cm x 4 m",                            "Wundversorgung",       "Stück", "",           "PZN 04019746"),
    ("12565",           "Haft-Fixierbinde 8 cm x 4 m",                            "Wundversorgung",       "Stück", "",           "PZN 01329067"),
    ("101340",          "Schutzkappen für Thermo Scan 3000/4000/6000",             "Diagnostik",           "Pck.",  "2x20 Stück", "PZN 07437651"),
    ("053030",          "Mulltupfer steril 2x2 Stück (20x20 cm)",                  "Wundversorgung",       "Pck.",  "50 Sets",    "PZN 01364129"),
    ("106600",          "Sterillium Händedesinfektionsmittel 1 Liter",             "Desinfektion",         "Fl.",   "1 Liter",    "PZN 01494079"),
    ("9566",            "Leukosilk Rollenpflaster 1,25 cm x 9,2 m",               "Wundversorgung",       "Rolle", "",           "PZN 04593675"),
    ("9567",            "Leukosilk Rollenpflaster 2,5 cm x 9,2 m",                "Wundversorgung",       "Rolle", "",           "PZN 04593681"),
    ("1222233",         "Laborbecher 0,2 Liter weiß",                              "Verbrauchsmaterial",   "Pck.",  "100 Stück",  ""),
    ("SP-00-S",         "EKG Klebeelektrode Blue Sensor SP-00-S 38mm",             "Diagnostik",           "Pck.",  "50 Stück",   ""),
    ("976802",          "Cutasept F Hautdesinfektion Sprühflasche 250 ml",          "Desinfektion",         "Fl.",   "250 ml",     "PZN 03917271"),
    ("9085",            "Fixomull stretch 10 m x 10 cm",                           "Wundversorgung",       "Stück", "",           "PZN 04539523"),
    ("325012000",       "Ambu SPUR II Beatmungsbeutel Erwachsene",                 "Notfallausrüstung",    "Stück", "",           ""),
    ("719060",          "Lindesa Hautschutzcreme 50 ml",                           "Schutzausrüstung",     "Stück", "50 ml",      "PZN 1281030"),
    ("10670",           "Hansaplast Kinderpflaster Disney Mickey (20 Strips)",      "Wundversorgung",       "Pck.",  "20 Strips",  "PZN 16760150"),
    ("9193006",         "Multi Safe med 6 Kanülensammler ca. 5,1 Liter",           "Entsorgung",           "Stück", "5,1 Liter",  ""),
    ("202210401",       "Infusionstasche S PAX-Light rot (11x25x12 cm)",            "Notfallausrüstung",    "Stück", "",           ""),
    ("22011",           "Spritze Injekt 1 ml 2-teilig ohne Kanüle",                "Verbrauchsmaterial",   "Pck.",  "100 Stück",  "PZN 00896456"),
    ("014211",          "Saugkompresse steril 10x20 cm",                           "Wundversorgung",       "Pck.",  "25 Stück",   "PZN 11606013"),
    ("1003373",         "Aluderm Verbandpäckchen groß 4m x 10cm (Kompr. 10x12cm)", "Wundversorgung",       "Stück", "",           "PZN 03147525"),
    ("01003371",        "Aluderm Verbandpäckchen klein 3m x 6cm (Kompr. 6x8cm)",   "Wundversorgung",       "Stück", "",           ""),
    ("1003372",         "Aluderm Verbandpäckchen mittel 4m x 8cm (Kompr. 8x10cm)", "Wundversorgung",       "Stück", "",           "PZN 03147519"),
    ("10330-2",         "SAM SPLINT-Fingerschiene",                                 "Notfallausrüstung",    "Stück", "",           ""),
    ("10330-1",         "SAM SPLINT Standard 11x92 cm gerollt",                     "Notfallausrüstung",    "Stück", "",           ""),
    ("4063000-100",     "Infusionssystem Intrafix SafeSet 1,8m AirStop",            "Verbrauchsmaterial",   "Kin.",  "100 Stück",  "PZN 01900697"),
    ("973389",          "Bacillol AF 5 Liter Kanister",                             "Desinfektion",         "Kan.",  "5 Liter",    "PZN 00182685"),
    ("1024",            "Leukosilk Rollenpflaster 5 cm x 5 m",                      "Wundversorgung",       "Rolle", "",           "PZN 00626231"),
    ("1022",            "Leukosilk Rollenpflaster 2,5 cm x 5 m",                    "Wundversorgung",       "Rolle", "",           "PZN 00626225"),
    ("672700",          "Bode Messbecher 250 ml (Desinfektionsherstellung)",         "Desinfektion",         "Stück", "250 ml",     "PZN 03650951"),
    ("09160515",        "Alkoholtupfer 6x3 cm gefaltet einzeln verpackt",           "Desinfektion",         "Pck.",  "100 Stück",  "PZN 08468837"),
    ("1003208",         "aluderm Kompresse 10 x 10 cm",                             "Wundversorgung",       "Pck.",  "",           ""),
    ("072585",          "Wundverband steril 7x5 cm",                                "Wundversorgung",       "Pck.",  "50 Stück",   "PZN 07092666"),
    ("40156",           "Wundpflaster 6 cm x 5 m im Spenderkarton",                "Wundversorgung",       "Pck.",  "",           "PZN 04002852"),
    ("1009152",         "aluderm-aluplast Sortiment klein Fingerverband",            "Wundversorgung",       "Pck.",  "",           ""),
    ("1009163",         "aluderm Fingerverband 4x2 cm",                             "Wundversorgung",       "Pck.",  "10 Stück",   ""),
    ("1009184",         "aluderm Fingerkuppenverband 4,3x7,2 cm",                   "Wundversorgung",       "Pck.",  "10 Stück",   ""),
    ("35406",           "Auffangbeutel Universal mit Verschlussring",               "Patientenversorgung",  "Pck.",  "50 Stück",   ""),
    ("70003349",        "Schülke Wipes Wischtücher safe & easy",                    "Desinfektion",         "Karton","6x111 Tücher","PZN 18050464"),
    ("84156",           "Universal Wischtücher blau/weiß 30x33 cm",                "Verbrauchsmaterial",   "Pck.",  "50 Stück",   ""),
    ("16018182",        "Absaugschlauch mit Fingertip und Trichter CH25 (2,0 m)",   "Notfallausrüstung",    "Stück", "",           ""),
    ("0704908210",      "Yankauer Saugansatz CH24 abgewinkelt steril",              "Notfallausrüstung",    "Stück", "",           ""),
    ("1453",            "Cirrus™2 Verneblerset Erwachsene mit EcoLite™ Maske 2,1m", "Notfallausrüstung",   "Stück", "",           ""),
    ("60501502",        "Gänsegurgel gerade 25 cm mit Doppel-Drehkonnektor",        "Notfallausrüstung",    "Stück", "",           ""),
    ("5402822",         "Sauerstoffmaske Erwachsene mit Reservoirbeutel 2m Schlauch","Notfallausrüstung",   "Stück", "",           ""),
    ("000252956",       "Beatmungsmaske Ambu Plus 6 Erwachsene groß",               "Notfallausrüstung",    "Stück", "",           ""),
    ("5402880",         "Hyperventilationsmaske mit Rückatembeutel",                "Notfallausrüstung",    "Stück", "",           ""),
    ("MAD100",          "Medikamentenvernebler MAD 100 Nasalzerstäuber",            "Notfallausrüstung",    "Stück", "",           "PZN 10134233"),
    ("193880",          "Kanülenabwurfbehälter Multi Safe Sani 200 (0,2 Liter)",    "Entsorgung",           "Stück", "0,2 Liter",  ""),
    ("79401805",        "Kombistopfen blau Luer-Lock",                              "Verbrauchsmaterial",   "Pck.",  "100 Stück",  ""),
    ("716270",          "EKG Papier für Corpuls C II (106 mm x 22 m)",              "Diagnostik",           "Stück", "",           ""),
    ("121402",          "Octenisept 500 ml Schleimhautantiseptikum",                "Desinfektion",         "Fl.",   "500 ml",     ""),
    ("972553",          "Baktolan balm Hautschutzcreme 350 ml",                     "Schutzausrüstung",     "Fl.",   "350 ml",     ""),
]


class SanmatDB:
    """Sanitätsmaterial-Datenbankzugriff für Nesk3."""

    def __init__(self):
        self.sanmat_db = SANMAT_DB_PATH

    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.sanmat_db)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def initialize(self):
        os.makedirs(os.path.dirname(self.sanmat_db), exist_ok=True)
        conn = sqlite3.connect(self.sanmat_db)
        conn.executescript(_CREATE_TABLES_SQL)
        conn.commit()
        conn.close()

    def sanmat_db_exists(self) -> bool:
        return os.path.exists(self.sanmat_db)

    def hat_artikel(self) -> bool:
        conn = self._conn()
        count = conn.execute("SELECT COUNT(*) FROM artikel").fetchone()[0]
        conn.close()
        return count > 0

    def upsert_initial_artikel(self) -> int:
        conn = self._conn()
        count = 0
        for (artikelnr, bezeichnung, kategorie, einheit, packungsinhalt, pzn) in _INITIAL_ARTIKEL:
            exists = conn.execute(
                "SELECT id FROM artikel WHERE artikelnr = ?", (artikelnr,)
            ).fetchone()
            if exists:
                continue
            cur = conn.execute(
                "INSERT INTO artikel (artikelnr, bezeichnung, kategorie, einheit, packungsinhalt, pzn) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (artikelnr, bezeichnung, kategorie, einheit, packungsinhalt, pzn)
            )
            artikel_id = cur.lastrowid
            conn.execute("INSERT INTO bestand (artikel_id, menge, min_menge) VALUES (?, 0, 0)", (artikel_id,))
            count += 1
        conn.commit()
        conn.close()
        return count

    # ── Artikel ──────────────────────────────────────────────────────────

    def get_artikel(self, nur_aktiv: bool = True) -> list[dict]:
        conn = self._conn()
        where = "WHERE a.aktiv = 1" if nur_aktiv else ""
        cur = conn.execute(f"""
            SELECT a.id, a.artikelnr, a.bezeichnung, a.kategorie, a.einheit,
                   a.packungsinhalt, a.hersteller, a.pzn, a.aktiv, a.bemerkung,
                   COALESCE(b.menge, 0) AS menge,
                   COALESCE(b.min_menge, 0) AS min_menge,
                   COALESCE(b.lagerort, '') AS lagerort
            FROM artikel a
            LEFT JOIN bestand b ON b.artikel_id = a.id
            {where}
            ORDER BY a.kategorie, a.bezeichnung
        """)
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def get_artikel_by_id(self, aid: int) -> dict | None:
        conn = self._conn()
        cur = conn.execute("""
            SELECT a.id, a.artikelnr, a.bezeichnung, a.kategorie, a.einheit,
                   a.packungsinhalt, a.hersteller, a.pzn, a.aktiv, a.bemerkung,
                   COALESCE(b.menge, 0) AS menge,
                   COALESCE(b.min_menge, 0) AS min_menge,
                   COALESCE(b.lagerort, '') AS lagerort
            FROM artikel a
            LEFT JOIN bestand b ON b.artikel_id = a.id
            WHERE a.id = ?
        """, (aid,))
        row = cur.fetchone()
        conn.close()
        return dict(row) if row else None

    def add_artikel(self, bezeichnung: str, artikelnr: str = "", kategorie: str = "",
                    einheit: str = "Stück", packungsinhalt: str = "", hersteller: str = "meetB",
                    pzn: str = "", bemerkung: str = "") -> tuple[int, str]:
        conn = self._conn()
        try:
            cur = conn.execute(
                "INSERT INTO artikel (artikelnr, bezeichnung, kategorie, einheit, "
                "packungsinhalt, hersteller, pzn, bemerkung) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (artikelnr, bezeichnung, kategorie, einheit, packungsinhalt, hersteller, pzn, bemerkung)
            )
            artikel_id = cur.lastrowid
            conn.execute("INSERT INTO bestand (artikel_id, menge, min_menge) VALUES (?, 0, 0)", (artikel_id,))
            conn.commit()
            return artikel_id, f"Artikel '{bezeichnung}' wurde angelegt."
        except Exception as e:
            return 0, str(e)
        finally:
            conn.close()

    def update_artikel(self, aid: int, bezeichnung: str, artikelnr: str = "",
                       kategorie: str = "", einheit: str = "Stück", packungsinhalt: str = "",
                       hersteller: str = "meetB", pzn: str = "", bemerkung: str = "") -> tuple[bool, str]:
        conn = self._conn()
        try:
            conn.execute(
                "UPDATE artikel SET artikelnr=?, bezeichnung=?, kategorie=?, einheit=?, "
                "packungsinhalt=?, hersteller=?, pzn=?, bemerkung=? WHERE id=?",
                (artikelnr, bezeichnung, kategorie, einheit, packungsinhalt, hersteller, pzn, bemerkung, aid)
            )
            conn.commit()
            return True, "Artikel aktualisiert."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def deactivate_artikel(self, aid: int) -> tuple[bool, str]:
        conn = self._conn()
        try:
            conn.execute("UPDATE artikel SET aktiv=0 WHERE id=?", (aid,))
            conn.commit()
            return True, "Artikel deaktiviert."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def get_kategorien(self) -> list[str]:
        conn = self._conn()
        cur = conn.execute(
            "SELECT DISTINCT kategorie FROM artikel WHERE aktiv=1 AND kategorie!='' ORDER BY kategorie"
        )
        cats = [r[0] for r in cur.fetchall()]
        conn.close()
        return cats

    # ── Bestand ──────────────────────────────────────────────────────────

    def get_bestand(self) -> list[dict]:
        return self.get_artikel(nur_aktiv=True)

    def get_niedrig_bestand(self) -> list[dict]:
        conn = self._conn()
        cur = conn.execute("""
            SELECT a.id, a.bezeichnung, a.einheit, a.kategorie,
                   COALESCE(b.menge, 0) AS menge,
                   COALESCE(b.min_menge, 0) AS min_menge,
                   COALESCE(b.lagerort, '') AS lagerort
            FROM artikel a
            LEFT JOIN bestand b ON b.artikel_id = a.id
            WHERE a.aktiv = 1
              AND COALESCE(b.min_menge, 0) > 0
              AND COALESCE(b.menge, 0) <= COALESCE(b.min_menge, 0)
            ORDER BY COALESCE(b.menge, 0) ASC, a.bezeichnung
        """)
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def update_bestand(self, artikel_id: int, menge: int, min_menge: int,
                       lagerort: str = "", bemerkung: str = "") -> tuple[bool, str]:
        conn = self._conn()
        try:
            conn.execute("""
                INSERT INTO bestand (artikel_id, menge, min_menge, lagerort, bemerkung, geaendert_am)
                VALUES (?, ?, ?, ?, ?, datetime('now','localtime'))
                ON CONFLICT(artikel_id) DO UPDATE SET
                    menge=excluded.menge,
                    min_menge=excluded.min_menge,
                    lagerort=excluded.lagerort,
                    bemerkung=excluded.bemerkung,
                    geaendert_am=datetime('now','localtime')
            """, (artikel_id, menge, min_menge, lagerort, bemerkung))
            conn.commit()
            return True, "Bestand aktualisiert."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def einlagern(self, artikel_id: int, artikel_name: str, menge: int,
                  datum: str, von: str = "", bemerkung: str = "",
                  typ: str = "einlagerung") -> tuple[bool, str]:
        conn = self._conn()
        try:
            conn.execute("""
                UPDATE bestand SET menge = menge + ?, geaendert_am = datetime('now','localtime')
                WHERE artikel_id = ?
            """, (menge, artikel_id))
            conn.execute(
                "INSERT INTO buchungen (artikel_id, artikel_name, menge, typ, von, bemerkung, datum) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (artikel_id, artikel_name, menge, typ, von, bemerkung, datum)
            )
            conn.commit()
            return True, f"+{menge} {artikel_name} eingelagert."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def entnehmen(self, artikel_id: int, artikel_name: str, menge: int,
                  datum: str, typ: str = "entnahme", von: str = "",
                  bemerkung: str = "", negativ_erlaubt: bool = False) -> tuple[bool, str]:
        conn = self._conn()
        try:
            cur = conn.execute("SELECT menge FROM bestand WHERE artikel_id = ?", (artikel_id,))
            row = cur.fetchone()
            aktuell = row["menge"] if row else 0
            if not negativ_erlaubt and menge > aktuell:
                conn.close()
                return False, f"Nicht genug auf Lager (Bestand: {aktuell})."
            conn.execute("""
                UPDATE bestand SET menge = menge - ?, geaendert_am = datetime('now','localtime')
                WHERE artikel_id = ?
            """, (menge, artikel_id))
            conn.execute(
                "INSERT INTO buchungen (artikel_id, artikel_name, menge, typ, von, bemerkung, datum) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (artikel_id, artikel_name, -menge, typ, von, bemerkung, datum)
            )
            conn.commit()
            return True, f"-{menge} {artikel_name} gebucht."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def korrektur(self, artikel_id: int, artikel_name: str, neue_menge: int,
                  datum: str, von: str = "", bemerkung: str = "") -> tuple[bool, str]:
        conn = self._conn()
        try:
            cur = conn.execute("SELECT menge FROM bestand WHERE artikel_id = ?", (artikel_id,))
            row = cur.fetchone()
            alt = row["menge"] if row else 0
            diff = neue_menge - alt
            conn.execute(
                "UPDATE bestand SET menge=?, geaendert_am=datetime('now','localtime') WHERE artikel_id=?",
                (neue_menge, artikel_id)
            )
            bem = f"Korrektur: {alt}→{neue_menge}" + (f" | {bemerkung}" if bemerkung else "")
            conn.execute(
                "INSERT INTO buchungen (artikel_id, artikel_name, menge, typ, von, bemerkung, datum) "
                "VALUES (?, ?, ?, 'korrektur', ?, ?, ?)",
                (artikel_id, artikel_name, diff, von, bem, datum)
            )
            conn.commit()
            return True, f"Bestand korrigiert auf {neue_menge}."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    # ── Buchungen ─────────────────────────────────────────────────────────

    def get_buchungen(self, limit: int = 200, offset: int = 0,
                      artikel_id: int = None, typ: str = None,
                      datum_von: str = None, datum_bis: str = None,
                      suche: str = None) -> list[dict]:
        conn = self._conn()
        where_parts = []
        params: list = []
        if artikel_id:
            where_parts.append("b.artikel_id = ?")
            params.append(artikel_id)
        if typ and typ != "Alle":
            where_parts.append("b.typ = ?")
            params.append(typ)
        if datum_von:
            where_parts.append("b.datum >= ?")
            params.append(datum_von)
        if datum_bis:
            where_parts.append("b.datum <= ?")
            params.append(datum_bis)
        if suche:
            where_parts.append("(b.artikel_name LIKE ? OR b.von LIKE ? OR b.bemerkung LIKE ?)")
            params.extend([f"%{suche}%", f"%{suche}%", f"%{suche}%"])
        where = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""
        cur = conn.execute(f"""
            SELECT b.id, b.artikel_name, b.menge, b.typ, b.von, b.bemerkung, b.datum,
                   b.erstellt_am, b.artikel_id
            FROM buchungen b
            {where}
            ORDER BY b.datum DESC, b.erstellt_am DESC
            LIMIT ? OFFSET ?
        """, params + [limit, offset])
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def count_buchungen(self, artikel_id: int = None, typ: str = None,
                        datum_von: str = None, datum_bis: str = None,
                        suche: str = None) -> int:
        conn = self._conn()
        where_parts = []
        params: list = []
        if artikel_id:
            where_parts.append("artikel_id = ?")
            params.append(artikel_id)
        if typ and typ != "Alle":
            where_parts.append("typ = ?")
            params.append(typ)
        if datum_von:
            where_parts.append("datum >= ?")
            params.append(datum_von)
        if datum_bis:
            where_parts.append("datum <= ?")
            params.append(datum_bis)
        if suche:
            where_parts.append("(artikel_name LIKE ? OR von LIKE ? OR bemerkung LIKE ?)")
            params.extend([f"%{suche}%", f"%{suche}%", f"%{suche}%"])
        where = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""
        cur = conn.execute(f"SELECT COUNT(*) FROM buchungen {where}", params)
        count = cur.fetchone()[0]
        conn.close()
        return count

    def buche_verbrauch_gruppe(self, stichwort: str, artikel_liste: list[dict],
                               datum: str, entnehmer: str = "",
                               negativ_erlaubt: bool = False) -> tuple[bool, str]:
        """Bucht mehrere Artikel als manuelle Verbrauchsgruppe (GID = Zeitstempel)."""
        import time
        gid = int(time.time())
        bemerkung = f"{stichwort}  (GID {gid})"
        conn = self._conn()
        try:
            fehler = []
            for pos in artikel_liste:
                art_id   = pos["artikel_id"]
                art_name = pos["bezeichnung"]
                menge    = int(pos["menge"])
                cur = conn.execute("SELECT menge FROM bestand WHERE artikel_id = ?", (art_id,))
                row = cur.fetchone()
                aktuell = row["menge"] if row else 0
                if not negativ_erlaubt and menge > aktuell:
                    fehler.append(f"{art_name}: Nicht genug Bestand ({aktuell})")
                    continue
                conn.execute("""
                    UPDATE bestand SET menge = menge - ?, geaendert_am = datetime('now','localtime')
                    WHERE artikel_id = ?
                """, (menge, art_id))
                conn.execute(
                    "INSERT INTO buchungen (artikel_id, artikel_name, menge, typ, von, bemerkung, datum) "
                    "VALUES (?, ?, ?, 'verbrauch', ?, ?, ?)",
                    (art_id, art_name, -menge, entnehmer, bemerkung, datum)
                )
            if fehler:
                conn.rollback()
                return False, "\n".join(fehler)
            conn.commit()
            return True, f"{len(artikel_liste)} Artikel als Verbrauchsgruppe gebucht."
        except Exception as e:
            conn.rollback()
            return False, str(e)
        finally:
            conn.close()

    def get_buchung_by_id(self, bid: int) -> dict | None:
        conn = self._conn()
        cur = conn.execute(
            "SELECT id, artikel_id, artikel_name, menge, typ, von, bemerkung, datum FROM buchungen WHERE id=?",
            (bid,)
        )
        row = cur.fetchone()
        conn.close()
        return dict(row) if row else None

    def delete_buchung(self, bid: int) -> tuple[bool, str]:
        """Löscht eine Buchung und stellt den Bestand wieder her."""
        conn = self._conn()
        try:
            cur = conn.execute(
                "SELECT artikel_id, menge FROM buchungen WHERE id=?", (bid,)
            )
            row = cur.fetchone()
            if not row:
                conn.close()
                return False, "Buchung nicht gefunden."
            artikel_id, menge = row["artikel_id"], row["menge"]
            # Bestand rückgängig machen (menge war negativ bei verbrauch)
            conn.execute(
                "UPDATE bestand SET menge = menge - ?, geaendert_am = datetime('now','localtime') WHERE artikel_id = ?",
                (menge, artikel_id)
            )
            conn.execute("DELETE FROM buchungen WHERE id=?", (bid,))
            conn.commit()
            return True, "Buchung gelöscht."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def update_buchung(self, bid: int, neue_menge: int, von: str,
                       bemerkung: str, datum: str) -> tuple[bool, str]:
        """Ändert Menge/Von/Bemerkung/Datum und passt Bestand an."""
        conn = self._conn()
        try:
            cur = conn.execute(
                "SELECT artikel_id, menge FROM buchungen WHERE id=?", (bid,)
            )
            row = cur.fetchone()
            if not row:
                conn.close()
                return False, "Buchung nicht gefunden."
            artikel_id, alte_menge = row["artikel_id"], row["menge"]
            # alte_menge ist negativ (verbrauch), neue_menge ist positiv (Eingabe)
            neue_menge_db = -abs(neue_menge)
            diff = neue_menge_db - alte_menge  # wie viel sich der Bestand ändert
            conn.execute(
                "UPDATE bestand SET menge = menge - ?, geaendert_am = datetime('now','localtime') WHERE artikel_id = ?",
                (diff, artikel_id)
            )
            conn.execute(
                "UPDATE buchungen SET menge=?, von=?, bemerkung=?, datum=? WHERE id=?",
                (neue_menge_db, von, bemerkung, datum, bid)
            )
            conn.commit()
            return True, "Buchung aktualisiert."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def restore_buchung(self, snapshot: dict) -> tuple[bool, str]:
        """Stellt eine gelöschte Buchung aus einem Snapshot wieder her."""
        conn = self._conn()
        try:
            conn.execute("""
                INSERT INTO buchungen (id, artikel_id, artikel_name, menge, typ, von, bemerkung, datum)
                VALUES (:id, :artikel_id, :artikel_name, :menge, :typ, :von, :bemerkung, :datum)
            """, snapshot)
            # Bestand zurücksetzen
            conn.execute(
                "UPDATE bestand SET menge = menge + ?, geaendert_am = datetime('now','localtime') WHERE artikel_id = ?",
                (snapshot["menge"], snapshot["artikel_id"])
            )
            conn.commit()
            return True, "Buchung wiederhergestellt."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def get_statistik(self) -> dict:
        conn = self._conn()
        art_count = conn.execute("SELECT COUNT(*) FROM artikel WHERE aktiv=1").fetchone()[0]
        niedrig = conn.execute("""
            SELECT COUNT(*) FROM artikel a
            LEFT JOIN bestand b ON b.artikel_id=a.id
            WHERE a.aktiv=1 AND COALESCE(b.min_menge,0)>0
              AND COALESCE(b.menge,0) <= COALESCE(b.min_menge,0)
        """).fetchone()[0]
        leer = conn.execute("""
            SELECT COUNT(*) FROM artikel a
            LEFT JOIN bestand b ON b.artikel_id=a.id
            WHERE a.aktiv=1 AND COALESCE(b.menge,0)=0
        """).fetchone()[0]
        buchungen_heute = conn.execute(
            "SELECT COUNT(*) FROM buchungen WHERE datum=date('now','localtime')"
        ).fetchone()[0]
        conn.close()
        return {
            "artikel_gesamt": art_count,
            "niedrig_bestand": niedrig,
            "leer": leer,
            "buchungen_heute": buchungen_heute,
        }

    def set_default_min_menge(self, min_menge: int) -> int:
        conn = self._conn()
        cur = conn.execute("UPDATE bestand SET min_menge=? WHERE min_menge=0", (min_menge,))
        count = cur.rowcount
        conn.commit()
        conn.close()
        return count
