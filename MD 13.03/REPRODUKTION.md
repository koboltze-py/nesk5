# Nesk3 – Reproduktionsprotokoll

**Stand:** 15.04.2026 – v3.9.0  
**Ziel:** Vollständige Neuerstellung der Nesk3-Anwendung auf einem neuen System

---

## Voraussetzungen

| Komponente | Version | Hinweis |
|-----------|---------|---------|
| Python | 3.13+ | App ist auf 3.13 entwickelt |
| PySide6 | aktuell | GUI-Framework |
| openpyxl | aktuell | Excel-Lesen/Schreiben |
| python-docx | aktuell | Word-Export |
| win32com (pywin32) | aktuell | Outlook-Integration (optional) |
| SQLite | 3.x | Eingebaut in Python |

```powershell
pip install PySide6 openpyxl python-docx pypdf pywin32
```

---

## 1. Projektstruktur anlegen

```
Nesk3/
├── main.py                          # Einstiegspunkt
├── config.py                        # Globale Konstanten (Pfade, Farben)
│
├── gui/
│   ├── __init__.py
│   ├── main_window.py               # Hauptfenster, 17-Seiten-Navigation
│   ├── dashboard.py                 # Dashboard (Statistik-Karten, Flugzeug-Animation)
│   ├── dienstplan.py                # Dienstplan (Excel-Import, Tabelle, Export)
│   ├── dienstliches.py              # Einsatzprotokoll + Übersicht [NEU v3.x]
│   ├── aufgaben.py                  # Aufgaben Nacht
│   ├── aufgaben_tag.py              # Aufgaben Tag (Code19Mail, FreieMail, Checklisten)
│   ├── sonderaufgaben.py            # Sonderaufgaben (Bulmor, E-Mobby)
│   ├── uebergabe.py                 # Übergabe-Protokoll
│   ├── fahrzeuge.py                 # Fahrzeugverwaltung
│   ├── code19.py                    # Code-19 (Taschenuhr-Animation, Protokoll)
│   ├── mitarbeiter.py               # Mitarbeiter-Verwaltung
│   ├── mitarbeiter_dokumente.py     # Dokumente + Stellungnahmen + Verspätung [NEU v3.x]
│   ├── einstellungen.py             # Einstellungen (Pfade, E-Mobby)
│   ├── checklisten.py               # Checklisten
│   ├── hilfe_dialog.py              # Animierter Hilfe-Dialog
│   ├── beschwerden.py               # Beschwerden-Verwaltung [NEU v3.5]
│   └── passagieranfragen.py         # Passagieranfragen (Outlook-Inbox, Extraktion, Szenarien) [NEU v3.5]
│
├── functions/
│   ├── __init__.py
│   ├── archiv_functions.py          # Archiv-DB (WAL) [NEU v3.x]
│   ├── dienstplan_html_export.py    # HTML-Dienstplan generieren [NEU v3.x]
│   ├── dienstplan_parser.py         # Excel-Parser (Krank-Typen, Dispo-Abschnitt)
│   ├── dienstplan_functions.py      # DB CRUD für Dienstplan
│   ├── emobby_functions.py          # E-Mobby-Fahrerliste (TXT↔DB-Sync)
│   ├── fahrzeug_functions.py        # DB CRUD für Fahrzeuge
│   ├── mail_functions.py            # Outlook-COM-Integration
│   ├── mitarbeiter_dokumente_functions.py  # Kategorien, Vorlagen [NEU v3.x]
│   ├── mitarbeiter_functions.py     # DB CRUD für Mitarbeiter
│   ├── settings_functions.py        # Key-Value-Einstellungen
│   ├── staerkemeldung_export.py     # Word-Export Stärkemeldung
│   ├── stellungnahmen_db.py         # Stellungnahmen-DB (WAL) [NEU v3.x]
│   ├── stellungnahmen_html_export.py # HTML-Ansicht Stellungnahmen [NEU v3.x]
│   ├── uebergabe_functions.py       # DB CRUD für Übergabe
│   ├── verspaetung_db.py            # Verspätungs-DB (WAL, _connect()) [NEU v3.x]
│   └── beschwerden_db.py            # Beschwerden-DB [NEU v3.5]
│
├── database/
│   ├── __init__.py
│   ├── connection.py                # SQLite-Verbindung (WAL)
│   ├── migrations.py               # DB-Migrationen (beim Start)
│   └── models.py                   # ORM-Modelle
│
├── backup/
│   ├── __init__.py
│   └── backup_manager.py            # ZIP Backup/Restore
│
├── Daten/
│   └── E-Mobby/
│       └── mobby.txt                ← E-Mobby-Fahrerliste (ein Name/Zeile)
│
├── database SQL/                    # Alle 5 SQLite-DBs (seit 05.03.2026)
│   ├── nesk3.db                     # Haupt-DB (WAL) ← wird automatisch erstellt
│   ├── archiv.db                    # Archiv (WAL)
│   ├── stellungnahmen.db            # Stellungnahmen (WAL)
│   ├── einsaetze.db                 # Einsatzprotokoll (WAL)
│   └── verspaetungen.db             # Verspätungs-Meldungen (WAL)
│
└── docs/
    ├── FUNKTIONEN.md
    └── REPRODUKTION.md
```

---

## 2. Datenbank-Setup

Die Datenbank wird beim ersten Start automatisch angelegt via `database/migrations.py`.

### `database/connection.py`
```python
import sqlite3
from pathlib import Path
from config import DB_PATH

def get_connection():
    conn = sqlite3.connect(DB_PATH, timeout=5, check_same_thread=False)
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous  = NORMAL")
    conn.execute("PRAGMA busy_timeout  = 5000")
    conn.row_factory = sqlite3.Row
    return conn
```

> **Hinweis:** Alle 5 Datenbanken nutzen WAL-Modus. Die anderen 4 DBs haben eigene `_connect()`-Helferfunktionen in ihren jeweiligen Modulen (`stellungnahmen_db.py`, `dienstliches.py`, `verspaetung_db.py`, `archiv_functions.py`).

### `database/migrations.py`
Erstellt beim App-Start alle benötigten Tabellen:
```python
from database.connection import get_connection

def run_migrations():
    conn = get_connection()
    c = conn.cursor()
    # Tabellen anlegen (CREATE TABLE IF NOT EXISTS)
    c.execute("""CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY, value TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS mitarbeiter (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, funktion TEXT, qualifikationen TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS fahrzeuge (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, status TEXT, zuletzt_geaendert TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS uebergabe (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        typ TEXT, beginn TEXT, ende TEXT,
        besonderheiten TEXT, erstellt_am TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS dienstplan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        datum TEXT, name TEXT, kuerzel TEXT,
        von TEXT, bis TEXT, ist_dispo INTEGER
    )""")
    conn.commit()
    conn.close()
```

---

## 3. Konfiguration (`config.py`)

```python
from pathlib import Path

BASE_DIR = Path(__file__).parent.resolve()
DB_PATH  = BASE_DIR / "database SQL" / "nesk3.db"
SHARED_DIR = BASE_DIR.parent.parent  # OneDrive-Stammordner
```

Alle weiteren DB-Pfade werden in den jeweiligen Modulen direkt über `BASE_DIR` aufgebaut:
```python
# Beispiel aus functions/stellungnahmen_db.py
DB_ORDNER = os.path.join(BASE_DIR, "database SQL")
DB_PFAD   = os.path.join(DB_ORDNER, "stellungnahmen.db")
```

---

## 4. Einstiegspunkt (`main.py`)

```python
import sys
from PySide6.QtWidgets import QApplication
from database.migrations import run_migrations
from gui.main_window import MainWindow

def main():
    run_migrations()
    app = QApplication(sys.argv)
    app.setApplicationName("Nesk3")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
```

---

## 5. Hauptfenster (`gui/main_window.py`)

```python
from PySide6.QtWidgets import QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QPushButton, QStackedWidget
from PySide6.QtGui import QFont

NAV_ITEMS = [
    ("🏠", "Dashboard",           0),
    ("👥", "Mitarbeiter",         1),
    ("☕️", "Dienstliches",       2),
    ("☀️", "Aufgaben Tag",        3),
    ("🌙", "Aufgaben Nacht",      4),
    ("📅", "Dienstplan",          5),
    ("📋", "Übergabe",            6),
    ("🚗", "Fahrzeuge",           7),
    ("🕐", "Code 19",             8),
    ("🖨️", "Ma. Ausdrucke",      9),
    ("🤒", "Krankmeldungen",     10),
    ("📞", "Telefonnummern",     11),
    ("♿",  "Call Transcription", 12),
    ("💾", "Backup",             13),
    ("⚙️",  "Einstellungen",     14),
    ("📣",  "Beschwerden",       15),
    ("✉️",  "Passagieranfragen", 16),
]

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setMinimumSize(1200, 700)
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Linke Navigation
        nav = QWidget()
        nav.setFixedWidth(160)
        nav_layout = QVBoxLayout(nav)
        # ... Buttons aus NAV_ITEMS ...

        # Rechter Stack
        self._stack = QStackedWidget()
        # ... Widgets hinzufügen ...

        layout.addWidget(nav)
        layout.addWidget(self._stack)
```

---

## 6. Section: E-Mobby Fahrer

### `functions/emobby_functions.py` aufbauen

```python
import json
from pathlib import Path
from config import BASE_DIR

_TXT_PATH     = Path(BASE_DIR) / "Daten" / "E-Mobby" / "mobby.txt"
_SETTINGS_KEY = "emobby_fahrer"

def _names_from_txt() -> list[str]:
    if not _TXT_PATH.exists():
        return []
    names = []
    for line in _TXT_PATH.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            names.append(line)
    return names

def get_emobby_fahrer() -> list[str]:
    """Merged TXT + DB, synct neue TXT-Einträge → DB."""
    from functions.settings_functions import get_setting, set_setting
    txt_names = _names_from_txt()
    db_raw = get_setting(_SETTINGS_KEY, "")
    try:
        db_names = json.loads(db_raw) if db_raw else []
    except Exception:
        db_names = []
    # Neue TXT-Namen in DB eintragen
    changed = False
    for n in txt_names:
        if n not in db_names:
            db_names.append(n)
            changed = True
    if changed:
        set_setting(_SETTINGS_KEY, json.dumps(db_names, ensure_ascii=False))
    return db_names

def is_emobby_fahrer(name: str) -> bool:
    """Case-insensitiver Substring-Match."""
    fahrer = get_emobby_fahrer()
    name_lower = name.lower()
    return any(f.lower() in name_lower or name_lower in f.lower() for f in fahrer)

def add_emobby_fahrer(name: str) -> bool:
    """Fügt Namen zur DB-Liste hinzu. Returns False wenn bereits vorhanden."""
    from functions.settings_functions import get_setting, set_setting
    db_raw = get_setting(_SETTINGS_KEY, "")
    try:
        db_names = json.loads(db_raw) if db_raw else []
    except Exception:
        db_names = []
    if name in db_names:
        return False
    db_names.append(name)
    set_setting(_SETTINGS_KEY, json.dumps(db_names, ensure_ascii=False))
    return True
```

### `Daten/E-Mobby/mobby.txt` – Format
```
# Kommentarzeilen werden ignoriert (beginnen mit #)
Bouladhne
Ormaba
Dobrani
... (ein Nachname pro Zeile)
```

---

## 7. Dienstplan-Parser – Aufbau

### `functions/dienstplan_parser.py` – Grundstruktur

```python
import openpyxl
from dataclasses import dataclass, field
from typing import Optional

@dataclass
class Mitarbeiter:
    name: str
    kuerzel: str
    von: str
    bis: str
    ist_dispo: bool = False
    ist_krank: bool = False
    krank_schicht_typ: str = ""      # tagdienst / nachtdienst / sonderdienst
    krank_abgeleiteter_dienst: str = ""

class DienstplanParser:
    def parse(self, xlsx_path: str) -> list[Mitarbeiter]:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        result = []
        aktueller_abschnitt = "betreuer"  # dispo / betreuer

        for row in ws.iter_rows(values_only=True):
            abschnitt = self._detect_abschnitt_header(list(row))
            if abschnitt:
                aktueller_abschnitt = abschnitt
                continue
            # Name-Spalte auslesen, Dienst, Zeiten...
            # ...
        return result

    def _detect_abschnitt_header(self, row_list) -> Optional[str]:
        for cell in row_list:
            if cell and "dispo" in str(cell).lower():
                return "dispo"
            if cell and any(w in str(cell).lower() for w in ["stamm", "betreuer"]):
                return "betreuer"
        return None
```

---

## 8. E-Mobby in Sonderaufgaben

### Combo-Logik in `_add_aufgabe_row()`

```python
def _add_aufgabe_row(self, grid, name, row, nur_bulmor=False):
    is_emobby = (name == "E-mobby Check")

    for schicht in ("tag", "nacht"):
        is_tag = (schicht == "tag")
        if nur_bulmor:
            mitarbeiter = self._tag_bulmor if is_tag else self._nacht_bulmor
        elif is_emobby:
            mitarbeiter = self._tag_emobby if is_tag else self._nacht_emobby
        else:
            mitarbeiter = self._tag_mitarbeiter if is_tag else self._nacht_mitarbeiter

        combo = QComboBox()
        if mitarbeiter:
            combo.addItems(["— bitte wählen —"] + mitarbeiter)
        elif is_emobby and self._dienstplan_geladen:
            combo.addItems(["⚠ Kein E-Mobby-Fahrer auf dieser Schicht – bitte prüfen!"])
            combo.setStyleSheet("QComboBox { color: #cc6600; font-weight: bold; }")
        else:
            combo.addItems(["— Dienstplan laden —"])
```

---

## 9. Übergabe: Automatisches Zeitbefüllen

```python
def _neues_protokoll(self, typ: str):
    # typ = "tagdienst" oder "nachtdienst"
    if typ == "tagdienst":
        self._f_beginn.setText("07:00")
        self._f_ende.setText("19:00")
    else:
        self._f_beginn.setText("19:00")
        self._f_ende.setText("07:00")
    # Felder freischalten, Formular leeren...
```

---

## 10. Code-19 Taschenuhr (`_PocketWatchWidget`)

```python
import math
from PySide6.QtWidgets import QWidget
from PySide6.QtCore import QTimer, Qt, QPointF
from PySide6.QtGui import QPainter, QColor, QRadialGradient, QPen, QFont
import datetime

class _PocketWatchWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(240, 300)
        self._swing_t   = 0.0
        self._swing_angle = 0.0
        self._blink_on  = True

        self._swing_timer = QTimer(self)
        self._swing_timer.timeout.connect(self._swing_step)
        self._swing_timer.start(25)   # ~40 FPS

        self._tick_timer = QTimer(self)
        self._tick_timer.timeout.connect(self._tick)
        self._tick_timer.start(1000)  # jede Sekunde

    def _swing_step(self):
        self._swing_t += 0.07
        self._swing_angle = 14.0 * math.sin(self._swing_t)
        self.update()

    def _tick(self):
        self._blink_on = not self._blink_on
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        # Pivot oben-mitte
        p.translate(120, 40)
        p.rotate(self._swing_angle)
        p.translate(-120, -40)
        # Uhr-Zentrum
        cx, cy = 120, 170
        # Gehäuse zeichnen (Radial-Gradient Gold)
        grad = QRadialGradient(cx, cy, 90)
        grad.setColorAt(0, QColor("#FFD700"))
        grad.setColorAt(1, QColor("#8B6914"))
        p.setBrush(grad)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(cx - 85, cy - 85, 170, 170)
        # Zifferblatt, Zeiger, röm. Ziffern...
        now = datetime.datetime.now()
        # ...
        p.end()
```

---

## 11. Einstellungen: E-Mobby GroupBox

```python
from PySide6.QtWidgets import QGroupBox, QListWidget, QLineEdit, QPushButton

grp_emobby = QGroupBox("🛵 E-Mobby Fahrer")
self._emobby_list  = QListWidget()
self._emobby_input = QLineEdit()
self._emobby_input.setPlaceholderText("Nachname eingeben …")
self._emobby_input.returnPressed.connect(self._add_emobby_entry)

btn_add = QPushButton("+ Hinzufügen")
btn_add.clicked.connect(self._add_emobby_entry)

btn_rem = QPushButton("🗑 Entfernen")
btn_rem.clicked.connect(self._remove_emobby_entry)

def _load_emobby_list(self):
    from functions.emobby_functions import get_emobby_fahrer
    self._emobby_list.clear()
    for n in get_emobby_fahrer():
        self._emobby_list.addItem(n)

def _add_emobby_entry(self):
    name = self._emobby_input.text().strip()
    if not name:
        return
    from functions.emobby_functions import add_emobby_fahrer
    if add_emobby_fahrer(name):
        self._emobby_input.clear()
        self._load_emobby_list()

def _remove_emobby_entry(self):
    selected = self._emobby_list.currentItem()
    if not selected:
        return
    import json
    from functions.settings_functions import get_setting, set_setting
    db_names = json.loads(get_setting("emobby_fahrer", "[]"))
    db_names.remove(selected.text())
    set_setting("emobby_fahrer", json.dumps(db_names, ensure_ascii=False))
    self._load_emobby_list()
```

---

## 12. App starten

```powershell
cd "...\Nesk\Nesk3"
python3.13 main.py
```

### Erste Inbetriebnahme Checkliste
- [ ] Python 3.13+ installiert
- [ ] `pip install PySide6 openpyxl python-docx pypdf pywin32`
- [ ] `Daten/E-Mobby/mobby.txt` mit Fahrernamen befüllt
- [ ] App starten → DB wird automatisch erstellt
- [ ] Einstellungen öffnen → Pfade zu Dienstplan-Excel, Sonderaufgaben, AOCC, Code-19 eintragen
- [ ] Dienstplan laden → ersten Dienstplan importieren

---

## 13. Passagieranfragen-Modul (`gui/passagieranfragen.py`)

Verarbeitet eingehende Passagieranfragen aus Outlook: liest den Posteingang, extrahiert Kontakt- und Flugdaten und erstellt passende Antwort-Entwürfe per win32com.

### Abhängigkeiten
```python
from pathlib import Path
from PySide6.QtWidgets import (
    QWidget, QDialog, QHBoxLayout, QVBoxLayout, QLabel,
    QPushButton, QTextEdit, QLineEdit, QComboBox, QCheckBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QSplitter,
    QDialogButtonBox, QMessageBox)
from PySide6.QtCore import Qt
from config import BASE_DIR

_LOGO_PATH = str(Path(BASE_DIR) / "Daten" / "Email" / "Logo.jpg")
```

### `OutlookInboxDialog` – Posteingang-Auswahl

```python
class OutlookInboxDialog(QDialog):
    """Lädt die letzten 75 Outlook-Inbox-Mails, zeigt sie als Tabelle."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Outlook-Posteingang")
        self.setMinimumSize(800, 500)
        self.selected_sender_email = ""
        self.selected_sender_name  = ""
        self._body = ""
        self._build_ui()
        self._load_mails()

    def _load_mails(self):
        try:
            import win32com.client
            try:
                ol = win32com.client.GetActiveObject("Outlook.Application")
            except Exception:
                ol = win32com.client.Dispatch("Outlook.Application")
            ns    = ol.GetNamespace("MAPI")
            inbox = ns.GetDefaultFolder(6)   # 6 = olFolderInbox
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)  # neueste zuerst
            for i in range(1, min(76, items.Count + 1)):
                msg = items[i]
                try:
                    datum   = str(msg.ReceivedTime)[:16]
                    sender  = msg.SenderName or ""
                    betreff = msg.Subject or ""
                    email   = msg.SenderEmailAddress or ""
                    # Exchange-interne Adressen (EX:/O=...) → SMTP-Fallback
                    if email.upper().startswith(("/O=", "EX:", "/CN=")):
                        try:
                            email = msg.ReplyRecipients(1).Address
                        except Exception:
                            email = ""
                    # Zeile in QTableWidget eintragen …
                except Exception:
                    continue
        except Exception as e:
            QMessageBox.warning(self, "Outlook", f"Fehler beim Laden: {e}")

    def get_body(self) -> str:
        return self._body
```

### `_parse_email_fields()` – 5-stufige Namensextraktion

```python
import re

def _parse_email_fields(text: str) -> dict:
    """Extrahiert Name, Anrede, E-Mail, Flugnummer, Datum aus E-Mail-Text."""
    result = {"name": "", "anrede": "", "email": "",
              "flugnummer": "", "datum": "", "rueckflug": ""}

    # Stufe 1: Vorname / Nachname als separate Felder (mehrzeilig)
    vn = re.search(r"Vorname[:\s]+(.+)", text, re.IGNORECASE)
    nn = re.search(r"Nachname[:\s]+(.+)", text, re.IGNORECASE)
    if vn and nn:
        result["name"] = f"{vn.group(1).strip()} {nn.group(1).strip()}"

    # Stufe 2: Anrede-Block  "Anrede: Herr Mustermann"
    m = re.search(r"Anrede[:\s]+(Herr|Frau)\s+(.+)", text, re.IGNORECASE)
    if m and not result["name"]:
        result["anrede"] = m.group(1).strip()
        result["name"]   = m.group(2).strip()

    # Stufe 3: "Name:" Label
    m = re.search(r"\bName[:\s]+(.+)", text, re.IGNORECASE)
    if m and not result["name"]:
        result["name"] = m.group(1).strip()

    # Stufe 4: "Herr / Frau" im Fließtext
    m = re.search(r"\b(Herr|Frau)\s+([A-ZÄÖÜ][a-zäöüß]+([\s-][A-ZÄÖÜ][a-zäöüß]+)*)", text)
    if m and not result["name"]:
        result["anrede"] = m.group(1)
        result["name"]   = m.group(2).strip()

    # Stufe 5: "Von:" Header als letzter Fallback
    m = re.search(r"Von:\s*(.+?)(?:\s*<|\s*$)", text, re.IGNORECASE)
    if m and not result["name"]:
        result["name"] = m.group(1).strip()

    # E-Mail-Adresse
    m = re.search(r"[\w.+-]+@[\w-]+\.[a-z]{2,}", text, re.IGNORECASE)
    if m:
        result["email"] = m.group(0)

    # Flugnummer (z. B. EW583, 4U123)
    m = re.search(r"\b([A-Z]{1,2}\d{2,4})\b", text)
    if m:
        result["flugnummer"] = m.group(1)

    # Datum (DD.MM.YYYY oder YYYY-MM-DD)
    m = re.search(r"\b(\d{1,2}\.\d{1,2}\.\d{4}|\d{4}-\d{2}-\d{2})\b", text)
    if m:
        result["datum"] = m.group(1)

    return result
```

### `PassagieranfragenWidget` – Aufbau

Linkes Panel: Posteingang-Button, E-Mail-Rohtext, 5 extrahierbare Felder (Name, Anrede, E-Mail, Flugnummer, Datum).  
Rechtes Panel: 4 Szenario-Buttons, Flugdaten-Checkbox, Antwort-Textarea, Outlook-Button.

```python
BUTTON_STYLE = (
    "background-color:#1e5799;color:white;font-weight:bold;"
    "border-radius:6px;padding:6px;"
)

class PassagieranfragenWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        # … QSplitter mit linkem und rechtem Panel …

        self._inbox_btn.clicked.connect(self._load_from_inbox)
        self._extract_btn.clicked.connect(self._extract)

        for label, slot in [
            ("✅ Antwort Vollständig",  self._antwort_komplett),
            ("❓ Fehlende Daten",       self._antwort_fehlende_daten),
            ("🅿️ Parkplatz-Info",      self._antwort_parkplatz),
            ("ℹ️ Allgemeine Info",      self._antwort_info_service),
        ]:
            btn = QPushButton(label)
            btn.setStyleSheet(BUTTON_STYLE)
            btn.clicked.connect(slot)

    def _load_from_inbox(self):
        dlg = OutlookInboxDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._email_text.setPlainText(dlg.get_body())
            self._extract()
            # Sender-E-Mail aus COM überschreibt ggf. extrahierte Adresse
            if dlg.selected_sender_email:
                self._f_email.setText(dlg.selected_sender_email)
            if not self._f_name.text() and dlg.selected_sender_name:
                self._f_name.setText(dlg.selected_sender_name)

    def _extract(self):
        fields = _parse_email_fields(self._email_text.toPlainText())
        self._f_name.setText(fields["name"])
        self._f_anrede.setCurrentText(fields["anrede"])
        self._f_email.setText(fields["email"])
        self._f_flug.setText(fields["flugnummer"])
        self._f_datum.setText(fields["datum"])

    def _set_antwort(self, vorlage: str):
        name   = self._f_name.text().strip()
        anrede = self._f_anrede.currentText()
        flug   = self._f_flug.text().strip()
        datum  = self._f_datum.text().strip()

        if anrede and name:
            suffix = "r" if anrede == "Herr" else ""
            greet  = f"Sehr geehrte{suffix} {anrede} {name},"
        elif name:
            greet = f"Sehr geehrte/r {name},"
        else:
            greet = "Sehr geehrte Damen und Herren,"

        bezug = f"\nBezug: Flug {flug}, {datum}\n" if (flug or datum) else ""
        text  = f"{greet}{bezug}\n\n{vorlage}"

        if self._flugdaten_chk.isChecked():
            text += ("\n\n[Bitte teilen Sie uns Ihre genauen Flugdaten mit,"
                     " damit wir Ihnen weiterhelfen können.]")
        self._antwort_text.setPlainText(text)

    def _open_outlook(self):
        from functions.mail_functions import create_outlook_draft
        create_outlook_draft(
            to=self._f_email.text().strip(),
            subject=f"Ihre Anfrage – Flug {self._f_flug.text().strip()}",
            body_text=self._antwort_text.toPlainText(),
            logo_path=_LOGO_PATH,
        )
```

### In `main_window.py` registrieren

```python
from gui.passagieranfragen import PassagieranfragenWidget

# NAV_ITEMS Index 16:
("✉️", "Passagieranfragen", 16),

# In _build_ui() – Stack-Widget:
self._stack.addWidget(PassagieranfragenWidget())
```

---

## 14. EXE erstellen (PyInstaller)

```powershell
cd "...\Nesk\Nesk3"
python3.13 -m PyInstaller Nesk3.spec
```

Die `Nesk3.spec` ist bereits konfiguriert mit:
- `icon = 'Daten/Logo/nesk3.ico'`
- `datas`: alle `Daten/`-Ordner
- `upx = False` (verhindert Antivirus-Fehlalarme)
- Ausgabe: `dist/Nesk3.exe` (~69 MB)

### Backup erstellen (vor EXE-Build empfohlen)
```python
from backup.backup_manager import create_zip_backup
zip_pfad = create_zip_backup()
print(f"Backup: {zip_pfad}")
```
Ausgeschlossen: `Backup Data/`, `build_tmp/`, `Exe/`, `__pycache__/`, `dist/` → Größe ~8 MB
