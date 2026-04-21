# Nesk3 – Vollständige Funktionsübersicht

**Stand:** 20.04.2026 – v3.10.0  
**App:** Nesk3 – DRK Erste-Hilfe-Station Flughafen Köln/Bonn  
**Zweck:** Dienstplan-Verwaltung, Stärkemeldung, Mitarbeiterdokumente, Einsatzprotokoll, Verspätungs-Meldungen, Übergabe, Code-19, Telefonnummern-Verzeichnis, Anrufprotokoll, Beschwerdemanagement, Passagieranfragen-Bearbeitung, **Schulungs-Verwaltung**, **Sanitätsmaterial-Verbrauch**, **Vorkommnisberichte**

---

## Inhaltsverzeichnis

1. [Einstiegspunkt & Hauptfenster](#1-einstiegspunkt--hauptfenster)
2. [Dashboard](#2-dashboard)
3. [Dienstplan](#3-dienstplan)
4. [Aufgaben Nacht](#4-aufgaben-nacht)
5. [Aufgaben Tag (Tabs)](#5-aufgaben-tag-tabs)
6. [Sonderaufgaben](#6-sonderaufgaben)
7. [Übergabe](#7-übergabe)
8. [Fahrzeuge](#8-fahrzeuge)
9. [Code 19](#9-code-19)
10. [Mitarbeiter](#10-mitarbeiter)
11. [Dienstliches (Einsatzprotokoll)](#11-dienstliches-einsatzprotokoll)
12. [Mitarbeiterdokumente / Stellungnahmen / Verspätung](#12-mitarbeiterdokumente--stellungnahmen--verspätung)
13. [Einstellungen](#13-einstellungen)
14. [Beschwerden](#14-beschwerden)
15. [Passagieranfragen](#15-passagieranfragen)
16. [Checklisten / Backup / Ma. Ausdrucke / Telefonnummern / Anrufprotokoll](#16-checklisten--backup--ma-ausdrucke--telefonnummern--anrufprotokoll)
17. [Datenbanken (SQLite)](#17-datenbanken-sqlite)
18. [Functions-Module](#18-functions-module)
19. [HilfeDialog](#19-hilfedialog)
20. [Konfiguration (config.py)](#20-konfiguration-configpy)
21. [Vorkommnisse](#21-vorkommnisse)

---

## 1. Einstiegspunkt & Hauptfenster

### `main.py`
- Startet `QApplication` und zeigt `MainWindow`
- Setzt App-Name und Icon

### `gui/main_window.py` – `MainWindow(QMainWindow)`
- Linke Navigationsleiste mit Icon + Label-Buttons
- NAV_ITEMS (Index, Icon, Label, Widget-Index):
  | Icon | Label | Seite |
  |------|-------|-------|
  | 🏠 | Dashboard | 0 |
  | � | Mitarbeiter | 1 |
  | ☕️ | Dienstliches | 2 |
  | ☀️ | Aufgaben Tag | 3 |
  | 🌙 | Aufgaben Nacht | 4 |
  | 📅 | Dienstplan | 5 |
  | 📋 | Übergabe | 6 |
  | 🚗 | Fahrzeuge | 7 |
  | 🕐 | Code 19 | 8 |
  | 🖨️ | Ma. Ausdrucke | 9 |
  | 🤒 | Krankmeldungen | 10 |
  | � | Telefonnummern | 11 |
  | ♿ | Call Transcription | 12 |
  | 💾 | Backup | 13 |
  | ⚙️ | Einstellungen | 14 |
  | 📣 | Beschwerden | 15 |
  | ✉️ | Passagieranfragen | 16 |
- `QStackedWidget` als Hauptbereich (17 Seiten)
- Automatisches Laden beim Start: Dienstplan-Status aus DB

---

## 2. Dashboard

### `gui/dashboard.py` – `DashboardWidget(QWidget)`
- Statistik-Karten (Mitarbeiter, Fahrzeuge, Protokolle, Aufgaben)
- DB-Status-Anzeige
- `_SkyWidget(QWidget)`: QPainter-Animation mit Himmels-Gradient, Wolken, Landebahn, fliegendem ✈-Emoji (~33 FPS via QTimer 30ms)
- `FlugzeugWidget(QFrame)`: Klickbare Karte mit hochzählendem Verspätungs-Ticker (1/s), QMessageBox bei Klick
- **Notizen-Kalender**: Kalender-Widget mit farbigen Punkten – blau = Termin, grün = Notiz; Doppelklick auf Tag öffnet Notiz/Termin-Dialog
- **`_neue_notiz_dialog()`**: Dialog zum Anlegen neuer Notizen mit Datum, Titel, Text und optionalem Fälligkeitsdatum
- **`functions/notizen_db.py`**: SQLite CRUD für `notizen.db` (WAL); Funktionen: `speichern`, `als_gelesen`, `als_erledigt`, `loeschen`, `lade_aktive`, `lade_alle`, `lade_fuer_datum`

---

## 3. Dienstplan

### `gui/dienstplan.py` – `DienstplanWidget(QWidget)`
Kernfunktionen:
- **Excel laden**: Öffnet `.xlsx`-Datei via Dateiauswahl oder gespeicherten Pfad aus Einstellungen
- **Tabelle anzeigen**: HTML-Tabelle mit farbcodierten Diensten
- **Statuszeile**: Tagdienst / Nachtdienst / Krank-Aufschlüsselung (Betreuer, Dispo, Tag/Nacht/Sonder-Krank)
- **Export**: Word-Stärkemeldung generieren (via `staerkemeldung_export.py`) – Dateiname mit korrektem Umlaut `Stärkemeldung`
- **Dienst-Typen** (`_TAG_DIENSTE`, `_NACHT_DIENSTE`, `_SONDER_DIENSTE`): Bestimmen Farbe und Kategorisierung
- **Info-Banner**: Oben erklärt „Bis zu 4 Dienstpläne gleichzeitig öffnen“
- **Export-Button-Text**: Inaktiv: `'Hier klicken um Datei als Wordexport auswählen'` / Aktiv: `'✓  Für Wordexport gewählt'`

### `functions/dienstplan_parser.py` – `DienstplanParser`
- `parse(xlsx_path)`: Liest Excel-Datei, extrahiert alle Mitarbeiter mit Dienst, Zeiten, Funktion
- `_detect_abschnitt_header(row)`: Erkennt Dispo/Betreuer-Abschnitte
- `_ermittle_krank_typ(start_zeit, end_zeit, vollname)`: Tag/Nacht/Sonder-Krank-Klassifizierung
- `_runde_auf_volle_stunde(zeit_str)`: Minutenabweichungen korrigieren (für Dispo)
- `_betr_zu_dispo_kuerzel(kuerzel)`: `N→DN`, `T→DT`, `T10→DT`, `N10→DN`
- Rückgabe je Mitarbeiter: `name`, `kuerzel`, `von`, `bis`, `ist_dispo`, `ist_krank`, `krank_schicht_typ`, etc.

### `functions/dienstplan_functions.py`
- `save_dienstplan(data)`: Speichert geparsten Dienstplan in SQLite
- `get_dienstplan(datum)`: Lädt Dienstplan aus DB
- `get_mitarbeiter_schicht(name, datum)`: Gibt Schicht eines Mitarbeiters zurück

---

## 4. Aufgaben Nacht

### `gui/aufgaben.py` – `AufgabenWidget(QWidget)`
Tabs:
| Nr. | Tab | Klasse |
|-----|-----|--------|
| 1 | 📋 Allgemein | `_AllgemeinTab` (in aufgaben.py) |
| 2 | 📋 Checklisten | `_ChecklistenTab` |
| 3 | 📋 Freie Mail | integriert |
| 4 | 📋 Code 19 Mail | `_Code19MailTab` (aus aufgaben_tag.py) |

---

## 5. Aufgaben Tag (Tabs)

### `gui/aufgaben_tag.py`
Tabs innerhalb der Tagdienst-Ansicht:
- **_Code19MailTab**: Mail-Assistent für Code-19-Benachrichtigungen
  - Empfänger, Betreff, Nachrichtentext vorausgefüllt
  - Outlook-Integration via COM (VBS-Script-Logik)
  - „Signatur einfügen"-Button
  - **Blauer Info-Kasten „Zeitraum“**: Erklärt welche Excel-Zeilen ausgelesen werden (Standard: letzte 7 Tage)
- **_FreieMailTab**: Frei konfigurierbarer Mail-Tab mit Anhang-Support
  - **Blauer Info-Kasten** nach Template-Buttons erklärt Checklisten- / Checks-Template
  - **Gelber Info-Kasten** nach Umbenennen-Checkbox erklärt `JJJJ_MM_TT`-Umbenennung
- **_ChecklistenTab**: Checklisten-Ansicht für Tagdienst (Symbol: `📋 Checklisten`)
- Weitere Tabs für tagesspezifische Aufgaben

---

## 6. Sonderaufgaben

### `gui/sonderaufgaben.py` – `SonderaufgabenWidget(QWidget)`
Kernfunktionen:
- **Dienstplan laden**: Liest Dienstplan-Excel, extrahiert:
  - `_tag_mitarbeiter` / `_nacht_mitarbeiter`: Alle Mitarbeiter je Schicht
  - `_tag_bulmor` / `_nacht_bulmor`: Bulmor-Fahrer (gefiltert nach Qualifikation)
  - `_tag_emobby` / `_nacht_emobby`: E-Mobby-Fahrer (abgeglichen mit DB-Liste via `is_emobby_fahrer()`)
  - Flag `_dienstplan_geladen` wird auf `True` gesetzt
- **`_build_form()`**: Baut Aufgaben-Formular mit Dropdown-Combos je Aufgabe und Schicht
- **`_add_aufgabe_row(grid, name, row, nur_bulmor)`**: Erstellt eine Zeile (Combo Tag + Textfeld + Combo Nacht + Textfeld)
  - Wenn `nur_bulmor=True` → nur Bulmor-Fahrer in Combo
  - Wenn `is_emobby=True` (Name == "E-mobby Check"):
    - Fahrer gefunden → auswählbar
    - Keine Fahrer + Dienstplan geladen → `⚠ Kein E-Mobby-Fahrer – bitte prüfen!` (orange)
    - Dienstplan nicht geladen → `— Dienstplan laden —`
- **Speichern**: Excel-Export der ausgefüllten Sonderaufgaben

### `functions/emobby_functions.py`
- `get_emobby_fahrer() → list[str]`: Liest `Daten/E-Mobby/mobby.txt`, synct neue Namen in DB
- `is_emobby_fahrer(name: str) → bool`: Case-insensitiver Substring-Match gegen DB-Liste
- `add_emobby_fahrer(name: str) → bool`: Fügt Namen zur DB hinzu (Duplikat-Check, returns False wenn bereits vorhanden)
- **DB-Key**: `emobby_fahrer` in `settings`-Tabelle als JSON-Array
- **TXT-Pfad**: `Daten/E-Mobby/mobby.txt` (33 Fahrer initial)

---

## 7. Übergabe

### `gui/uebergabe.py` – `UebergabeWidget(QWidget)`
Kernfunktionen:
- **Neues Protokoll**: Button Tagdienst / Nachtdienst
  - Tagdienst: Beginn 07:00, Ende 19:00 (automatisch)
  - Nachtdienst: Beginn 19:00, Ende 07:00 (automatisch)
- **Felder**: Besonderheiten, Fahrzeugstatus, sonstige Hinweise
- **Speichern**: Speichert Protokoll in SQLite (`uebergabe`-Tabelle) – Protokoll bleibt als `offen` bearbeitbar
- **Abschließen**: Setzt Status auf `abgeschlossen` – danach keine Bearbeitung mehr möglich; Abzeichner-Name erforderlich
- **E-Mail**: Erstellt Outlook-Entwurf mit Protokollinhalt
- **Verlauf laden**: Vorhandene Protokolle anzeigen und bearbeiten
- **Blaue Info-Box**: Erklärt Unterschied Speichern / Abschließen / E-Mail
- Tooltips auf allen 4 Buttons (Speichern, Abschließen, E-Mail, Löschen)
- (Entfernt: „Personal im Dienst" – kein Textfeld mehr)

### `functions/uebergabe_functions.py`
- `save_uebergabe(data)`: Speichert Übergabe-Protokoll
- `get_uebergaben(limit)`: Lädt letzte N Protokolle
- `get_uebergabe_by_id(id)`: Lädt einzelnes Protokoll

---

## 8. Fahrzeuge

### `gui/fahrzeuge.py` – `FahrzeugeWidget(QWidget)`
- Fahrzeugliste mit Status (verfügbar / in Wartung / außer Dienst)
- Hinzufügen / Bearbeiten / Löschen von Fahrzeugen
- Statusänderung mit Zeitstempel

### `functions/fahrzeug_functions.py`
- `get_fahrzeuge()`: Alle Fahrzeuge aus DB
- `save_fahrzeug(data)`: Fahrzeug speichern/aktualisieren
- `delete_fahrzeug(id)`: Fahrzeug löschen
- `update_status(id, status)`: Status aktualisieren

---

## 9. Code 19

### `gui/code19.py` – `Code19Widget(QWidget)`
- Titelleiste: `🕐 Code 19`
- **`_PocketWatchWidget(QWidget)`** (240×300 px): Alice-im-Wunderland Taschenuhr-Animation
  - `_swing_timer` (25 ms) → Pendel-Swing ±14° via `math.sin()`
  - `_tick_timer` (1000 ms) → Sekundenzeiger (ruckartig), Blink-Punkt toggle
  - `paintEvent()`: Radial-Gradient Golden (#FFD700→#8B6914), Zifferblatt, römische Ziffern, Echtzeit-Zeiger
  - Zitat: „Ich bin spät! Ich bin spät!"
- Code-19-Protokoll: Erstellung, Verwaltung, Excel-Export (`code19_datei` aus Einstellungen)

---

## 10. Mitarbeiter

### `gui/mitarbeiter.py` – `MitarbeiterWidget(QWidget)`
- Mitarbeiterliste mit Suche und Filter
- CRUD: Hinzufügen, Bearbeiten, Löschen
- **Spalte „Export“** (✅/🚫): Zeigt ob Mitarbeiter in der Stärkemeldungs-Word erscheint
- **🚫 Ausschließen-Button**: Schließt Mitarbeiter vom Word-Export aus ohne ihn zu löschen
- **Gelbe Info-Box**: Erklärt Unterschied Ausschließen (kein Export) vs. Löschen
- Ausgeschlossene Zeilen werden rot hinterlegt

### `functions/mitarbeiter_functions.py`
- `get_mitarbeiter()`: Alle Mitarbeiter aus DB
- `save_mitarbeiter(data)`: Mitarbeiter speichern
- `delete_mitarbeiter(id)`: Mitarbeiter löschen
- `search_mitarbeiter(query)`: Volltextsuche

---

## 11. Dienstliches (Einsatzprotokoll)

### `gui/dienstliches.py` – `DienstlichesWidget(QWidget)`
2 Tabs:
- **📋 Einsatzprotokoll**: Neues Einsatz-/Veranstaltungsprotokoll erfassen
  - Felder: Datum, Uhrzeit, Einsatzart, Anlass, Einsatzleiter, Kräfte, Beschreibung
  - Speichern in `einsaetze.db` (WAL)
  - Excel-Export pro Einsatz in `Daten/Einsatz/Protokolle/`
- **📊 Übersicht**: Alle gespeicherten Einsätze tabellarisch, filterbar
  - Rechtsklick-Menü: Bearbeiten, Löschen, Excel öffnen

**Datenbank:** `database SQL/einsaetze.db` (WAL, busy_timeout=5000)  
**Zugriff:** direkt in `gui/dienstliches.py` via `_ensured_db()` / `_db()` Context-Manager

---

## 12. Mitarbeiterdokumente / Stellungnahmen / Verspätung

### `gui/mitarbeiter_dokumente.py` – `MitarbeiterDokumenteWidget(QWidget)`
3 Tabs:

**Tab 1 – Mitarbeiterdokumente / Ausdrucke:**
- Ordner-Browser für `Daten/Mitarbeiterdokumente/` (kategorisiert)
- Direktes Öffnen von `.docx`, `.pdf`, `.xlsx` per Doppelklick
- Neue Dokumente aus Word-Vorlage erstellen (DRK-Briefkopf)

**Tab 2 – Stellungnahmen (Passagierbeschwerden):**
- Liste aller Stellungnahmen aus `stellungnahmen.db`
- Spalten: Datum, Mitarbeiter, Art, Flugnummer, Status
- Neue Stellungnahme via Formular: Mitarbeiter, Art, Flugnummer, Freitext, Anhänge
- Status-Verlauf: offen → in Bearbeitung → abgeschlossen
- Word-Vorlage automatisch ausfüllen + speichern
- Rechtsklick: Bearbeiten, Status ändern, Dokument öffnen, Löschen
- HTML-Export (`WebNesk/stellungnahmen_lokal.html`)

**Tab 3 – Verspätungs-Meldungen:**
- Liste aller Meldungen aus `verspaetungen.db`
- Spalten: Datum, Mitarbeiter, Minuten, Art, Status
- Neue Meldung: Name, Datum, Verspätungsminuten, Beschreibung
- Word-Vorlage `FO_CGN_27` automatisch ausfüllen

**Zugehörige Functions-Module:**

| Datei | Funktion |
|---|---|
| `functions/stellungnahmen_db.py` | CRUD für Stellungnahmen-DB (WAL) |
| `functions/stellungnahmen_html_export.py` | HTML-Ansicht generieren |
| `functions/verspaetung_db.py` | CRUD für Verspätungs-DB (WAL, `_connect()`-Helfer) |
| `functions/archiv_functions.py` | Archivierung abgeschlossener Protokolle (WAL) |
| `functions/mitarbeiter_dokumente_functions.py` | Kategorien, Dateipfade, Vorlagen-Handling |

---

## 13. Einstellungen

### `gui/einstellungen.py` – `EinstellungenWidget(QWidget)`
Gruppen:
| Gruppe | Inhalt |
|--------|--------|
| 📂 Dienstplan-Ordner | Pfad zur Excel-Dienstplan-Datei |
| 📂 Sonderaufgaben-Ordner | Speicherpfad für Sonderaufgaben-Excel |
| 📊 AOCC Lagebericht | Pfad zur AOCC-Excel-Datei |
| 🕐 Code-19-Datei | Pfad zur Code-19-Excel |
| 🛵 E-Mobby Fahrer | Liste + Hinzufügen/Entfernen |

E-Mobby-Verwaltung:
- `QListWidget` zeigt alle in DB gespeicherten Fahrer
- Textfeld + „+ Hinzufügen" (Enter-Taste und Button)
- „🗑 Entfernen" mit Bestätigung
- Zähler-Label „X Fahrer in der Liste"
- Änderungen sofort in DB gespeichert (kein separater Speichern-Button)- Beschreibungstext präzisiert: erklärt Verbindung zur Übergabe-Ansicht und dass nur Nachnamen einzutragen sind
### `functions/settings_functions.py`
- `get_setting(key, default='')`: Liest Wert aus `settings`-Tabelle
- `set_setting(key, value)`: Schreibt Wert in `settings`-Tabelle

---

## 14. Beschwerden

### `gui/beschwerden.py` – `BeschwerdenWidget(QWidget)`
- Erfassung von Beschwerden aus Freitext, Word-Dateien (`.docx`) oder PDF-Dateien
- **Import-Dialog (`BeschwerdeImportDialog`)**: lädt Freitext oder Datei, zeigt Vorschau; unterstützt Drag & Drop
- **Heuristische Extraktion**: Namens-Erkennung, E-Mail-Erkennung, Telefonnummer-Erkennung aus dem Beschwerde-Text
- **Anonymisierung**: automatisches Entfernen von Personendaten vor der KI-Analyse
- **Gemini-KI-Analyse**: Anfrage an Google Generative AI API (Hintergrund-Thread via `threading`)
- **Outlook-Antwort**: erstellt Mail-Entwurf via `mail_functions.create_outlook_draft()`
- **Bearbeitungs-Dialog (`BeschwerdeBearbeitenDialog`)**: Bearbeiten/Nachpflegen bestehender Einträge
- **Volltext-Viewer (`VolltextDialog`)**: scrollbare Anzeige langer Beschwerde-Texte
- **Statussystem**: offen → in Bearbeitung → abgeschlossen
- Speicherung in `beschwerden.db` via `functions/beschwerden_db.py`

### `functions/beschwerden_db.py`
- `init_db()`: Erstellt `beschwerden.db` mit WAL-Modus
- `beschwerde_speichern(daten)`: Neue Beschwerde anlegen
- `beschwerde_aktualisieren(id, daten)`: Eintrag aktualisieren
- `alle_beschwerden_laden()`: Alle Einträge aus DB
- `beschwerde_laden(id)`: Einzelner Eintrag
- `beschwerde_loeschen(id)`: Eintrag löschen

---

## 15. Passagieranfragen

### `gui/passagieranfragen.py` – `PassagieranfragenWidget(QWidget)`
- Verarbeitung eingehender Passagieranfragen aus dem Outlook-Posteingang
- **`OutlookInboxDialog(QDialog)`**: zeigt die letzten 75 Inbox-E-Mails via `win32com.client`; Doppelklick oder OK lädt E-Mail-Text und Absenderdaten
- **Exchange-Adressen**: prüft `SenderEmailAddress` auf `EX:`/`/O=` → SMTP-Fallback via `ReplyRecipients(1).Address`
- **`_parse_email_fields(text)`**: 5-stufige Namens-Extraktion (Vorname/Nachname-Felder → Anrede-Block → `Name:`-Label → Herr/Frau Fließtext → `Von:`-Header)
- **4 Antwort-Szenarien**: Antwort Vollständig, Fehlende Daten, Parkplatz-Info, Allgemeine Info
- **Personalisierte Begrüßung**: `Sehr geehrter Herr X,` / `Sehr geehrte Frau X,`
- **Bezug-Zeile**: automatisch aus extrahierter Flugnummer + Datum zusammengesetzt
- **Flugdaten-Checkbox**: fügt Anfrage nach vollständigen Flugdaten an Antworttext an
- **Outlook-Entwurf**: `create_outlook_draft()` mit eingebettetem DRK-Logo (CID `nesk_logo`)
- Kein eigenes DB-Modul – reine Outlook-Integration via `win32com`

---

## 16. Checklisten / Backup / Ma. Ausdrucke / Telefonnummern / Anrufprotokoll

### `gui/checklisten.py` – `ChecklistenWidget(QWidget)`
- Vordefinierte und benutzerdefinierte Checklisten
- Abhaken mit Zeitstempel
- Tages-Reset

### Nav-Seite „Ma. Ausdrucke" (Index 9)
- Öffnet direkt den Ordner `Daten/Vordrucke/` im Explorer

### Nav-Seite „Krankmeldungen" (Index 10)
- Öffnet Ordner `Daten/Krankmeldungen/` oder zeigt Dokument-Browser

### Nav-Seite „Telefonnummern" (Index 11)
→ vollständiges Widget: `gui/telefonnummern.py` – importiert Nummern-Verzeichnisse aus Excel (FKB Gate-/Check-In-Nummern, DRK-Kontakte), mit Suche, Kategorien und Bearbeitungs-Dialog; DB-Zugriff via `functions/telefonnummern_db.py`

### Nav-Seite „Call Transcription" (Index 12)
→ vollständiges Widget: `gui/call_transcription.py` (`CallTranscriptionWidget`) – Anrufprotokoll mit Textbausteinen, Flugnummer-Auswahl mit Ziel/Herkunft-Befüllung, PRM-Format-Export in Zwischenablage, Verlaufsliste; DB via `functions/call_transcription_db.py`

### Nav-Seite „Backup" (Index 13)
- ZIP-Backup erstellen (via `backup_manager.py`): `Backup Data/Nesk3_backup_YYYYMMDD_HHMMSS.zip`
- Backup-Liste anzeigen, Restore anstoßen
- **Ausgeschlossen**: `Backup Data/`, `build_tmp/`, `Exe/`, `__pycache__/` → Größe ~8 MB

---

## 17. Datenbanken (SQLite)

Alle 6 SQLite-Datenbanken liegen seit **05.03.2026** zentral in `database SQL/`.

| Datei | Beschreibung | WAL | Zugriff über |
|---|---|---|---|
| `nesk3.db` | Hauptdatenbank (Dienstplan, Mitarbeiter, Fahrzeuge, Übergabe, …) | ✅ | `database/connection.py` |
| `archiv.db` | Archivierte Protokolle | ✅ | `functions/archiv_functions.py` |
| `stellungnahmen.db` | Passagierbeschwerde-Stellungnahmen | ✅ | `functions/stellungnahmen_db.py` |
| `einsaetze.db` | Einsatzprotokoll FKB (Dienstliches) | ✅ | `gui/dienstliches.py` |
| `verspaetungen.db` | Verspätungs-Meldungen | ✅ | `functions/verspaetung_db.py` |
| `beschwerden.db` | Beschwerden-Verwaltung [NEU v3.5] | ✅ | `functions/beschwerden_db.py` |
| `vorkommnisse.db` | Vorkommnisberichte [NEU v3.10] | ✅ | `functions/vorkommnisse_db.py` |
| `notizen.db` | Dashboard-Notizen und Termine [NEU v3.10] | ✅ | `functions/notizen_db.py` |

**WAL-Konfiguration** (alle DBs):
```python
PRAGMA journal_mode = WAL
PRAGMA synchronous  = NORMAL
PRAGMA busy_timeout  = 5000   # 5 Sekunden
sqlite3.connect(pfad, timeout=5)
```

**Tabellen in `nesk3.db`:**
| Tabelle | Inhalt |
|---------|--------|
| `mitarbeiter` | Mitarbeiterstammdaten |
| `dienstplan` | Dienstplan-Einträge (Name, Datum, Dienst, Zeiten) |
| `fahrzeuge` | Fahrzeugdaten mit Status |
| `uebergabe` | Übergabe-Protokolle |
| `settings` | Key-Value-Einstellungen |
| `code19` | Code-19-Protokolleinträge |
| `sonderaufgaben` | Gespeicherte Sonderaufgaben |

### `database/connection.py`
- `get_connection()`: SQLite-Verbindung mit WAL + `check_same_thread=False`

### `database/migrations.py`
- `run_migrations()`: Erstellt fehlende Tabellen und Spalten (wird bei App-Start ausgeführt)

### `database/models.py`
- Dataclass-ähnliche Modelle für DB-Entitäten

---

## 18. Functions-Module

| Datei | Hauptfunktionen |
|-------|----------------|
| `archiv_functions.py` | Archivierung abgeschlossener Protokolle (WAL) |
| `dienstplan_html_export.py` | Statische HTML-Ansicht des Dienstplans generieren |
| `dienstplan_parser.py` | Excel parsen, Krank-Typen, Abschnitt-Erkennung |
| `dienstplan_functions.py` | DB CRUD für Dienstplan |
| `emobby_functions.py` | E-Mobby-Fahrerliste (TXT↔DB-Sync, Matching) |
| `fahrzeug_functions.py` | DB CRUD für Fahrzeuge |
| `mail_functions.py` | Outlook-COM-Integration, Mail verschicken |
| `mitarbeiter_dokumente_functions.py` | Kategorien, Dateipfade, Vorlagen-Handling |
| `mitarbeiter_functions.py` | DB CRUD für Mitarbeiter |
| `settings_functions.py` | Key-Value-Einstellungen aus DB |
| `staerkemeldung_export.py` | Word-Dokument-Export (.docx) |
| `stellungnahmen_db.py` | CRUD für Stellungnahmen-DB (WAL) |
| `stellungnahmen_html_export.py` | HTML-Ansicht Stellungnahmen generieren |
| `uebergabe_functions.py` | DB CRUD für Übergabe-Protokolle |
| `verspaetung_db.py` | CRUD für Verspätungs-DB (WAL, `_connect()`-Helfer) |
| `beschwerden_db.py` | CRUD für Beschwerden-DB (WAL) [NEU v3.5] |
| `call_transcription_db.py` | CRUD für Anrufprotokoll-DB [NEU v3.5] |
| `telefonnummern_db.py` | Import und CRUD für Telefonnummern-Verzeichnis [NEU v3.5] |
| `dienstanweisungen_db.py` | DB für Dienstanweisungen [NEU v3.5] |
| `psa_db.py` | DB für PSA-Verwaltung [NEU v3.5] |
| `verspaetung_functions.py` | Hilfsfunktionen für Verspätungs-Meldungen [NEU v3.5] |
| `vorkommnisse_db.py` | CRUD für Vorkommnisberichte-DB (WAL) [NEU v3.10] |
| `notizen_db.py` | CRUD für Notizen/Termine-DB (WAL) [NEU v3.10] |

---

## 19. HilfeDialog

### `gui/hilfe_dialog.py` – `HilfeDialog(QDialog)`
Animierter Hilfe-Dialog mit 5 Tabs:
| Tab | Inhalt |
|-----|--------|
| 🏠 Übersicht | Willkommens-Banner, Schnellübersicht aller Module |
| 📦 Module | Jedes Modul mit 6–11 detaillierten Bullet-Points |
| 🔄 Workflow | 8 Schritte mit ausführlicher Beschreibung + Sondersituationen-Abschnitt |
| 💡 Tipps & FAQ | 14 Tipps + 5 FAQ-Einträge + Versionsinfo |
| 📖 Anleitungen | 5 vollständige Schritt-für-Schritt-Anleitungen |

Animationen:
- Fade+Slide beim Tab-Wechsel (`QPropertyAnimation`)
- Puls-Icon auf dem Hilfe-Button
- Laufbanner mit Datum + Version
- Workflow-Progress-Bar

---

## 20. Konfiguration (`config.py`)

```python
BASE_DIR    # Absoluter Pfad zu Nesk3/
DB_PATH     # Pfad zur SQLite-Datei
SHARED_DIR  # Pfad zum gemeinsamen OneDrive-Ordner
```

Farben für Dienstplan-Tabelle (HTML-Farben für verschiedene Dienst-Typen).

---

## 21. Vorkommnisse

### `gui/vorkommnisse.py` – `VorkommnisseWidget(QWidget)`
- Vollständiges Vorfallbericht-Formular für flugrelevante und allgemeine Vorkommnisse
- **Label**: „Flugnummer / Vorkommnis:" – universell für Flüge und sonstige Ereignisse
- **Felder**: Datum, Uhrzeit, Ort, Beteiligte Mitarbeiter, Kategorie, Beschreibung, Maßnahmen
- **Offblock-Felder (optional)**:
  - Plan-Offblock: `QTimeEdit` + `QCheckBox("angeben")` in Container; standardmäßig inaktiv
  - Ist-Offblock: ebenso; Checkbox aktiviert/deaktiviert das Zeitfeld
  - `_sammle_daten()`: gibt `""` zurück wenn Checkbox inaktiv; Verspätungsberechnung nur wenn beide Felder aktiv und Plan < Ist
- **1. Betroffene Personen** (`_EditTable` mit 4 Spalten):
  - Spalten: Person (Name), Typ, Kategorie, Anmerkung
  - Typ-Dropdown: `Passagier`, `PRM Passagier`, `Patient`, `Mitarbeiter`, `Sonstige`
  - Kategorie (Spalte 2) nur aktiv wenn Typ = „PRM Passagier"
  - PRM-Kategorien: `WCHS`, `WCHR`, `WCHC`, `BLND`, `DEAF`, `DPNA`, `UMNR`, `STCR`, `MEDA`, `Sonstiges`
- **`_EditTable` (generische editierbare Tabelle)**:
  - Parameter `conditional_columns: dict[int, tuple[int, list[str]]]`
  - Methode `_update_conditional(row, target_col, text, enabled_vals)`: deaktiviert + leert Spalte wenn Bedingung nicht erfüllt
  - Signal-Wiring in `_add_row()` und `set_data()`
- **Header-Buttons**: Neu, Speichern, Word-Export (`📄`), Ordner öffnen (`📁`), E-Mail-Entwurf (`✉️`)
- **Word-Export (`_erstelle_word`)**:
  - Dynamische Grunddaten-Tabelle: Offblock-Zeilen nur wenn Wert vorhanden
  - Personen-Tabelle mit 4 Spalten `[Person, Typ, Kategorie, Anmerkung]`
  - DRK-Logo in Kopfzeile
- **E-Mail-Entwurf (`_email_entwurf_dialog`)**:
  - QDialog mit vorausgefülltem Betreff `Vorkommnisbericht – Flug {flug}`
  - QTextEdit für Freitext-Body
  - QListWidget zeigt `.docx`-Dateien aus Berichts-Ordner (nach Änderungszeit sortiert, neueste oben)
  - Button „✉️ Entwurf erstellen" → `_erstelle_outlook_entwurf()` via `win32com.client`

### `functions/vorkommnisse_db.py`
- SQLite CRUD für `database SQL/vorkommnisse.db` (WAL-Modus)
- Tabelle `vorkommnisse`: `id`, `datum`, `uhrzeit`, `ort`, `flugnummer`, `kategorie`, `beschreibung`, `massnahmen`, `offblock_plan`, `offblock_ist`, `verspaetung_min`, `personen` (JSON), `mitarbeiter`, `erstellt_am`
- `speichern(daten) → int`: Neuen Bericht anlegen, gibt ID zurück
- `aktualisieren(id, daten)`: Bestehenden Bericht überschreiben
- `lade_alle() → list[dict]`: Alle Berichte (neueste zuerst)
- `lade_einen(id) → dict | None`: Einzelnen Bericht
- `loeschen(id)`: Bericht löschen
