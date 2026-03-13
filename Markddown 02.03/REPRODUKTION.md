# Nesk3 – Reproduktionsprotokoll

**Stand:** 12.03.2026 – v3.4.2

---

## Voraussetzungen

| Komponente | Version |
|-----------|---------|
| Python | 3.13+ |
| PySide6 | aktuell |
| openpyxl | aktuell |
| python-docx | aktuell |
| pywin32 | aktuell (Outlook) |

```powershell
pip install PySide6 openpyxl python-docx pywin32
```

---

## Projektstruktur

```
Nesk3/
├── main.py, config.py
├── gui/
│   ├── main_window.py, dashboard.py, mitarbeiter_dokumente.py
│   ├── uebergabe.py, fahrzeuge.py, dienstplan.py
│   ├── dienstliches.py              # Patienten-Station (v3.3.0+)
│   ├── telefonnummern.py            # Telefonnummern-Verzeichnis (v3.2.0+)
│   ├── sonderaufgaben.py            # Sonderaufgaben-Formular (v3.4.0 erweitert)
│   └── ...
├── functions/
│   ├── mitarbeiter_dokumente_functions.py
│   ├── stellungnahmen_db.py
│   ├── stellungnahmen_html_export.py
│   ├── telefonnummern_db.py         # NEU v3.2.0
│   └── ...
├── database/ (connection.py, migrations.py, models.py)
├── backup/ (backup_manager.py)
├── WebNesk/
│   ├── stellungnahmen_lokal.html   ← generiert, kein Server nötig
│   └── ...
└── Daten/
    ├── Mitarbeiterdokumente/
    │   ├── Stellungnahmen/, Bescheinigungen/, ...
    │   ├── Datenbank/stellungnahmen.db
    │   └── Mitarbeiter Vorlagen/Kopf und Fußzeile/...docx
    └── Patienten Station/
        └── Protokolle/              ← Word-Exporte (.docx)
```

---

## Datenbank initialisieren

```python
from database.migrations import run_migrations
run_migrations()
```

Die Stellungnahmen-DB wird automatisch beim ersten Aufruf von `stellungnahmen_db.py` erstellt.  
Die Telefonnummern-DB wird beim ersten Start von `TelefonnummernWidget` angelegt und befüllt.  
Die Patienten-DB (`patienten`, `verbrauchsmaterial`, `medikamente`) wird automatisch migriert – bestehende `patienten`-Tabellen werden per `ALTER TABLE` erweitert.

---

## Anwendung starten

```powershell
python main.py
```

---

## Stellungnahmen-Web-Ansicht

Die HTML-Datei `WebNesk/stellungnahmen_lokal.html` wird automatisch generiert.
Manuell regenerieren:

```python
from functions.stellungnahmen_html_export import generiere_html
generiere_html()
```

Browser-Aufruf: `file:///C:/...Nesk3/WebNesk/stellungnahmen_lokal.html`
Direktlink zu Datensatz 42: `...html#id-42`

---

## Bekannte Probleme

| Problem | Lösung |
|---------|--------|
| PySide6 nicht installierbar | System-Python nutzen |
| Outlook-Mail nicht erstellt | pywin32 installieren, Outlook öffnen |
| Vorlage nicht gefunden | Pfad in mitarbeiter_dokumente_functions.py prüfen |
| HTML-Seite leer | Noch keine Stellungnahmen angelegt oder generiere_html() aufrufen |
| Telefonnummern-Tab leer | Excel-Dateien in Daten/Telefonnummern/ ablegen, „📥 Excel neu einlesen" klicken |
| Patienten-Word-Export fehlt | python-docx installieren; Ordner Daten/Patienten Station/Protokolle/ anlegen |
| Dienstplan in Excel öffnen schlägt fehl | Excel installiert und als Standard-App für .xlsx gesetzt? |

---

## Backup erstellen

```powershell
cd "Nesk3"
python -c "
import os, zipfile
from datetime import datetime
BASE_DIR = os.getcwd()
BACKUP_DIR = os.path.join(BASE_DIR, 'Backup Data')
EXCLUDE = {'__pycache__', '.git', 'Backup Data', 'backup', 'build_tmp', 'Exe'}
os.makedirs(BACKUP_DIR, exist_ok=True)
stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
zip_path = os.path.join(BACKUP_DIR, f'Nesk3_backup_{stamp}.zip')
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(BASE_DIR):
        dirs[:] = [d for d in dirs if d not in EXCLUDE]
        for fname in files:
            full = os.path.join(root, fname)
            try:
                zf.write(full, os.path.relpath(full, BASE_DIR))
            except (PermissionError, OSError):
                pass
print('Backup:', zip_path)
"
```

---

## Notfall-Recovery – Komplettwiederherstellung aus Git

> Der **einzig zuverlässige Weg**, die App vollständig wiederherzustellen, ist das Git-Repository.
> Die MD-Dateien helfen beim Verstehen und Einrichten – aber sie ersetzen nicht den Quellcode.

### Repository

```
https://github.com/koboltze-py/nesk5
```

### Schritt-für-Schritt Recovery

#### 1. Repository klonen (neuer PC / neuer Ordner)

```powershell
git clone https://koboltze-py@github.com/koboltze-py/nesk5.git "Nesk3"
cd "Nesk3"
```

#### 2. Python-Abhängigkeiten installieren

```powershell
pip install PySide6 openpyxl python-docx pywin32
```

#### 3. Pflichtordner anlegen

```powershell
New-Item -ItemType Directory -Force "database SQL"
New-Item -ItemType Directory -Force "Backup Data\db_backups"
New-Item -ItemType Directory -Force "Daten\Hilfe\screenshots"
New-Item -ItemType Directory -Force "Daten\Dienstplan"
New-Item -ItemType Directory -Force "Daten\Ausdrucke"
New-Item -ItemType Directory -Force "json"
```

#### 4. App starten – Datenbanken werden auto-erstellt

```powershell
python main.py
```

Beim ersten Start erstellt `database/migrations.py` alle 8 Datenbankdateien automatisch:

| Datei | Inhalt |
|-------|--------|
| `database SQL/nesk3.db` | Hauptdatenbank (Fahrzeuge, Übergabe, Settings) |
| `database SQL/mitarbeiter.db` | Mitarbeiter-Stammdaten |
| `database SQL/verspaetungen.db` | Verspätungs-Protokolle |
| `database SQL/stellungnahmen.db` | Stellungnahmen |
| `database SQL/telefonnummern.db` | Telefonnummern |
| `database SQL/archiv.db` | Archiv |
| `database SQL/patienten_station.db` | Patienten DRK Station |
| `database SQL/einsaetze.db` | Einsätze |

#### 5. Produktionsdaten wiederherstellen (falls vorhanden)

Falls ein DB-Backup existiert (aus `Backup Data/db_backups/` oder SQL JSON-Backup):

```powershell
# Option A: DB-Dateien direkt kopieren
Copy-Item "backup_DATUM\*.db" "database SQL\"

# Option B: Aus SQL JSON-Backup importieren (App-intern)
# Sidebar → Backup → "SQL-Backup wiederherstellen"
```

### Was ist NICHT im Repository (muss separat gesichert werden)

| Was | Wo liegt es | Wie sichern |
|-----|------------|-------------|
| Datenbankdateien (`*.db`) | `database SQL/` | ZIP-Backup über App oder manuell |
| Dienstplan-Excel | `Daten/Dienstplan/` | OneDrive / manuell |
| Mitarbeiterdokumente | `Daten/Mitarbeiterdokumente/` | OneDrive / manuell |
| Ausdrucke / Word-Protokolle | `Daten/Ausdrucke/` | OneDrive / manuell |
| Einstellungen (Outlook-Signatur etc.) | In `nesk3.db` → `settings`-Tabelle | Enthalten im DB-Backup |

### Schnell-Checkliste Recovery

```
[ ] git clone https://koboltze-py@github.com/koboltze-py/nesk5.git
[ ] pip install PySide6 openpyxl python-docx pywin32
[ ] Ordner anlegen (database SQL, Backup Data, Daten/...)
[ ] DB-Backup einspielen (falls vorhanden)
[ ] python main.py → App startet, DBs werden angelegt
[ ] Einstellungen prüfen (Sidebar → Einstellungen)
[ ] Outlook-Verbindung testen (Übergabe → E-Mail)
```
