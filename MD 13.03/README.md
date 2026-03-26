# Nesk3

**DRK Flughafen Köln – Erste-Hilfe-Station**  
Dienstplan-Verwaltung, Stärkemeldung und Mitarbeiterverwaltung

**Version:** v3.6.0 (26.03.2026)  
**Module:** Dashboard · Mitarbeiter · Dienstliches · Aufgaben Tag/Nacht · Dienstplan · Übergabe · Fahrzeuge · Code 19 · **Telefonnummern** · Backup · Einstellungen · Hilfe · Passagieranfragen · **Schulungen**

### Neu in v3.6.0
- **Schulungen – Mitarbeiter-Liste** (neuer Tab neben Kalender): Suche nach Name, Filter nach Status & Schulungstyp; Matrix EH/Refresher/ZÜP/Ärztl./FS-K. mit Farbkodierung; Mitarbeiter ohne Einträge gesondert grau
- **Schulungen – Detailansicht**: Doppelklick auf MA → alle 14 Schulungstypen; fehlende Einträge leer/grau
- **Schulungen – Datum bearbeiten**: ✏️-Button oder Doppelklick je Schulungstyp → Datumspicker, automatische Gültig-bis-Berechnung, direktes Speichern in DB

### Neu in v3.5.1
- **Tab-Design harmonisiert**: Alle 11 GUI-Dateien nutzen einheitliches `#1565a8`-Blau, konsistente 3px/2px Underlines, Segoe UI, Hover-States
- **Fade-Animation**: Sanfter 180ms Fade-In (OutCubic) bei jedem Seitenwechsel in der Sidebar-Navigation
- **Mitarbeiter – Verwaltung**: Tab „Dokumente“ → „Verwaltung“, Ausdrucke + Krankmeldungen jetzt als Sidebar-Kategorien mit DokumentBrowser
- **Sonderaufgaben**: Treeview-Header „Dienstpläne“, neuer „Ordner öffnen“-Button, „Wiederherstellen“-Button mit Dateiauswahl lädt gespeicherte Sonderaufgaben zurück ins Formular

### Neu in v3.5.0
- **Passagieranfragen** (neues Sidebar-Modul): Outlook-Posteingang direkt in der App lesen, E-Mail auswählen → Daten automatisch extrahieren (Name inkl. Anrede, E-Mail direkt aus Outlook, Flugnummer, Datum)
- **Personalisierte Antworten**: Anrede „Sehr geehrter Herr / Sehr geehrte Frau" + Bezug-Zeile mit Flugdaten werden automatisch eingefügt
- **4 Szenarien**: Alle Angaben vorhanden / Fehlende Daten / Parkplatz-Abholung / PRM-Info
- **„+ Flugdaten anfordern"-Checkbox**: fügt bei jedem Szenario Bullet-Liste der fehlenden Felder ein
- **Outlook-Entwurf** via win32com mit DRK-Logo und automatischer Signatur

### Neu in v3.4.5
- **Sidebar – Animiertes Logo**: `_NeskLogoWidget` mit rotierenden Ringen, Shimmer-Effekt und genau passender Sidebar-Farbe (`#354a5e`), keine Zierstreifen
- **Sidebar scrollbar**: `QScrollArea` verhindert, dass Logo und Buttons bei kleinem Fenster überdeckt werden
- **Übergabe – HTML-E-Mail** komplett überarbeitet: DRK-roter Header-Banner, Info-Tabelle, farbige Abschnitts-Boxen
- **Übergabe – Fahrzeuge in E-Mail**: nur KZ + Notiz (Status entfernt)
- **Übergabe – Neue Sektion** „Patienten DRK Station“ mit eigenen Checkboxen

### Neu in v3.4.4
- **Dienstplan Word-Export**: Speicherort-Button entfernt – Speicherdialog öffnet direkt beim Klick auf „Exportieren“
- **Dienstplan Word-Export**: Kein doppeltes Speichern mehr – zweiter „Kopie speichern“-Dialog entfernt

### Neu in v3.4.2
- **Übergabe – Verspätungen**: 3 Bugs behoben (Vortag nach Speichern sichtbar, Auto-Einträge werden persistiert, kein Duplikat-Problem)
- **Übergabe – E-Mail**: Datum-von/bis-Filter für verspätete Mitarbeiter
- **Sonderaufgaben**: Namen-Übertragung auf Vorlage repariert
- **Einsätze / Patienten**: Sortierung aufsteigend – neue Einträge am Ende; keine Pflichtfelder mehr

### Neu in v3.4.1
- **Hilfe – Screenshot-Galerie**: Neuer Tab „📸 Vorschau" im Hilfe-Dialog mit allen 14 Seiten als klickbare Kacheln
- **Hilfe – Vollbild-Vorschau**: Klick auf Kachel öffnet Screenshot in Vollbild (maximierbar)
- **Screenshots automatisch erstellen**: Schaltfläche durchläuft alle Seiten und speichert PNG-Dateien in `Daten/Hilfe/screenshots/`
- **Benutzeranleitung**: neue `docs/BENUTZERANLEITUNG.md` mit 17 Abschnitten, Mockups und Diagrammen

### Neu in v3.4.0
- **Medikamentengabe als Tabelle**: Medikament, Dosis, Applikation – direkt im Patientendialog erfassbar (statt Checkbox)
- **Sonderaufgaben – Bulmor**: Option „a.D." immer verfügbar; Fahrzeugstatus-Badge in der Tabelle
- **Sonderaufgaben – Dienstplan öffnen**: Button öffnet geladene Dienstplan-Datei direkt in Excel
- **Dienstplan-Tab**: „In Excel öffnen"-Button in jedem Dienstplan-Pane
- **Stärkemeldungs-Export**: nach Export Frage „Jetzt in Word öffnen?" + optionaler zweiter Speicherort

### Neu in v3.3.0
- **Patienten DRK Station**: vollständiges medizinisches Protokoll (ABCDE, Monitoring, Arbeitsunfall)
- **Word-Export**: Protokoll als formatiertes .docx mit DRK-Logo
- **E-Mail**: Protokoll als Outlook-Entwurf mit Word-Anhang versendbar

---

## Starten

```powershell
cd "...\Nesk\Nesk3"
python main.py
```

Erfordert Python 3.13+ und folgende Pakete:
```
PySide6, openpyxl, python-docx
```

## Backup erstellen

```python
from backup.backup_manager import create_zip_backup
zip_pfad = create_zip_backup()
print(f"Backup: {zip_pfad}")
```

Oder direkt per Skript:
```powershell
python C:\Users\DRKairport\AppData\Local\Temp\do_backup.py
```

**Ausgeschlossen:** `Backup Data/`, `build_tmp/`, `Exe/`, `__pycache__/` → Größe ~8 MB

## Backup wiederherstellen

```python
from backup.backup_manager import restore_from_zip
restore_from_zip(r"...\Backup Data\Nesk3_backup_YYYYMMDD_HHMMSS.zip")
```

## Dokumentation

→ [DOKUMENTATION.md](DOKUMENTATION.md)  
→ [CHANGELOG.md](CHANGELOG.md)
