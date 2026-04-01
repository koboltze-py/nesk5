# PAX-Import aus Word-Dokumenten - Zusammenfassung

## Aufgabe
Extraktion der PAX-Zahlen aus Stärkemeldung Word-Dokumenten in drei Ordnern (Januar, Februar, März 2026) und Import in die Datenbank.

## Branch
`pax-aus-word`

## Ergebnis

### Import-Statistik
- **Januar 2026**: 29 erfolgreich, 2 Fehler (31 Dokumente)
- **Februar 2026**: 28 erfolgreich, 0 Fehler (28 Dokumente)
- **März 2026**: 31 erfolgreich, 0 Fehler (31 Dokumente)
- **Gesamt**: 88 Dokumente erfolgreich verarbeitet

### Fehler
1. `Stärkemeldung 01.01.2026 bis 02.01.2026.docx` - Datei beschädigt (Package not found)
2. `Stärkemeldung 03.01.2026 bis 04.01.2026.docx` - Keine Daten gefunden

## Implementierung

### Erstellte Dateien
- `_import_pax_from_word.py` - Hauptskript für den PAX-Import
- `_test_pax_extraction.py` - Analyse-Tool für Dokument-Struktur
- `_test_single_extraction.py` - Einzeltest für Extraktion
- `_verify_pax_import.py` - Verifikation der importierten Daten

### Funktionsweise
1. **Dateinamen-Parsing**: Extrahiert das Von-Datum aus dem Dateinamen
   - Beispiel: `Stärkemeldung 01.03.2026 - 02.03.2026.docx` → `2026-03-01`

2. **PAX-Extraktion**: Sucht in den letzten 30 Absätzen nach dem Muster `- XXX -`
   - Beispiel: `- 191 -` → PAX-Zahl = 191

3. **Datenbank-Update**: Speichert die PAX-Zahl mit `speichere_tages_pax()`
   - Bei erneutem Aufruf wird der Wert überschrieben
   - Bestehende SL-Einsätze-Werte bleiben erhalten

## Verifikation

Stichproben bestätigen erfolgreichen Import:
- 2026-01-02: PAX = 236 ✓
- 2026-02-14: PAX = 248 ✓
- 2026-03-01: PAX = 191 ✓
- 2026-03-29: PAX = 288 ✓ (SL-Einsätze = 4, bereits vorhanden)
- 2026-03-31: PAX = 274 ✓ (SL-Einsätze = 6, bereits vorhanden)

## Verwendung

```bash
# Import ausführen
python _import_pax_from_word.py

# Verifikation
python _verify_pax_import.py
```

## Hinweise
- Das Skript kann jederzeit erneut ausgeführt werden
- Vorhandene PAX-Werte werden überschrieben
- SL-Einsätze-Werte bleiben unberührt
- Die Ordnerpfade sind fest im Skript kodiert
