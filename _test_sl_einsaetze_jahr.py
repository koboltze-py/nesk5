"""
Test: SL-Einsätze Jahressumme im Word-Export (Dashboard-Format)
Erzeugt ein Test-Dokument aus echter Excel-Datei und öffnet es direkt.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from datetime import datetime
from functions.staerkemeldung_dashboard_export import StaerkemeldungDashboardExport
from functions.dienstplan_parser import DienstplanParser

# -- Excel-Datei parsen -------------------------------------------------------
EXCEL = (
    r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V"
    r"\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26"
    r"\04_Tagesdienstpläne\04_April\27.04.2026.xlsx"
)

print(f"Lese Excel: {EXCEL}")
result = DienstplanParser(EXCEL, alle_anzeigen=True).parse()
if not result.get("success"):
    print(f"❌ Fehler beim Parsen: {result}")
    sys.exit(1)

# Datum aus Dateiname ableiten
von = datetime(2026, 4, 27)
bis = datetime(2026, 4, 27)

AUSGABE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_test_sl_einsaetze_jahr_OUTPUT.docx"
)

exporter = StaerkemeldungDashboardExport(
    dienstplan_data          = result,
    ausgabe_pfad             = AUSGABE,
    von_datum                = von,
    bis_datum                = bis,
    pax_zahl                 = 38000,
    bulmor_aktiv             = 5,
    einsaetze_zahl           = 7,
    sl_tag_name              = "",
    sl_nacht_name            = "",
    ausgeschlossene_vollnamen= set(),
)

pfad, warnungen = exporter.export()
print(f"✓ Dokument erstellt: {pfad}")
if warnungen:
    for w in warnungen:
        print(f"  ⚠ {w}")

# Dokument direkt öffnen
os.startfile(pfad)
