"""
Setzt alle SL-Einsätze-Werte auf 0
"""
from database.pax_db import lade_alle_eintraege, speichere_tages_einsaetze

print("="*80)
print("Setze alle SL-Einsätze auf 0")
print("="*80)

# Jahr 2026 laden
eintraege = lade_alle_eintraege(2026)

aktualisiert = 0
schon_null = 0

for eintrag in eintraege:
    datum = eintrag['datum']
    einsaetze = eintrag['einsaetze_zahl']
    
    if einsaetze != 0:
        speichere_tages_einsaetze(datum, 0)
        print(f"✓ {datum}: {einsaetze} → 0")
        aktualisiert += 1
    else:
        schon_null += 1

print("\n" + "="*80)
print(f"Abgeschlossen!")
print(f"  {aktualisiert} Einträge auf 0 gesetzt")
print(f"  {schon_null} Einträge waren bereits 0")
print("="*80)
