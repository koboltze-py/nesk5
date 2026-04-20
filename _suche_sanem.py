import openpyxl
import os
import datetime
from pathlib import Path

BASIS = r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\04_Tagesdienstpläne"
SUCHE = "toptas"

eintraege = []
krank = []

dateien = sorted(Path(BASIS).rglob("*.xlsx"))
print(f"{len(dateien)} Dateien gefunden, durchsuche ...")

for pfad in dateien:
    try:
        wb = openpyxl.load_workbook(pfad, data_only=True, read_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=3, values_only=True):
            name_val = str(row[3] or "")
            if SUCHE not in name_val.lower():
                continue

            datum_val = row[2]
            dienst    = str(row[4] or "").strip()
            beginn    = row[5]
            ende      = row[6]

            def fmt_t(t):
                if isinstance(t, datetime.datetime):
                    return t.strftime("%H:%M")
                if isinstance(t, datetime.time):
                    return t.strftime("%H:%M")
                return str(t) if t else "?"

            def fmt_d(d):
                if isinstance(d, datetime.datetime):
                    return d.strftime("%d.%m.%Y")
                return str(d)[:10] if d else pfad.name

            datum_str = fmt_d(datum_val)
            b_str = fmt_t(beginn)
            e_str = fmt_t(ende)

            eintrag = {
                "datum": datum_str,
                "dienst": dienst,
                "beginn": b_str,
                "ende": e_str,
                "name": name_val.strip(),
            }
            if "krank" in dienst.lower():
                krank.append(eintrag)
            else:
                eintraege.append(eintrag)
        wb.close()
    except Exception as ex:
        print(f"  FEHLER {pfad.name}: {ex}")

print()
print("=" * 65)
print("  SANEM TOPTAS – Dienstübersicht")
print("=" * 65)

if eintraege:
    eintraege_sorted = sorted(eintraege, key=lambda x: x["datum"])
    print(f"  {len(eintraege_sorted)} Arbeitstage:\n")
    for e in eintraege_sorted:
        art = ""
        d = e["dienst"].upper()
        if d in ("T", "T12") or d.startswith("T"):
            art = "Tagdienst"
        elif d in ("N", "DN") or d.startswith("N"):
            art = "Nachtdienst"
        elif d == "EH":
            art = "EH-Dienst"
        else:
            art = d
        print(f"  {e['datum']}  [{e['dienst']:6s}]  {e['beginn']} – {e['ende']}  → {art}")
else:
    print("  Keine Arbeitstage gefunden.")

print()
print("-" * 65)
print(f"  Krankmeldungen gesamt: {len(krank)}")
print("-" * 65)
if krank:
    for k in sorted(krank, key=lambda x: x["datum"]):
        print(f"  {k['datum']}  KRANK  ({k['beginn']} – {k['ende']})")
else:
    print("  Keine Krankmeldungen gefunden.")

print()
print("=" * 65)
print(f"  ZUSAMMENFASSUNG")
print(f"  Arbeitstage: {len(eintraege)}")
print(f"  Kranktage:   {len(krank)}")
print("=" * 65)
