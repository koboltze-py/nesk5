"""
Test: Verifiziert dass alle DB-Dateien push_row-Verdrahtung haben.
Nur Source-Inspektion, keine echten DB-Verbindungen.
"""
import sys, os, inspect
sys.path.insert(0, os.path.dirname(__file__))
os.chdir(os.path.dirname(__file__))

results = []

# ── 1. psa_db ──────────────────────────────────────────────────────────────
try:
    import functions.psa_db as m
    assert hasattr(m, '_push'), 'kein _push definiert'
    assert '_push' in inspect.getsource(m.psa_speichern), '_push fehlt in psa_speichern'
    assert '_push' in inspect.getsource(m.psa_aktualisieren), '_push fehlt in psa_aktualisieren'
    assert '_push' in inspect.getsource(m.markiere_psa_gesendet), '_push fehlt in markiere_psa_gesendet'
    results.append(('psa_db.py                ', True, None))
except Exception as e:
    results.append(('psa_db.py                ', False, str(e)))

# ── 2. verspaetung_db ──────────────────────────────────────────────────────
try:
    import functions.verspaetung_db as m
    assert hasattr(m, '_push'), 'kein _push definiert'
    assert '_push' in inspect.getsource(m.verspaetung_speichern), '_push fehlt in verspaetung_speichern'
    assert '_push' in inspect.getsource(m.verspaetung_aktualisieren), '_push fehlt in verspaetung_aktualisieren'
    results.append(('verspaetung_db.py        ', True, None))
except Exception as e:
    results.append(('verspaetung_db.py        ', False, str(e)))

# ── 3. stellungnahmen_db ────────────────────────────────────────────────────
try:
    import functions.stellungnahmen_db as m
    assert hasattr(m, '_push'), 'kein _push definiert'
    assert '_push' in inspect.getsource(m.eintrag_speichern), '_push fehlt in eintrag_speichern'
    results.append(('stellungnahmen_db.py     ', True, None))
except Exception as e:
    results.append(('stellungnahmen_db.py     ', False, str(e)))

# ── 4. call_transcription_db ───────────────────────────────────────────────
try:
    import functions.call_transcription_db as m
    assert hasattr(m, '_push'), 'kein _push definiert'
    assert '_push' in inspect.getsource(m.speichern), '_push fehlt in speichern'
    assert '_push' in inspect.getsource(m.textbaustein_speichern), '_push fehlt in textbaustein_speichern'
    results.append(('call_transcription_db.py ', True, None))
except Exception as e:
    results.append(('call_transcription_db.py ', False, str(e)))

# ── 5. mitarbeiter_functions ───────────────────────────────────────────────
try:
    src_text = open('functions/mitarbeiter_functions.py', encoding='utf-8').read()
    assert 'def _push_ma' in src_text or 'def _push' in src_text, 'kein _push_ma definiert'
    import functions.mitarbeiter_functions as m
    src_erstellen = inspect.getsource(m.mitarbeiter_erstellen)
    src_aktu = inspect.getsource(m.mitarbeiter_aktualisieren)
    assert 'push' in src_erstellen, '_push fehlt in mitarbeiter_erstellen'
    assert 'push' in src_aktu, '_push fehlt in mitarbeiter_aktualisieren'
    results.append(('mitarbeiter_functions.py ', True, None))
except Exception as e:
    results.append(('mitarbeiter_functions.py ', False, str(e)))

# ── 6. uebergabe_functions ─────────────────────────────────────────────────
try:
    import functions.uebergabe_functions as m
    assert hasattr(m, '_push_ue'), 'kein _push_ue definiert'
    assert '_push_ue' in inspect.getsource(m.erstelle_protokoll), '_push_ue fehlt in erstelle_protokoll'
    assert '_push_ue' in inspect.getsource(m.aktualisiere_protokoll), '_push_ue fehlt in aktualisiere_protokoll'
    assert '_push_ue' in inspect.getsource(m.schliesse_protokoll_ab), '_push_ue fehlt in schliesse_protokoll_ab'
    results.append(('uebergabe_functions.py   ', True, None))
except Exception as e:
    results.append(('uebergabe_functions.py   ', False, str(e)))

# ── 7. telefonnummern_db ───────────────────────────────────────────────────
try:
    import functions.telefonnummern_db as m
    assert hasattr(m, '_push'), 'kein _push definiert'
    assert '_push' in inspect.getsource(m.eintrag_speichern), '_push fehlt in eintrag_speichern'
    assert '_push' in inspect.getsource(m.eintrag_aktualisieren), '_push fehlt in eintrag_aktualisieren'
    results.append(('telefonnummern_db.py     ', True, None))
except Exception as e:
    results.append(('telefonnummern_db.py     ', False, str(e)))

# ── 8. gui/dienstliches (text-only, kein GUI-Import) ──────────────────────
try:
    src = open('gui/dienstliches.py', encoding='utf-8').read()
    assert 'def _push_pat' in src, 'kein _push_pat definiert'
    assert 'def _push_ein' in src, 'kein _push_ein definiert'
    assert '_push_pat("patienten"' in src, '_push_pat fehlt in patient_speichern'
    assert '_push_ein("einsaetze"' in src, '_push_ein fehlt in einsatz_speichern'
    # auch patient_aktualisieren
    idx = src.index('def patient_aktualisieren')
    assert '_push_pat' in src[idx:idx+6000], '_push_pat fehlt in patient_aktualisieren'
    # markiere_einsatz_gesendet
    idx2 = src.index('def markiere_einsatz_gesendet')
    assert '_push_ein' in src[idx2:idx2+500], '_push_ein fehlt in markiere_einsatz_gesendet'
    results.append(('gui/dienstliches.py      ', True, None))
except Exception as e:
    results.append(('gui/dienstliches.py      ', False, str(e)))

# ── 9. fahrzeug_functions ──────────────────────────────────────────────────
try:
    import functions.fahrzeug_functions as m
    assert hasattr(m, '_push'), 'kein _push definiert'
    assert '_push' in inspect.getsource(m.erstelle_fahrzeug), '_push fehlt in erstelle_fahrzeug'
    assert '_push' in inspect.getsource(m.aktualisiere_fahrzeug), '_push fehlt in aktualisiere_fahrzeug'
    results.append(('fahrzeug_functions.py    ', True, None))
except Exception as e:
    results.append(('fahrzeug_functions.py    ', False, str(e)))

# ── 10. nesk3.db – archiv_functions (reimport) ──────────────────────────────
try:
    src = open('functions/archiv_functions.py', encoding='utf-8').read()
    idx = src.index('def importiere_aus_archiv')
    assert 'push_row' in src[idx:idx+4000], 'push_row fehlt in importiere_aus_archiv'
    assert 'push_replace_by_fk' in src[idx:idx+4000], 'push_replace_by_fk fehlt in importiere_aus_archiv'
    results.append(('archiv_functions.py      ', True, None))
except Exception as e:
    results.append(('archiv_functions.py      ', False, str(e)))

# ── 11. nesk3.db – settings (PC-spezifisch, kein Sync nötig) ───────────────
results.append(('settings_functions.py    ', True, 'PC-spezifisch, kein Turso-Sync nötig'))

# ── Ausgabe ────────────────────────────────────────────────────────────────
print()
print('=== Turso push_row Wiring-Test ===')
print()
print('  [ DATENBANK ]            Datei / Status')
print('  ─────────────────────────────────────────────────────────────────')

nesk3_files = {
    'fahrzeug_functions.py    ', 'uebergabe_functions.py   ',
    'archiv_functions.py      ', 'settings_functions.py    ',
    'mitarbeiter_functions.py ',
}
other_dbs = {
    'psa_db.py                ', 'verspaetung_db.py        ',
    'stellungnahmen_db.py     ', 'call_transcription_db.py ',
    'telefonnummern_db.py     ', 'gui/dienstliches.py      ',
}

for name, ok, err in results:
    db_tag = '[nesk3.db]     ' if name in nesk3_files else '[eigene DB]    '
    status = 'PASS' if ok else 'FAIL'
    note = f'  ({err})' if err and ok else ''
    line = f'  [{status}]  {db_tag}  {name}{note}'
    if err and not ok:
        line += f'\n              -> {err}'
    print(line)

passed = sum(1 for _, ok, _ in results if ok)
total = len(results)
print(f'\n  {passed}/{total} OK')
import sys as _sys
_sys.exit(0 if passed == total else 1)
