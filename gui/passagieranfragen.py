"""
Passagieranfragen-Widget
Passagier-E-Mails verarbeiten, Daten extrahieren und
Antworten per Outlook (win32com) versenden.
"""
from __future__ import annotations

import os
import re
from pathlib import Path

from PySide6.QtCore import Qt, QDateTime
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox, QPushButton,
    QSizePolicy, QSplitter, QTableWidget, QTableWidgetItem, QTextEdit,
    QVBoxLayout, QWidget,
)

from config import BASE_DIR, FIORI_BLUE, FIORI_TEXT

_LOGO_PATH = str(Path(BASE_DIR) / "Daten" / "Email" / "Logo.jpg")


# ── Outlook-Posteingang-Dialog ──────────────────────────────────────────────────────────

class OutlookInboxDialog(QDialog):
    """Zeigt die letzten Posteingang-E-Mails und lädt die gewählte zurück."""

    def __init__(self, parent=None, max_items: int = 75):
        super().__init__(parent)
        self.setWindowTitle("📬  Outlook-Posteingang")
        self.resize(860, 480)
        self.selected_body: str = ""
        self.selected_sender_email: str = ""
        self.selected_sender_name: str = ""
        self._items: list = []   # (datum_str, absender_name, absender_email, betreff, body)

        self._load_mails(max_items)
        self._setup_ui()

    def _load_mails(self, max_items: int):
        try:
            import win32com.client
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except Exception:
                outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            inbox = ns.GetDefaultFolder(6)   # 6 = olFolderInbox
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)   # neueste zuerst

            count = 0
            for msg in messages:
                if count >= max_items:
                    break
                try:
                    datum = str(msg.ReceivedTime)[:16] if msg.ReceivedTime else ""
                    sender_name  = (msg.SenderName or "").strip()
                    sender_email = (msg.SenderEmailAddress or "").strip()
                    # Exchange-interne Adressen (EX:/CN=...) überspringen – SMTP bevorzugen
                    if sender_email.upper().startswith(("EX:", "/O=", "/CN=")):
                        sender_email = ""
                    # Versuche Antwort-An-Adresse als SMTP-Fallback
                    if not sender_email:
                        try:
                            sender_email = msg.ReplyRecipients(1).Address or ""
                        except Exception:
                            pass
                    absender_display = sender_name or sender_email
                    betreff  = msg.Subject or ""
                    # Reinen Text bevorzugen; bei HTML-only Body aus HTMLBody extrahieren
                    body = (msg.Body or "").strip()
                    if not body:
                        import re as _re
                        body = _re.sub(r'<[^>]+>', ' ', msg.HTMLBody or "").strip()
                    self._items.append((datum, absender_display, sender_email, betreff, body))
                    count += 1
                except Exception:
                    pass
        except Exception as exc:
            self._items = []
            self._load_error = str(exc)
        else:
            self._load_error = ""

    def _setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(8)

        if self._load_error:
            err = QLabel(f"⚠️  Outlook konnte nicht gelesen werden:\n{self._load_error}")
            err.setWordWrap(True)
            err.setStyleSheet("color: #c0392b; font-size: 11px;")
            lay.addWidget(err)
        else:
            info = QLabel(
                f"{len(self._items)} E-Mails geladen — "
                "Doppelklick oder Auswählen + OK um eine E-Mail zu übernehmen."
            )
            info.setStyleSheet("color: #555; font-size: 11px;")
            lay.addWidget(info)

        self._table = QTableWidget(len(self._items), 3)
        self._table.setHorizontalHeaderLabels(["Datum", "Von", "Betreff"])
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self._table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        self._table.setStyleSheet("font-size: 11px;")

        for row, (datum, absender_display, _sender_email, betreff, _body) in enumerate(self._items):
            self._table.setItem(row, 0, QTableWidgetItem(datum))
            self._table.setItem(row, 1, QTableWidgetItem(absender_display))
            self._table.setItem(row, 2, QTableWidgetItem(betreff))

        self._table.cellDoubleClicked.connect(self._accept_row)
        lay.addWidget(self._table, stretch=1)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("✅  E-Mail übernehmen")
        btns.button(QDialogButtonBox.StandardButton.Cancel).setText("Abbrechen")
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _accept_row(self, row: int, _col: int):
        self._table.selectRow(row)
        self._on_ok()

    def _on_ok(self):
        rows = self._table.selectionModel().selectedRows()
        if not rows:
            return
        row = rows[0].row()
        _datum, sender_name, sender_email, _betreff, body = self._items[row]
        self.selected_body         = body
        self.selected_sender_email = sender_email
        self.selected_sender_name  = sender_name
        self.accept()


# ── Hilfsfunktion: Deutschen E-Mail-Text in strukturierte Felder aufteilen ──────

def _parse_email_fields(text: str) -> dict:
    """
    Versucht aus einem deutschen E-Mail-Freitext folgende Felder zu extrahieren:
    name, email, flugnummer, datum, rueckflug

    Strategie (Priorität oben → unten):
    1. Explizite Label-Zeilen: "Vorname:", "Nachname:", "Name:", "Von:", "Absender:"
    2. Anredezeilen: "Anrede: Herr" + "Vorname: Hans" + "Nachname: ..."
    3. "Herr/Frau Vorname Nachname" im Fließtext
    4. From-Header / Absenderzeile
    """
    result = {"name": "", "anrede": "", "email": "", "flugnummer": "", "datum": "", "rueckflug": ""}

    # ── Anrede ──────────────────────────────────────────────────────────────────
    m_anr = re.search(r'Anrede[:\s]+(Herr|Frau)', text, re.IGNORECASE)
    if m_anr:
        result["anrede"] = m_anr.group(1).capitalize()

    lines = text.splitlines()

    # ── E-Mail ──────────────────────────────────────────────────────────────
    m = re.search(r'[\w.+\-]+@[\w\-]+\.[a-zA-Z]{2,}', text)
    result["email"] = m.group(0) if m else ""

    # ── Flugnummer  (EW583, FR 1234, LH456, 4U100) ──────────────────────────
    m = re.search(r'\b([A-Z0-9]{2})\s*(\d{3,4})\b', text)
    result["flugnummer"] = (m.group(1) + m.group(2)) if m else ""

    # ── Datum + Rückflug ─────────────────────────────────────────────────────
    _MONATE = {
        'januar': '01', 'februar': '02', 'märz': '03', 'april': '04',
        'mai': '05', 'juni': '06', 'juli': '07', 'august': '08',
        'september': '09', 'oktober': '10', 'november': '11', 'dezember': '12',
    }
    dates: list[str] = []
    # dd.MM.yyyy / dd/MM/yyyy / dd-MM-yyyy
    for raw in re.finditer(r'\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})\b', text):
        d, mo, y = raw.group(1), raw.group(2), raw.group(3)
        if len(y) == 2:
            y = "20" + y
        dates.append(f"{int(d):02d}.{int(mo):02d}.{y}")
    # "12. März 2026" / "12 März 2026"
    if not dates:
        for raw2 in re.finditer(
            r'(\d{1,2})[.\s]+(januar|februar|märz|april|mai|juni|juli|august|'
            r'september|oktober|november|dezember)[,.\s]+(\d{4})',
            text, re.IGNORECASE,
        ):
            mo_key = raw2.group(2).lower()
            dates.append(f"{int(raw2.group(1)):02d}.{_MONATE[mo_key]}.{raw2.group(3)}")
    result["datum"] = dates[0] if dates else ""
    result["rueckflug"] = dates[1] if len(dates) >= 2 else ""

    # Explizites Rückflug-Label überschreibt zweites Datum
    m_r = re.search(r'[Rr]ückflug[:\s]+(.{3,40}?)(?:\n|,|\.|$)', text)
    if m_r:
        result["rueckflug"] = m_r.group(1).strip()

    # ── Name ─────────────────────────────────────────────────────────────────
    # Strategie 1: explizite Zeilen "Vorname: ...", "Nachname: ..."
    vorname = ""
    nachname = ""
    for line in lines:
        if re.match(r'(?:Vorname|First\s*name)[:\s]+(.+)', line, re.IGNORECASE):
            vorname = re.split(r':\s*', line, 1)[-1].strip()
        if re.match(r'(?:Nachname|Last\s*name|Familienname)[:\s]+(.+)', line, re.IGNORECASE):
            nachname = re.split(r':\s*', line, 1)[-1].strip()
    if vorname or nachname:
        result["name"] = f"{vorname} {nachname}".strip()
        return result

    # Strategie 2: Label "Name: ..." in einer Zeile
    m_n = re.search(
        r'(?:^|\n)\s*(?:Name|Passagier)[:\s]+([A-ZÄÖÜ][^\n,]{2,40})',
        text, re.IGNORECASE,
    )
    if m_n:
        result["name"] = m_n.group(1).strip()
        return result

    # Strategie 3: Anrede-Block "Anrede: Herr\nVorname: Hans\nNachname: Muster"
    anrede_idx = -1
    for i, line in enumerate(lines):
        if re.match(r'Anrede\s*:', line, re.IGNORECASE):
            anrede_idx = i
            break
    if anrede_idx >= 0:
        for line in lines[anrede_idx + 1: anrede_idx + 5]:
            lm = re.match(r'(?:Vorname|Nachname)[:\s]+(.+)', line, re.IGNORECASE)
            if lm:
                part = lm.group(1).strip()
                if not vorname:
                    vorname = part
                else:
                    nachname = part
        if vorname or nachname:
            result["name"] = f"{vorname} {nachname}".strip()
            return result

    # Strategie 4: "Herr/Frau Vorname Nachname" im Fließtext
    m_n = re.search(
        r'\b(?:Herr(?:n)?|Frau)\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)+)',
        text,
    )
    if m_n:
        result["name"] = m_n.group(1).strip()
        return result

    # Strategie 5: "Von: Hans Muster <...>" (E-Mail-Header)
    m_n = re.search(
        r'^Von:\s*([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)*)\s*[<\(]',
        text, re.MULTILINE,
    )
    if m_n:
        result["name"] = m_n.group(1).strip()

    return result


# ── Antwort-Vorlagen ───────────────────────────────────────────────────────────

_FLUGDATEN_BITTE = (
    "\n\nFür die Bearbeitung Ihrer Anfrage benötigen wir noch folgende Angaben:\n"
    "• Abflugdatum und -uhrzeit\n"
    "• Flugnummer und Reiseziel\n"
    "• Vor- und Nachname der zu betreuenden Person\n"
    "• Art der Einschränkung (WCH-R, WCH-S oder WCH-C)"
)

_SIG = (
    "\n\nMit freundlichen Grüßen\n\nIhr Team vom PRM-Service\n"
    "Am Köln-Bonn-Airport · Kennedystraße · 51147 Köln\n\n"
    "Telefon: +49 2203 40 - 2323  (24 Stunden täglich erreichbar)\n"
    "E-Mail:  flughafen@drk-koeln.de"
)

ANTWORT_KOMPLETT = (
    "Sehr geehrte Damen und Herren,\n\n"
    "wir haben die Passagiere in unserem System eingetragen.\n\n"
    "Die Buchung des PRM Services erfolgt jedoch nur für den Flughafen Köln / Bonn.\n"
    "Bei Buchungen, die nicht über die Airline erfolgen, kann sich die Airline das Recht "
    "nehmen, den PRM Service vor Ort abzulehnen.\n\n"
    "Um einen reibungslosen Ablauf zu gewährleisten, sind folgende Dinge zu beachten:\n\n"
    "• Bitte kommen Sie mindestens 2–2,5 Stunden vor Abflug am Flughafen Köln/Bonn an.\n"
    "• Melden Sie sich beim zuständigen Check-In-Schalter und weisen Sie darauf hin,\n"
    "  dass Sie den PRM Service in Anspruch nehmen möchten.\n"
    "  Wichtig: Die Bestätigung muss über den Check-In erfolgen, da sonst keine\n"
    "  Abholung am Service-Point stattfindet.\n"
    "• Das Check-In-Personal verweist Sie zu einem Service-Point. Dort holen wir Sie\n"
    "  in der Regel eine Stunde vor Abflug ab. In seltenen Fällen kann es zu leichten "
    "Verzögerungen kommen.\n\n"
    "Bei Fragen stehen wir Ihnen jederzeit zur Verfügung – unser Team ist 24 Stunden "
    "täglich für Sie erreichbar. Gerne beantworten wir Ihre Fragen auch telefonisch."
) + _SIG

ANTWORT_FEHLENDE_DATEN = (
    "Sehr geehrte Damen und Herren,\n\n"
    "gerne sind wir Ihnen bei Ihrem Flug behilflich. Leider liegen uns derzeit noch "
    "nicht alle erforderlichen Informationen vor.\n\n"
    "Für die Organisation des Services benötigen wir noch folgende Angaben:\n\n"
    "• Abflugdatum und -uhrzeit\n"
    "• Flugnummer und Reiseziel\n"
    "• Vor- und Nachname der zu betreuenden Person\n"
    "• Art der Einschränkung (WCH-R, WCH-S oder WCH-C)\n\n"
    "Sobald uns diese Informationen vorliegen, kümmern wir uns umgehend um die "
    "weitere Koordination.\n\n"
    "Unser Team ist 24 Stunden täglich für Sie erreichbar – zögern Sie nicht, "
    "uns auch telefonisch zu kontaktieren."
) + _SIG

ANTWORT_PARKPLATZ = (
    "Sehr geehrte Damen und Herren,\n\n"
    "vielen Dank für Ihre Anfrage.\n\n"
    "Eine Abholung direkt am Parkplatz ist kein Problem. Bitte rufen Sie uns am "
    "Reisetag nochmals an, damit wir die genaue Abholung koordinieren können.\n\n"
    "Unser Team ist 24 Stunden täglich für Sie erreichbar und freut sich auf Ihren Anruf."
) + _SIG

ANTWORT_INFO_SERVICE = (
    "Sehr geehrte Damen und Herren,\n\n"
    "vielen Dank für Ihre Anfrage zum PRM-Service am Flughafen Köln/Bonn.\n\n"
    "Der PRM-Service (Persons with Reduced Mobility) steht allen Reisenden zur "
    "Verfügung, die am Flughafen Unterstützung benötigen – ob aufgrund körperlicher "
    "Einschränkungen, vorübergehender Verletzungen oder aus anderen Gründen.\n\n"
    "So funktioniert der PRM-Service am Flughafen Köln/Bonn:\n\n"
    "1. Anmeldung\n"
    "   Bitte melden Sie den Bedarf möglichst frühzeitig bei Ihrer Airline an.\n"
    "   Die Airline informiert uns direkt über Ihre Anforderung.\n\n"
    "2. Ankunft am Flughafen\n"
    "   Bitte kommen Sie mindestens 2–2,5 Stunden vor Abflug am Flughafen an.\n\n"
    "3. Check-In\n"
    "   Melden Sie sich beim zuständigen Check-In-Schalter und weisen Sie auf den\n"
    "   gebuchten PRM-Service hin. Die Bestätigung durch den Check-In ist zwingend\n"
    "   erforderlich, da sonst keine Abholung am Service-Point erfolgt.\n\n"
    "4. Abholung am Service-Point\n"
    "   Das Check-In-Personal verweist Sie zu einem unserer Service-Points.\n"
    "   Dort holen wir Sie in der Regel rund eine Stunde vor Abflug ab.\n\n"
    "5. Begleitung\n"
    "   Unser Team begleitet Sie durch alle Sicherheitskontrollen, zum Gate und –\n"
    "   falls gewünscht – bis an Bord des Flugzeugs.\n\n"
    "Gerne beantworten wir Ihre Fragen auch telefonisch – unser Team ist "
    "24 Stunden täglich für Sie erreichbar."
) + _SIG


# ── Widget ─────────────────────────────────────────────────────────────────────

class PassagieranfragenWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    # ── UI-Aufbau ──────────────────────────────────────────────────────────────

    def _setup_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 10, 16, 10)
        root.setSpacing(8)

        # Titelzeile
        title = QLabel("✉️  Passagieranfragen")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Light))
        title.setStyleSheet(f"color: {FIORI_TEXT};")
        root.addWidget(title)

        hint = QLabel(
            "E-Mail einfügen → Daten extrahieren → Szenario wählen → Outlook-Entwurf öffnen"
        )
        hint.setStyleSheet("color: #666; font-size: 11px;")
        root.addWidget(hint)

        # Haupt-Splitter: Links Eingabe / Rechts Ausgabe
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, stretch=1)

        splitter.addWidget(self._build_left())
        splitter.addWidget(self._build_right())
        splitter.setSizes([440, 720])

    def _build_left(self) -> QWidget:
        left = QWidget()
        layout = QVBoxLayout(left)
        layout.setContentsMargins(0, 0, 8, 0)
        layout.setSpacing(6)

        # Eingabebereich
        grp_in = QGroupBox("📥  Passagier-E-Mail einfügen")
        grp_in.setStyleSheet(self._grp_style())
        in_lay = QVBoxLayout(grp_in)
        in_lay.setSpacing(6)

        self._text_input = QTextEdit()
        self._text_input.setPlaceholderText(
            "E-Mail-Text des Passagiers hier einfügen …"
        )
        self._text_input.setFont(QFont("Segoe UI", 10))
        in_lay.addWidget(self._text_input)

        btn_row = QHBoxLayout()
        btn_posteingang = QPushButton("📬  Posteingang")
        btn_posteingang.setToolTip("Outlook-Posteingang öffnen und E-Mail auswählen")
        btn_posteingang.setStyleSheet(f"""
            QPushButton {{
                background-color: #5d6d7e;
                color: white; border: none; border-radius: 4px;
                padding: 6px 12px; font-size: 11px; font-weight: bold;
            }}
            QPushButton:hover  {{ background-color: #485460; color: white; }}
            QPushButton:pressed {{ background-color: #34495e; }}
        """)
        btn_posteingang.setMinimumHeight(34)
        btn_posteingang.clicked.connect(self._load_from_inbox)
        btn_row.addWidget(btn_posteingang)

        btn_analyse = QPushButton("🔍  Daten extrahieren")
        btn_analyse.setStyleSheet(self._btn_primary_style())
        btn_analyse.setMinimumHeight(34)
        btn_analyse.clicked.connect(self._extract)
        btn_row.addWidget(btn_analyse)
        in_lay.addLayout(btn_row)

        layout.addWidget(grp_in, stretch=1)

        # Extrahierte Felder
        grp_fields = QGroupBox("📋  Extrahierte Daten  (bearbeitbar)")
        grp_fields.setStyleSheet(self._grp_style())
        f_lay = QVBoxLayout(grp_fields)
        f_lay.setSpacing(5)

        self._f_name       = self._add_field(f_lay, "Name:")

        # Anrede-Combo
        anr_row = QHBoxLayout()
        anr_lbl = QLabel("Anrede:")
        anr_lbl.setFixedWidth(90)
        anr_lbl.setStyleSheet(f"color: {FIORI_TEXT}; font-weight: bold; font-size: 11px;")
        self._f_anrede = QComboBox()
        self._f_anrede.addItems(["–", "Herr", "Frau"])
        self._f_anrede.setMinimumHeight(26)
        self._f_anrede.setStyleSheet("""
            QComboBox { border: 1px solid #c8d0d8; border-radius: 4px;
                        padding: 3px 8px; font-size: 11px; }
            QComboBox:focus { border-color: #0070f3; }
        """)
        anr_row.addWidget(anr_lbl)
        anr_row.addWidget(self._f_anrede)
        f_lay.addLayout(anr_row)

        self._f_email      = self._add_field(f_lay, "E-Mail:")
        self._f_flugnummer = self._add_field(f_lay, "Flugnummer:")
        self._f_datum      = self._add_field(f_lay, "Datum:")
        self._f_rueckflug  = self._add_field(f_lay, "Rückflug:")

        layout.addWidget(grp_fields)
        return left

    def _add_field(self, parent_layout: QVBoxLayout, label: str) -> QLineEdit:
        row = QHBoxLayout()
        lbl = QLabel(label)
        lbl.setFixedWidth(90)
        lbl.setStyleSheet(f"color: {FIORI_TEXT}; font-weight: bold; font-size: 11px;")
        edit = QLineEdit()
        edit.setStyleSheet(self._input_style())
        edit.setMinimumHeight(26)
        row.addWidget(lbl)
        row.addWidget(edit)
        parent_layout.addLayout(row)
        return edit

    def _build_right(self) -> QWidget:
        right = QWidget()
        layout = QVBoxLayout(right)
        layout.setContentsMargins(8, 0, 0, 0)
        layout.setSpacing(6)

        # Szenario-Buttons
        grp_sz = QGroupBox("🎯  Szenario wählen")
        grp_sz.setStyleSheet(self._grp_style())
        sz_lay = QVBoxLayout(grp_sz)
        sz_lay.setSpacing(5)

        _BTN_COLOR = "#1e5799"   # einheitliches Dunkelblau – gut lesbar bei Hover
        _BTN_HOVER  = "#154060"

        scenarios = [
            ("✅  Szenario 1 – Alle Angaben vorhanden",      ANTWORT_KOMPLETT),
            ("⚠️  Szenario 2 – Fehlende Informationen",      ANTWORT_FEHLENDE_DATEN),
            ("🅿️  Szenario 3 – Abholung am Parkplatz",       ANTWORT_PARKPLATZ),
            ("ℹ️  Szenario 4 – Allgemeine PRM-Service-Info", ANTWORT_INFO_SERVICE),
        ]
        for label, text in scenarios:
            btn = QPushButton(label)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {_BTN_COLOR};
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 7px 14px;
                    font-size: 11px;
                    font-weight: bold;
                    text-align: left;
                }}
                QPushButton:hover  {{ background-color: {_BTN_HOVER}; color: #ffffff; }}
                QPushButton:pressed {{ background-color: #0e2d42; }}
            """)
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn.setMinimumHeight(36)
            btn.clicked.connect(lambda _, t=text: self._set_antwort(t))
            sz_lay.addWidget(btn)

        # Checkbox: Flugdaten anfordern
        self._chk_flugdaten = QCheckBox("  + Flugdaten anfordern (Hinweis in Antwort einfügen)")
        self._chk_flugdaten.setStyleSheet(f"color: {FIORI_TEXT}; font-size: 11px;")
        sz_lay.addWidget(self._chk_flugdaten)

        layout.addWidget(grp_sz)

        # Antwort-Bereich
        grp_ant = QGroupBox("📝  Antwort  (bearbeitbar)")
        grp_ant.setStyleSheet(self._grp_style())
        ant_lay = QVBoxLayout(grp_ant)
        ant_lay.setSpacing(6)

        self._text_antwort = QTextEdit()
        self._text_antwort.setPlaceholderText(
            "Szenario wählen oder Antwort hier manuell eingeben …"
        )
        self._text_antwort.setFont(QFont("Segoe UI", 10))
        ant_lay.addWidget(self._text_antwort)

        btn_outlook = QPushButton("📧  Outlook-Entwurf erstellen")
        btn_outlook.setStyleSheet(self._btn_primary_style())
        btn_outlook.setMinimumHeight(38)
        btn_outlook.clicked.connect(self._open_outlook)
        ant_lay.addWidget(btn_outlook)

        layout.addWidget(grp_ant, stretch=1)
        return right

    # ── Logik ──────────────────────────────────────────────────────────────────

    def _load_from_inbox(self):
        """Outlook-Posteingang öffnen, E-Mail auswählen und Felder befüllen."""
        dlg = OutlookInboxDialog(self)
        if dlg.exec() == OutlookInboxDialog.DialogCode.Accepted and dlg.selected_body:
            self._text_input.setPlainText(dlg.selected_body)
            # Zuerst Textextraktion laufen lassen
            self._extract()
            # Absender-E-Mail direkt aus Outlook überschreiben (Regex findet sie nicht im Body)
            if dlg.selected_sender_email:
                self._f_email.setText(dlg.selected_sender_email)
            # Absender-Name ergänzen, falls Extraktion leer war
            if dlg.selected_sender_name and not self._f_name.text().strip():
                self._f_name.setText(dlg.selected_sender_name)

    def _extract(self):
        text = self._text_input.toPlainText()
        fields = _parse_email_fields(text)
        self._f_name.setText(fields["name"])
        anrede = fields.get("anrede", "")
        self._f_anrede.setCurrentText(anrede if anrede in ("Herr", "Frau") else "–")
        self._f_email.setText(fields["email"])
        self._f_flugnummer.setText(fields["flugnummer"])
        self._f_datum.setText(fields["datum"])
        self._f_rueckflug.setText(fields["rueckflug"])

    def _set_antwort(self, template: str):
        text = template.strip()

        # 1. Anrede personalisieren
        name   = self._f_name.text().strip()
        anrede = self._f_anrede.currentText()   # "–", "Herr", "Frau"
        nachname = name.split()[-1] if name else ""

        if nachname and anrede in ("Herr", "Frau"):
            gen = "geehrter" if anrede == "Herr" else "geehrte"
            greeting = f"Sehr {gen} {anrede} {nachname},"
        elif nachname:
            greeting = f"Sehr geehrte/r {nachname},"
        else:
            greeting = "Sehr geehrte Damen und Herren,"

        text = text.replace("Sehr geehrte Damen und Herren,", greeting, 1)

        # 2. Bezug-Zeile mit Flugdaten einfügen (direkt nach Anrede + Leerzeile)
        flug  = self._f_flugnummer.text().strip()
        datum = self._f_datum.text().strip()
        ref_parts = []
        if flug:
            ref_parts.append(f"Flug {flug}")
        if datum:
            ref_parts.append(datum)
        if ref_parts:
            bezug = "Bezug: " + ", ".join(ref_parts) + "\n\n"
            # nach "greeting\n\n" einfügen
            marker = greeting + "\n\n"
            if marker in text:
                text = text.replace(marker, marker + bezug, 1)

        # 3. Flugdaten-Bitte vor der Signatur einfügen
        if self._chk_flugdaten.isChecked():
            sig_marker = "\n\nMit freundlichen Grüßen"
            if sig_marker in text:
                text = text.replace(sig_marker, _FLUGDATEN_BITTE + sig_marker, 1)
            else:
                text += _FLUGDATEN_BITTE

        self._text_antwort.setPlainText(text)

    def _open_outlook(self):
        """Erstellt Outlook-Entwurf via win32com mit DRK-Logo als Signatur."""
        from functions.mail_functions import create_outlook_draft

        recipient = self._f_email.text().strip()
        body = self._text_antwort.toPlainText().strip()

        subject_parts = ["PRM-Service – Flughafen Köln/Bonn"]
        if self._f_name.text().strip():
            subject_parts.append(self._f_name.text().strip())
        if self._f_flugnummer.text().strip():
            subject_parts.append(f"Flug {self._f_flugnummer.text().strip()}")
        if self._f_datum.text().strip():
            subject_parts.append(self._f_datum.text().strip())
        subject = " | ".join(subject_parts)

        logo = _LOGO_PATH if os.path.isfile(_LOGO_PATH) else None

        try:
            create_outlook_draft(
                to=recipient,
                subject=subject,
                body_text=body,
                logo_path=logo,
            )
            QMessageBox.information(
                self,
                "Outlook",
                "Outlook-Entwurf wurde geöffnet.\nBitte prüfen und absenden.",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Outlook-Fehler",
                f"Entwurf konnte nicht erstellt werden:\n{exc}\n\n"
                "Bitte sicherstellen, dass Outlook geöffnet und pywin32 installiert ist.",
            )

    def refresh(self):
        pass

    # ── Styles ─────────────────────────────────────────────────────────────────

    def _grp_style(self) -> str:
        return f"""
            QGroupBox {{
                border: 1px solid #c8d0d8;
                border-radius: 6px;
                margin-top: 14px;
                padding-top: 6px;
                font-size: 11px;
                font-weight: bold;
                color: {FIORI_TEXT};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 10px;
                padding: 0 6px;
                color: {FIORI_BLUE};
            }}
        """

    def _btn_primary_style(self) -> str:
        return f"""
            QPushButton {{
                background-color: {FIORI_BLUE};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 7px 18px;
                font-size: 12px;
                font-weight: bold;
            }}
            QPushButton:hover  {{ background-color: #1a5276; color: white; }}
            QPushButton:pressed {{ background-color: #154360; }}
        """

    def _input_style(self) -> str:
        return f"""
            QLineEdit {{
                border: 1px solid #c8d0d8;
                border-radius: 4px;
                padding: 3px 8px;
                font-size: 11px;
                color: {FIORI_TEXT};
            }}
            QLineEdit:focus {{ border-color: {FIORI_BLUE}; }}
        """
