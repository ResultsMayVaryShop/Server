"""
RMV Merch Shop — Cloud Server
Läuft auf Render.com — kein lokaler Laptop nötig, immer online.
Einstellungen: Umgebungsvariablen im Render-Dashboard eintragen.
"""

import os, json, io
from datetime import datetime, timedelta
from http.server import HTTPServer, BaseHTTPRequestHandler

# ── Konfiguration (aus Render-Umgebungsvariablen) ─────────────
PORT              = int(os.environ.get("PORT", 8787))
SMTP_USER         = os.environ.get("EMAIL_ADRESSE", "")
SMTP_PASSWORD     = os.environ.get("EMAIL_PASSWORT", "")
OWNER_EMAIL       = os.environ.get("BENACHRICHTIGUNG_AN", SMTP_USER)
STRIPE_SECRET_KEY = os.environ.get("STRIPE_SECRET_KEY", "")
SHOP_URL          = os.environ.get("SHOP_URL", "")   # z.B. https://rmv-merch.netlify.app

SMTP_SERVER = "smtp-mail.outlook.com"
SMTP_PORT   = 587


# ── In-Memory-Zustand ─────────────────────────────────────────
# Lagerbestand — wird beim Server-Start einmalig geladen und
# dann im Arbeitsspeicher aktualisiert. Zurücksetzen bei Neustart
# ist für diesen kleinen Shop akzeptabel (Emails sind das Backup).

INVENTORY = {
    # Drop 1 — Lagerware (aktuelle Zahlen aus Excel)
    "sweater_drop1_violet":     {"available": 1,   "status": "Wenige übrig",  "sizes": ["M"]},
    "sweater_drop1_khaki":      {"available": 3,   "status": "Wenige übrig",  "sizes": ["M", "L"]},
    "sweater_drop1_naturalraw": {"available": 2,   "status": "Wenige übrig",  "sizes": ["M", "L"]},
    "tshirt_drop1_naturalraw":  {"available": 6,   "status": "Verfügbar",     "sizes": ["S", "M", "L"]},
    # Drop 2 — Auf Bestellung (unbegrenzt)
    "sweater_drop2_naturalraw": {"available": "∞", "status": "Auf Bestellung","sizes": []},
    "sweater_drop2_violet":     {"available": "∞", "status": "Auf Bestellung","sizes": []},
    "sweater_drop2_khaki":      {"available": "∞", "status": "Auf Bestellung","sizes": []},
    "tshirt_drop2_naturalraw":  {"available": "∞", "status": "Auf Bestellung","sizes": []},
}

PENDING_ORDERS = {}   # stripe_session_id → order-dict (in-memory)
NEXT_INV_NUM   = [59] # Liste damit Änderung in Funktionen möglich ist


# ── HTTP-Handler ──────────────────────────────────────────────

class OrderHandler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] " + fmt % args)

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_GET(self):
        if self.path == "/inventory":
            self._respond(200, INVENTORY)
        elif self.path in ("/", "/index.html"):
            self._serve_html()
        else:
            self._respond(200, {"status": "RMV Merch Shop Server läuft \u2713", "version": "2.1-render"})

    def _serve_html(self):
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
        if os.path.exists(html_path):
            with open(html_path, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        else:
            self._respond(404, {"error": "index.html nicht gefunden"})

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        try:
            body = self.rfile.read(length)
            data = json.loads(body)
        except Exception as e:
            self._respond(400, {"error": str(e)})
            return

        if self.path == "/checkout":
            try:
                self._respond(200, create_checkout_session(data))
            except Exception as e:
                import traceback; traceback.print_exc()
                self._respond(500, {"error": str(e)})

        elif self.path == "/verify-payment":
            try:
                self._respond(200, verify_and_process(data))
            except Exception as e:
                import traceback; traceback.print_exc()
                self._respond(500, {"error": str(e)})

        elif self.path == "/waitlist":
            try:
                self._respond(200, handle_waitlist(data))
            except Exception as e:
                import traceback; traceback.print_exc()
                self._respond(500, {"error": str(e)})

        else:
            self._respond(404, {"error": "Unbekannter Endpoint"})

    def _respond(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._cors()
        self.end_headers()
        self.wfile.write(body)


# ── Stripe Checkout ───────────────────────────────────────────

def create_checkout_session(order):
    if not STRIPE_SECRET_KEY:
        raise RuntimeError("STRIPE_SECRET_KEY nicht gesetzt (Render-Umgebungsvariablen)")
    import stripe
    stripe.api_key = STRIPE_SECRET_KEY

    gesamt_cent = int(round(float(order.get("gesamt", 0)) * 100))
    prod_label  = f"{order.get('produkt','')} · {order.get('farbe','')} · Gr. {order.get('groesse','')}"

    # Shop-URL: zuerst aus Order (vom Browser), dann aus Env-Var, dann Fallback
    shop_url = order.pop("shopUrl", None) or SHOP_URL or "http://localhost:8080"

    session = stripe.checkout.Session.create(
        payment_method_types=["card", "paypal", "sepa_debit"],
        line_items=[{
            "price_data": {
                "currency": "eur",
                "product_data": {"name": prod_label, "description": "RMV Merch"},
                "unit_amount": gesamt_cent,
            },
            "quantity": 1,
        }],
        mode="payment",
        customer_email=order.get("email", "") or None,
        success_url=shop_url + "?paid=1&sid={CHECKOUT_SESSION_ID}",
        cancel_url=shop_url + "?cancelled=1",
    )

    # Bestellung im Arbeitsspeicher merken bis Zahlung bestätigt
    PENDING_ORDERS[session.id] = order
    print(f"  ✓ Stripe Session: {session.id}")
    return {"checkoutUrl": session.url}


def verify_and_process(data):
    session_id = data.get("sessionId", "")
    if not session_id:
        raise ValueError("sessionId fehlt")
    if not STRIPE_SECRET_KEY:
        raise RuntimeError("STRIPE_SECRET_KEY nicht gesetzt")

    import stripe
    stripe.api_key = STRIPE_SECRET_KEY
    session = stripe.checkout.Session.retrieve(session_id)

    if session.payment_status != "paid":
        return {"success": False, "reason": "Zahlung noch nicht abgeschlossen"}

    order = PENDING_ORDERS.get(session_id)
    if not order:
        return {"success": False, "reason": "Bestellung nicht gefunden (evtl. schon verarbeitet)"}

    result = process_order(order, paid=True)
    del PENDING_ORDERS[session_id]
    return result


# ── Warteliste ───────────────────────────────────────────────

def handle_waitlist(data):
    """Wartelisten-Eintrag speichern und Email-Benachrichtigung senden."""
    vn    = data.get("vorname", "").strip()
    nn    = data.get("nachname", "").strip()
    em    = data.get("email", "").strip()
    typ   = data.get("type", "event")        # "event" oder "retreat"
    msg   = data.get("nachricht", "").strip()
    datum = data.get("datum", "")

    if not vn or not nn or not em:
        raise ValueError("Pflichtfelder fehlen")

    label = "Event-Warteliste" if typ == "event" else "Retreat-Warteliste"
    print(f"\n✉  {label}: {vn} {nn} <{em}>")

    # Email-Benachrichtigung an Carola
    _send_waitlist_email(vn, nn, em, typ, msg, datum)

    return {"success": True}


def _send_waitlist_email(vn, nn, em, typ, msg, datum):
    """Sendet eine Benachrichtigungs-Email bei neuem Wartelisten-Eintrag."""
    if not SMTP_USER or not SMTP_PASSWORD:
        print("  ⚠  SMTP nicht konfiguriert — Email übersprungen")
        return
    try:
        import smtplib
        from email.message import EmailMessage
        label = "Event" if typ == "event" else "Retreat"
        subject = f"[RMV] Neue {label}-Warteliste: {vn} {nn}"
        body = (
            f"Neuer Wartelisten-Eintrag ({label})\n"
            f"{'─'*40}\n"
            f"Name:    {vn} {nn}\n"
            f"Email:   {em}\n"
            f"Datum:   {datum}\n"
        )
        if msg:
            body += f"Nachricht: {msg}\n"

        eml = EmailMessage()
        eml["From"]    = SMTP_USER
        eml["To"]      = OWNER_EMAIL
        eml["Subject"] = subject
        eml.set_content(body)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.ehlo(); s.starttls(); s.ehlo()
            s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(eml)
        print(f"  ✓ Warteliste-Email gesendet an {OWNER_EMAIL}")
    except Exception as e:
        print(f"  ✗ Email-Fehler: {e}")


# ── Haupt-Logik ───────────────────────────────────────────────

def process_order(order, paid=False):
    name = f"{order.get('vorname','')} {order.get('nachname','')}".strip()
    print(f"\n{'='*50}")
    print(f"  Neue Bestellung: {name} {'(BEZAHLT)' if paid else ''}")
    print(f"  {order.get('produkt')} · {order.get('farbe')} · Gr. {order.get('groesse')}")
    print(f"{'='*50}")

    inv_num = NEXT_INV_NUM[0]
    NEXT_INV_NUM[0] += 1
    inv_str = str(inv_num).zfill(3)

    # Lagerbestand aktualisieren (nur Drop 1)
    if order.get("drop", "") == "Drop 1":
        _update_inventory(order)

    send_emails(order, inv_str, paid=paid)
    print(f"\n✓ Fertig — Rechnung Nr. {inv_str}")
    return {"success": True, "invoiceNumber": inv_str}


def _norm(s):
    return str(s or "").strip().lower().replace(" ", "").replace("-", "").replace("(neu)", "")


def _update_inventory(order):
    prod    = _norm(order.get("produkt", "")).replace("drop1","").replace("drop2","")
    farbe   = _norm(order.get("farbe", ""))
    drop    = _norm(order.get("drop", ""))
    anzahl  = int(order.get("anzahl", 1))

    key = f"{prod}_{drop}_{farbe}"
    if key in INVENTORY and INVENTORY[key]["available"] != "∞":
        old = INVENTORY[key]["available"]
        INVENTORY[key]["available"] = max(0, int(old) - anzahl)
        avail = INVENTORY[key]["available"]
        if avail == 0:
            INVENTORY[key]["status"] = "Ausverkauft"
        elif avail <= 2:
            INVENTORY[key]["status"] = "Wenige übrig"


# ── DOCX Rechnung ────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DOCX = os.path.join(BASE_DIR, "template_rechnung.docx")

def _generate_invoice_docx(order, inv_str, paid, today_str, due_str):
    """Füllt die Rechnungsvorlage aus und gibt DOCX-Bytes zurück."""
    if not os.path.exists(TEMPLATE_DOCX):
        print(f"  ⚠  Rechnungsvorlage nicht gefunden: {TEMPLATE_DOCX}")
        return None
    try:
        import zipfile, io as _io
        with open(TEMPLATE_DOCX, "rb") as f:
            raw = f.read()

        buf_out = _io.BytesIO()
        with zipfile.ZipFile(_io.BytesIO(raw), "r") as zin, \
             zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    xml = data.decode("utf-8", errors="ignore")

                    preis   = str(order.get("preis", ""))
                    versand = float(order.get("versand", 0) or 0)
                    gesamt  = str(order.get("gesamt", ""))
                    prod    = order.get("produkt", "")
                    farbe   = order.get("farbe", "")
                    groesse = order.get("groesse", "")
                    drop    = order.get("drop", "")
                    anzahl  = str(order.get("anzahl", 1))
                    vorname = order.get("vorname", "")
                    nachname= order.get("nachname", "")
                    strasse = order.get("strasse", "—")
                    plz     = order.get("plz", "")
                    stadt   = order.get("stadt", "München")
                    datum   = order.get("datum", today_str)

                    prod_line = f"RMV Merch — {prod} · {farbe} · {drop} · Gr. {groesse}"
                    versand_zeile = (
                        f"inkl. Versand DHL: {versand:.2f} \u20ac"
                        if versand > 0 else "Abholung beim Run Club"
                    )
                    zahlung = (
                        f"Bezahlt am {today_str} \u2713" if paid
                        else f"F\u00e4llig am {due_str}"
                    )

                    for old, new in [
                        ("14. M\u00e4rz 2026",                                   datum),
                        ("011",                                                   inv_str),
                        ("xx",                                                    vorname),
                        ("xx stra\u00dfe",                                        f"{nachname} {strasse}"),
                        ("M\u00fcnchen",                                          f"{plz} {stadt}"),
                        ("F\u00e4llig am 03.04.2026 ",                           zahlung + " "),
                        ("01",                                                    anzahl.zfill(2)),
                        ("RMV Cycling Retreat \u2013 Rider Package (07.\u201310. Mai 2026)", prod_line),
                        ("379 Euro ",                                             f"{preis} Euro "),
                        ("Teilnahme am Results May Vary Cycling Retreat inkl.:", versand_zeile),
                        ("Unterkunft (3 N\u00e4chte)",                            ""),
                        ("Verpflegung (Brunch, Dinner, Snacks &amp; Drinks \u2013 au\u00dfer Lunch w\u00e4hrend der Touren)", ""),
                        ("Organisation &amp; Betreuung",                          ""),
                        ("Gef\u00fchrte Rennrad-Touren / gemeinsame Rides",       ""),
                        ("Yoga / Mobility Sessions",                              ""),
                        ("Community-Programm (DJ, gemeinsame Abende etc.)",       ""),
                        ("Goodie Bag",                                            ""),
                        ("379 ",                                                  f"{gesamt} "),
                    ]:
                        xml = xml.replace(old, new)

                    data = xml.encode("utf-8")
                zout.writestr(item, data)

        print(f"  ✓ DOCX Rechnung generiert: {inv_str}_Merch_{order.get('vorname','')}{order.get('nachname','')}_2026.docx")
        return buf_out.getvalue()
    except Exception as e:
        print(f"  ⚠  DOCX Generierung fehlgeschlagen: {e}")
        import traceback; traceback.print_exc()
        return None


# ── Emails ────────────────────────────────────────────────────

def send_emails(order, inv_str, paid=False):
    if not SMTP_PASSWORD:
        print("  ⚠  Email übersprungen — EMAIL_PASSWORT nicht gesetzt")
        return False

    name        = f"{order.get('vorname','')} {order.get('nachname','')}".strip()
    vorname     = order.get("vorname", "")
    kunde_email = order.get("email", "")
    produkt     = order.get("produkt", "")
    farbe       = order.get("farbe", "")
    groesse     = order.get("groesse", "")
    drop        = order.get("drop", "")
    anzahl      = order.get("anzahl", 1)
    preis       = order.get("preis", 0)
    versand     = order.get("versand", 0)
    gesamt      = order.get("gesamt", 0)
    lieferung   = order.get("lieferung", "")
    runclub     = order.get("runclub", "")
    datum       = order.get("datum", datetime.now().strftime("%d.%m.%Y"))

    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText

    def _send(msg):
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as srv:
            srv.ehlo(); srv.starttls(); srv.login(SMTP_USER, SMTP_PASSWORD)
            srv.send_message(msg)

    # ── 1. Team-Benachrichtigung (Plaintext) ──────────────────
    liefertext = f"Run Club am {runclub}" if runclub else f"Versand nach {order.get('strasse','')}, {order.get('plz','')} {order.get('stadt','')}"
    team_body = f"""🛍️  NEUE BESTELLUNG — RMV Merch Shop
{'='*45}

Rechnung Nr.:  {inv_str}
Datum:         {datum}

KUNDE:
  {name}
  {kunde_email}

BESTELLUNG:
  {anzahl}x  {produkt} · {farbe} · {drop} · Gr. {groesse}
  Preis:   {preis} €
  Versand: {versand} €
  Gesamt:  {gesamt} €

LIEFERUNG:  {liefertext}

💳 Zahlungsstatus: {'BEZAHLT (Stripe) ✓' if paid else 'Ausstehend'}
"""
    if order.get("anmerkung"):
        team_body += f"\nAnmerkung: {order.get('anmerkung')}\n"

    msg1 = MIMEMultipart("alternative")
    msg1["From"]    = SMTP_USER
    msg1["To"]      = OWNER_EMAIL
    msg1["Subject"] = f"[RMV Shop] Nr. {inv_str} — {name} — {produkt} {farbe} Gr.{groesse}"
    msg1.attach(MIMEText(team_body, "plain", "utf-8"))

    try:
        _send(msg1)
        print(f"  ✓ Team-Email → {OWNER_EMAIL}")
    except Exception as e:
        print(f"  ⚠  Team-Email fehlgeschlagen: {e}")

    # ── 2. Kunden-Bestätigung + DOCX-Rechnung ────────────────────
    if not kunde_email:
        return True

    today     = datetime.now()
    due       = today + timedelta(days=14)
    monate    = ["Januar","Februar","März","April","Mai","Juni",
                 "Juli","August","September","Oktober","November","Dezember"]
    today_str = f"{today.day:02d}. {monate[today.month-1]} {today.year}"
    due_str   = due.strftime("%d.%m.%Y")

    msg2 = MIMEMultipart()
    msg2["From"]    = SMTP_USER
    msg2["To"]      = kunde_email
    msg2["Subject"] = f"Deine RMV Merch Rechnung Nr. {inv_str} 🛍️"

    # Kunden-Bestätigungs-Body (plain text)
    confirm_body = f"""Hej {vorname}! 👋

Danke für deine Bestellung — im Anhang findest du deine Rechnung.

  {anzahl}x {produkt} · {farbe} · Gr. {groesse}
  Gesamt: {gesamt} €

{"Deine Zahlung ist eingegangen — alles erledigt! ✓" if paid else "Zahlungsdetails stehen auf der Rechnung."}

Liebe Grüße,
Carola & das RMV Team 🏃
results.mv@outlook.com
"""
    msg2.attach(MIMEText(confirm_body, "plain", "utf-8"))

    # DOCX Rechnung als Anhang
    docx_bytes = _generate_invoice_docx(order, inv_str, paid, today_str, due_str)
    if docx_bytes:
        from email.mime.base import MIMEBase
        from email import encoders as _enc
        docx_filename = f"{inv_str}_Merch_{order.get('vorname','')}{order.get('nachname','')}_2026.docx"
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
        part.set_payload(docx_bytes)
        _enc.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={docx_filename}")
        msg2.attach(part)

    try:
        _send(msg2)
        print(f"  ✓ Kunden-Email → {kunde_email}")
    except Exception as e:
        print(f"  ⚠  Kunden-Email fehlgeschlagen: {e}")

    return True


# ── Start ─────────────────────────────────────────────────────

if __name__ == "__main__":
    missing = []
    if not SMTP_USER:     missing.append("EMAIL_ADRESSE")
    if not SMTP_PASSWORD: missing.append("EMAIL_PASSWORT")
    if not STRIPE_SECRET_KEY: missing.append("STRIPE_SECRET_KEY")

    print(f"""
╔══════════════════════════════════════════╗
║   RMV Merch Shop — Cloud Server v2.1   ║
╚══════════════════════════════════════════╝
  Port:      {PORT}
  Email:     {SMTP_USER or '⚠ NICHT GESETZT'}
  Stripe:    {'✓ konfiguriert' if STRIPE_SECRET_KEY else '⚠ NICHT GESETZT'}
  Shop-URL:  {SHOP_URL or '(wird vom Browser übergeben)'}
{''.join(chr(10)+'  ⚠  Bitte setze: '+v+' (Render-Umgebungsvariablen)' for v in missing)}
  Läuft — Strg+C zum Stoppen
""")

    server = HTTPServer(("0.0.0.0", PORT), OrderHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer gestoppt.")
