"""
RMV Merch Shop â Cloud Server
LÃ¤uft auf Render.com â kein lokaler Laptop nÃ¶tig, immer online.
Einstellungen: Umgebungsvariablen im Render-Dashboard eintragen.
"""

import os, json, io
from datetime import datetime, timedelta
from http.server import HTTPServer, BaseHTTPRequestHandler

# ââ Konfiguration (aus Render-Umgebungsvariablen) âââââââââââââ
PORT              = int(os.environ.get("PORT", 8787))
SMTP_USER         = os.environ.get("EMAIL_ADRESSE", "")
SMTP_PASSWORD     = os.environ.get("EMAIL_PASSWORT", "")
OWNER_EMAIL       = os.environ.get("BENACHRICHTIGUNG_AN", SMTP_USER)
STRIPE_SECRET_KEY = os.environ.get("STRIPE_SECRET_KEY", "")
SHOP_URL          = os.environ.get("SHOP_URL", "")   # z.B. https://rmv-merch.netlify.app

SMTP_SERVER = "smtp-mail.outlook.com"
SMTP_PORT   = 587


# ââ In-Memory-Zustand âââââââââââââââââââââââââââââââââââââââââ
# Lagerbestand â wird beim Server-Start einmalig geladen und
# dann im Arbeitsspeicher aktualisiert. ZurÃ¼cksetzen bei Neustart
# ist fÃ¼r diesen kleinen Shop akzeptabel (Emails sind das Backup).

INVENTORY = {
    # Drop 1 â Lagerware (aktuelle Zahlen aus Excel)
    "sweater_drop1_violet":     {"available": 1,   "status": "Wenige Ã¼brig",  "sizes": ["M"]},
    "sweater_drop1_khaki":      {"available": 3,   "status": "Wenige Ã¼brig",  "sizes": ["M", "L"]},
    "sweater_drop1_naturalraw": {"available": 2,   "status": "Wenige Ã¼brig",  "sizes": ["M", "L"]},
    "tshirt_drop1_naturalraw":  {"available": 6,   "status": "VerfÃ¼gbar",     "sizes": ["S", "M", "L"]},
    # Drop 2 â Auf Bestellung (unbegrenzt)
    "sweater_drop2_naturalraw": {"available": "â", "status": "Auf Bestellung","sizes": []},
    "sweater_drop2_violet":     {"available": "â", "status": "Auf Bestellung","sizes": []},
    "sweater_drop2_khaki":      {"available": "â", "status": "Auf Bestellung","sizes": []},
    "tshirt_drop2_naturalraw":  {"available": "â", "status": "Auf Bestellung","sizes": []},
}

PENDING_ORDERS = {}   # stripe_session_id â order-dict (in-memory)

# ââ Rechnungsnummer persistent in Datei speichern ââââââââââââ
_INV_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inv_counter.txt")

def _load_inv_num():
    try:
        if os.path.exists(_INV_FILE):
            with open(_INV_FILE, "r") as f:
                return int(f.read().strip())
    except Exception:
        pass
    return 59  # Startwert falls keine Datei vorhanden

def _save_inv_num(n):
    try:
        with open(_INV_FILE, "w") as f:
            f.write(str(n))
    except Exception as e:
        print(f"  â   Konnte Rechnungsnummer nicht speichern: {e}")

NEXT_INV_NUM = [_load_inv_num()]  # Liste damit Ãnderung in Funktionen mÃ¶glich ist


# ââ HTTP-Handler ââââââââââââââââââââââââââââââââââââââââââââââ

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
            self._respond(200, {"status": "RMV Merch Shop Server lÃ¤uft \u2713", "version": "2.1-render"})

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
            except ValueError as e:
                msg = str(e)
                if msg.startswith("SOLD_OUT:"):
                    # Freundliche Fehlermeldung fÃ¼r ausverkaufte GrÃ¶Ãen
                    self._respond(409, {"error": msg[9:]})
                else:
                    self._respond(400, {"error": msg})
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


# ââ Stripe Checkout âââââââââââââââââââââââââââââââââââââââââââ

def create_checkout_session(order):
    if not STRIPE_SECRET_KEY:
        raise RuntimeError("STRIPE_SECRET_KEY nicht gesetzt (Render-Umgebungsvariablen)")
    import stripe
    stripe.api_key = STRIPE_SECRET_KEY

    # Shop-URL: zuerst aus Order (vom Browser), dann aus Env-Var, dann Fallback
    shop_url = order.pop("shopUrl", None) or SHOP_URL or "http://localhost:8080"

    # ââ GrÃ¶Ãen-Validierung (nur Drop 1 â Drop 2 ist immer auf Bestellung) ââ
    cart_items = order.get("cart", [])
    for item in cart_items:
        drop    = item.get("drop", "")
        # Drop 2 ist immer unbegrenzt verfÃ¼gbar â Ã¼berspringen
        if "2" in str(drop):
            continue
        produkt = item.get("produkt", "")
        farbe   = item.get("farbe", "").lower().replace(" ", "")
        groesse = item.get("groesse", "")
        # Inventory-Key fÃ¼r Drop 1 suchen (z.B. sweater_drop1_violet)
        inv_key = None
        for key in INVENTORY:
            if "drop1" in key and farbe in key and groesse:
                inv_key = key
                break
        if inv_key and inv_key in INVENTORY:
            inv = INVENTORY[inv_key]
            avail_sizes = inv.get("sizes", [])
            # PrÃ¼fen ob gewÃ¤hlte GrÃ¶Ãe noch verfÃ¼gbar ist
            if isinstance(avail_sizes, list) and avail_sizes and groesse not in avail_sizes:
                raise ValueError(
                    f"SOLD_OUT:{produkt} in GrÃ¶Ãe {groesse} ist leider ausverkauft. "
                    f"Noch verfÃ¼gbar: {', '.join(avail_sizes) if avail_sizes else 'keine'}."
                )

    # Stripe Line Items: einzelne Artikel wenn vorhanden, sonst Gesamt
    cart_items = order.get("cart", [])
    versand    = float(order.get("versand", 0) or 0)

    if cart_items:
        line_items = []
        for item in cart_items:
            name = f"RMV {item.get('produkt','')} Â· {item.get('farbe','')} Â· Gr. {item.get('groesse','')}"
            cent = int(round(float(item.get("preis", 0)) * 100))
            line_items.append({
                "price_data": {
                    "currency": "eur",
                    "product_data": {"name": name, "description": "RMV Merch"},
                    "unit_amount": cent,
                },
                "quantity": int(item.get("anzahl", 1)),
            })
        if versand > 0:
            line_items.append({
                "price_data": {
                    "currency": "eur",
                    "product_data": {"name": "Versand (DHL)"},
                    "unit_amount": int(round(versand * 100)),
                },
                "quantity": 1,
            })
    else:
        # Fallback: Gesamt als einzelne Position
        gesamt_cent = int(round(float(order.get("gesamt", 0)) * 100))
        prod_label  = f"{order.get('produkt','')} Â· {order.get('farbe','')} Â· Gr. {order.get('groesse','')}"
        line_items = [{
            "price_data": {
                "currency": "eur",
                "product_data": {"name": prod_label, "description": "RMV Merch"},
                "unit_amount": gesamt_cent,
            },
            "quantity": 1,
        }]

    # Wichtige Bestelldaten als Stripe-Metadata sichern (Backup bei Server-Neustart)
    meta = {
        "vorname":  str(order.get("vorname",  ""))[:100],
        "nachname": str(order.get("nachname", ""))[:100],
        "email":    str(order.get("email",    ""))[:200],
        "produkt":  str(order.get("produkt",  ""))[:200],
        "farbe":    str(order.get("farbe",    ""))[:100],
        "groesse":  str(order.get("groesse",  ""))[:50],
        "drop":     str(order.get("drop",     ""))[:50],
        "anzahl":   str(order.get("anzahl",   1)),
        "preis":    str(order.get("preis",    0)),
        "gesamt":   str(order.get("gesamt",   0)),
        "versand":  str(order.get("versand",  0)),
        "strasse":  str(order.get("strasse",  ""))[:200],
        "plz":      str(order.get("plz",      ""))[:20],
        "stadt":    str(order.get("stadt",    ""))[:100],
        "lieferung":str(order.get("lieferung",""))[:200],
        "runclub":  str(order.get("runclub",  ""))[:100],
        "cart_json":json.dumps(order.get("cart", []), ensure_ascii=False)[:490],
    }

    session = stripe.checkout.Session.create(
        payment_method_types=["card"],
        line_items=line_items,
        mode="payment",
        customer_email=order.get("email", "") or None,
        success_url=shop_url + "?paid=1&sid={CHECKOUT_SESSION_ID}",
        cancel_url=shop_url + "?cancelled=1",
        metadata=meta,
    )

    # Bestellung im Arbeitsspeicher merken bis Zahlung bestÃ¤tigt
    PENDING_ORDERS[session.id] = order
    print(f"  â Stripe Session: {session.id}")
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
        # Fallback: Bestelldaten aus Stripe-Metadata wiederherstellen (nach Server-Neustart)
        meta = session.metadata or {}
        if meta.get("vorname"):
            print(f"  â¹  Order aus Stripe-Metadata wiederhergestellt (Server-Neustart)")
            order = {
                "vorname":  meta.get("vorname",  ""),
                "nachname": meta.get("nachname", ""),
                "email":    meta.get("email",    ""),
                "produkt":  meta.get("produkt",  ""),
                "farbe":    meta.get("farbe",    ""),
                "groesse":  meta.get("groesse",  ""),
                "drop":     meta.get("drop",     ""),
                "anzahl":   int(meta.get("anzahl",  1)),
                "preis":    float(meta.get("preis",  0)),
                "gesamt":   float(meta.get("gesamt", 0)),
                "versand":  float(meta.get("versand",0)),
                "strasse":  meta.get("strasse",  ""),
                "plz":      meta.get("plz",      ""),
                "stadt":    meta.get("stadt",    "MÃ¼nchen"),
                "lieferung":meta.get("lieferung",""),
                "runclub":  meta.get("runclub",  ""),
            }
        else:
            return {"success": False, "reason": "Bestellung nicht gefunden (evtl. schon verarbeitet)"}

    result = process_order(order, paid=True)
    if session_id in PENDING_ORDERS:
        del PENDING_ORDERS[session_id]
    return result


# ââ Warteliste âââââââââââââââââââââââââââââââââââââââââââââââ

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
    print(f"\nâ  {label}: {vn} {nn} <{em}>")

    # Backup: In Textdatei loggen (Render-Filesystem, leert sich bei Redeploy)
    _log_waitlist(vn, nn, em, typ, msg, datum)

    # Email-Benachrichtigung an Carola
    _send_waitlist_email(vn, nn, em, typ, msg, datum)

    return {"success": True}


def _log_waitlist(vn, nn, em, typ, msg, datum):
    """Schreibt Wartelisten-Eintrag in waitlist.txt als Backup."""
    try:
        log_path = os.path.join(BASE_DIR, "waitlist.txt")
        line = f"{datum} | {typ.upper()} | {vn} {nn} | {em}"
        if msg:
            line += f" | {msg}"
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")
        print(f"  â Warteliste geloggt â waitlist.txt")
    except Exception as e:
        print(f"  â Log-Fehler: {e}")


def _send_waitlist_email(vn, nn, em, typ, msg, datum):
    """Sendet eine Benachrichtigungs-Email bei neuem Wartelisten-Eintrag."""
    if not SMTP_USER or not SMTP_PASSWORD:
        print("  â   SMTP nicht konfiguriert â Email Ã¼bersprungen")
        return
    try:
        import smtplib
        from email.message import EmailMessage
        label = "Event" if typ == "event" else "Retreat"
        subject = f"[RMV] Neue {label}-Warteliste: {vn} {nn}"
        body = (
            f"Neuer Wartelisten-Eintrag ({label})\n"
            f"{'â'*40}\n"
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
        print(f"  â Warteliste-Email gesendet an {OWNER_EMAIL}")

        # BestÃ¤tigungs-Email an Person, die sich eingetragen hat
        confirm_eml = EmailMessage()
        confirm_eml["From"]    = SMTP_USER
        confirm_eml["To"]      = em
        confirm_eml["Subject"] = f"Du bist auf der {label} â"
        confirm_eml.set_content(
            f"Hej {vn}! ð\n\n"
            f"Du bist auf unserer {label} eingetragen â wir melden uns,\n"
            f"sobald es Neuigkeiten gibt!\n\n"
            f"Liebe GrÃ¼Ãe,\nCarola & das RMV Team ð\n"
            f"results.mv@outlook.com\n"
        )
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.ehlo(); s.starttls(); s.ehlo()
            s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(confirm_eml)
        print(f"  â Warteliste-BestÃ¤tigung â {em}")
    except Exception as e:
        print(f"  â Email-Fehler: {e}")


# ââ Haupt-Logik âââââââââââââââââââââââââââââââââââââââââââââââ

def process_order(order, paid=False):
    name = f"{order.get('vorname','')} {order.get('nachname','')}".strip()
    print(f"\n{'='*50}")
    print(f"  Neue Bestellung: {name} {'(BEZAHLT)' if paid else ''}")
    print(f"  {order.get('produkt')} Â· {order.get('farbe')} Â· Gr. {order.get('groesse')}")
    print(f"{'='*50}")

    inv_num = NEXT_INV_NUM[0]
    NEXT_INV_NUM[0] += 1
    _save_inv_num(NEXT_INV_NUM[0])  # Persistent speichern
    inv_str = str(inv_num).zfill(3)

    # Lagerbestand aktualisieren (nur Drop 1 â Cart-Items einzeln verarbeiten)
    cart_items_inv = order.get("cart", [])
    if cart_items_inv:
        for ci in cart_items_inv:
            if "2" not in str(ci.get("drop", "")):
                _update_inventory(ci)
    elif "2" not in str(order.get("drop", "")):
        _update_inventory(order)

    send_emails(order, inv_str, paid=paid)
    print(f"\nâ Fertig â Rechnung Nr. {inv_str}")
    return {"success": True, "invoiceNumber": inv_str}


def _norm(s):
    return str(s or "").strip().lower().replace(" ", "").replace("-", "").replace("(new)", "")


def _update_inventory(order):
    prod    = _norm(order.get("produkt", "")).replace("drop1","").replace("drop2","")
    farbe   = _norm(order.get("farbe", ""))
    drop    = _norm(order.get("drop", ""))
    anzahl  = int(order.get("anzahl", 1))

    key = f"{prod}_{drop}_{farbe}"
    if key in INVENTORY and INVENTORY[key]["available"] != "â":
        old = INVENTORY[key]["available"]
        INVENTORY[key]["available"] = max(0, int(old) - anzahl)
        avail = INVENTORY[key]["available"]
        if avail == 0:
            INVENTORY[key]["status"] = "Ausverkauft"
        elif avail <= 2:
            INVENTORY[key]["status"] = "Wenige Ã¼brig"


# ââ DOCX Rechnung ââââââââââââââââââââââââââââââââââââââââââââ

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DOCX = os.path.join(BASE_DIR, "template_rechnung.docx")

def _generate_invoice_docx(order, inv_str, paid, today_str, due_str):
    """FÃ¼llt die Rechnungsvorlage aus und gibt DOCX-Bytes zurÃ¼ck."""
    if not os.path.exists(TEMPLATE_DOCX):
        print(f"  â   Rechnungsvorlage nicht gefunden: {TEMPLATE_DOCX}")
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
                    strasse = order.get("strasse", "â")
                    plz     = order.get("plz", "")
                    stadt   = order.get("stadt", "MÃ¼nchen")
                    datum   = order.get("datum", today_str)

                    prod_line = f"RMV Merch â {prod} Â· {farbe} Â· {drop} Â· Gr. {groesse}"
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
                        ("yy",                                                    vorname),
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

        print(f"  â DOCX Rechnung generiert: {inv_str}_Merch_{order.get('vorname','')}{order.get('nachname','')}_2026.docx")
        return buf_out.getvalue()
    except Exception as e:
        print(f"  â   DOCX Generierung fehlgeschlagen: {e}")
        import traceback; traceback.print_exc()
        return None


# ââ Emails ââââââââââââââââââââââââââââââââââââââââââââââââââââ

def send_emails(order, inv_str, paid=False):
    if not SMTP_PASSWORD:
        print("  â   Email Ã¼bersprungen â EMAIL_PASSWORT nicht gesetzt")
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

    # ââ 1. Team-Benachrichtigung (Plaintext) ââââââââââââââââââ
    liefertext = f"Run Club am {runclub}" if runclub else f"Versand nach {order.get('strasse','')}, {order.get('plz','')} {order.get('stadt','')}"

    # Artikel-Liste: Cart-Items wenn vorhanden, sonst Einzelartikel
    cart_items = order.get("cart", [])
    if cart_items:
        artikel_zeilen = "\n".join(
            f"  {item.get('anzahl',1)}x  {item.get('produkt','')} Â· {item.get('farbe','')} Â· {item.get('drop','')} Â· Gr. {item.get('groesse','')}  ({item.get('preis',0)} â¬)"
            for item in cart_items
        )
    else:
        artikel_zeilen = f"  {anzahl}x  {produkt} Â· {farbe} Â· {drop} Â· Gr. {groesse}  ({preis} â¬)"

    team_body = f"""ðï¸  NEUE BESTELLUNG â RMV Merch Shop
{'='*45}

Rechnung Nr.:  {inv_str}
Datum:         {datum}

KUNDE:
  {name}
  {kunde_email}

BESTELLUNG:
{artikel_zeilen}
  ââââââââââââââââââââââââââââââ
  Versand: {versand} â¬
  Gesamt:  {gesamt} â¬

LIEFERUNG:  {liefertext}

ð³ Zahlungsstatus: {'BEZAHLT (Stripe) â' if paid else 'Ausstehend'}
"""
    if order.get("anmerkung"):
        team_body += f"\nAnmerkung: {order.get('anmerkung')}\n"

    msg1 = MIMEMultipart("alternative")
    msg1["From"]    = SMTP_USER
    msg1["To"]      = OWNER_EMAIL
    msg1["Subject"] = f"[RMV Shop] Nr. {inv_str} â {name} â {produkt} {farbe} Gr.{groesse}"
    msg1.attach(MIMEText(team_body, "plain", "utf-8"))

    try:
        _send(msg1)
        print(f"  â Team-Email â {OWNER_EMAIL}")
    except Exception as e:
        print(f"  â   Team-Email fehlgeschlagen: {e}")

    # ââ 2. Kunden-BestÃ¤tigung + DOCX-Rechnung ââââââââââââââââââââ
    if not kunde_email:
        return True

    today     = datetime.now()
    due       = today + timedelta(days=14)
    monate    = ["Januar","Februar","MÃ¤rz","April","Mai","Juni",
                 "Juli","August","September","Oktober","November","Dezember"]
    today_str = f"{today.day:02d}. {monate[today.month-1]} {today.year}"
    due_str   = due.strftime("%d.%m.%Y")

    msg2 = MIMEMultipart()
    msg2["From"]    = SMTP_USER
    msg2["To"]      = kunde_email
    msg2["Subject"] = f"Deine RMV Merch Rechnung Nr. {inv_str} ðï¸"

    # Kunden-BestÃ¤tigungs-Body (plain text)
    if cart_items:
        artikel_confirm = "\n".join(
            f"  {item.get('anzahl',1)}x {item.get('produkt','')} Â· {item.get('farbe','')} Â· Gr. {item.get('groesse','')}"
            for item in cart_items
        )
    else:
        artikel_confirm = f"  {anzahl}x {produkt} Â· {farbe} Â· Gr. {groesse}"

    confirm_body = f"""Hej {vorname}! ð

Danke fÃ¸r deine Bestellung â im Anhang findest du deine Rechnung.

{artikel_confirm}
  Gesamt: {gesamt} â¬

{"Deine Zahlung ist eingegangen â alles erledigt! â" if paid else "Zahlungsdetails stehen auf der Rechnung."}

Liebe GrÃ¸Ãe,
Carola & das RMV Team ð
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
        print(f"  â Kunden-Email â {kunde_email}")
    except Exception as e:
        print(f"  â   Kunden-Email fehlgeschlagen: {e}")

    return True


# ââ Start âââââââââââââââââââââââââââââââââââââââââââââââââââââ

if __name__ == "__main__":
    missing = []
    if not SMTP_USER:     missing.append("EMAIL_ADRESSE")
    if not SMTP_PASSWORD: missing.append("EMAIL_PASSWORT")
    if not STRIPE_SECRET_KEY: missing.append("STRIPE_SECRET_KEY")

    print(f"""
ââââââââââââââââââââââââââââââââââââââââââââ
â   RMV Merch Shop â Cloud Server v2.1   â
ââââââââââââââââââââââââââââââââââââââââââââ
  Port:      {PORT}
  Email:     {SMTP_USER or 'â  NICHT GESETZT'}
  Stripe:    {'â konfiguriert' if STRIPE_SECRET_KEY else 'â  NICHT GESETZT'}
  Shop-URL:  {SHOP_URL or '(wird vom Browser Ã¼bergeben)'}
{''.join(chr(10)+'  â   Bitte setze: '+v+' (Render-Umgebungsvariablen)' for v in missing)}
  LÃ¤uft â Strg+C zum Stoppen
""")

    server = HTTPServer(("0.0.0.0", PORT), OrderHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer gestoppt.")
