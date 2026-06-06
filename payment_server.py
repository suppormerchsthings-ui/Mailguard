#!/usr/bin/env python3
"""
MailGuard — Serveur de paiement automatique
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Dès qu'un client paie, il reçoit automatiquement sa clé par email.

SETUP :
  1. pip install flask stripe
  2. Configurer les variables ci-dessous
  3. Créer 3 produits Stripe (un par plan) avec les metadata {"plan": "particulier"} etc.
  4. python payment_server.py
  5. Exposer avec ngrok en dev : ngrok http 5000
  6. En prod : déployer sur un VPS (ex: Hetzner 4€/mois) ou sur Railway/Render

STRIPE SETUP :
  - dashboard.stripe.com → Products → créer 3 produits :
      Particulier   29.99€  metadata: plan=particulier
      Professionnel 79.00€  metadata: plan=professionnel
      Entreprise    249.99€ metadata: plan=entreprise
  - Developers → Webhooks → Add endpoint → ton-domaine.com/webhook
  - Événement à écouter : checkout.session.completed
"""

from flask import Flask, request, jsonify, render_template_string
import hashlib, json, smtplib, os, stripe, logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from pathlib import Path

# ══ CONFIGURATION — À MODIFIER ══════════════════════════════
STRIPE_SECRET_KEY    = "sk_live_VOTRE_CLE_STRIPE"         # stripe.com/dashboard → Developers → API keys
STRIPE_WEBHOOK_SECRET = "whsec_VOTRE_SECRET_WEBHOOK"      # stripe.com/dashboard → Developers → Webhooks

# Email d'envoi (Gmail recommandé — créer un mot de passe d'application)
SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_USER     = "mailguard@votre-domaine.com"
SMTP_PASSWORD = "VOTRE_MOT_DE_PASSE_APPLICATION"          # Gmail → Sécurité → Mots de passe d'application
EMAIL_FROM    = "MailGuard <mailguard@votre-domaine.com>"

# Stripe Price IDs (créés dans le dashboard Stripe)
PRICE_IDS = {
    "particulier":   "price_XXXXXXXXXXXXX",   # à remplir après création dans Stripe
    "professionnel": "price_XXXXXXXXXXXXX",
    "entreprise":    "price_XXXXXXXXXXXXX",
}

LOG_FILE       = Path(__file__).parent / "transactions.json"
CUSTOMERS_FILE = Path(__file__).parent / "customers.json"
# ════════════════════════════════════════════════════════════

# ══ ALGO KEYGEN (identique à mailguard.py) ═══════════════════
_LK = bytes([77,71,95,115,101,99,114,101,116,95,50,48,50,52,
             95,120,57,107,50,112,55,113,49,95,109,103,50,52])

PLANS = {
    "particulier":   {"label":"Particulier",   "price":"29,99 €",  "max_acc":2},
    "professionnel": {"label":"Professionnel", "price":"79,00 €",  "max_acc":10},
    "entreprise":    {"label":"Entreprise",    "price":"249,99 €", "max_acc":0},
}

def _hx(v):
    return hashlib.sha256(f"{v}:{_LK.decode()}".encode()).hexdigest()

def gen_key(email: str, plan: str) -> str:
    r = _hx(f"{email.lower().strip()}:{plan}:KEY2024")[:32].upper()
    return f"{r[:8]}-{r[8:16]}-{r[16:24]}-{r[24:32]}"

# ══ ENVOI EMAIL ══════════════════════════════════════════════
EMAIL_TEMPLATE = """
<html><body style="font-family:Segoe UI,Arial;background:#f5f4f0;padding:30px">
<div style="max-width:520px;margin:0 auto;background:#fff;border-radius:8px;
            padding:32px;border:1px solid #dbd8d1">

  <div style="text-align:center;margin-bottom:24px">
    <h1 style="color:#2f7dc8;margin:0">✉ MailGuard</h1>
    <p style="color:#9c9890;margin:4px 0">Merci pour votre achat !</p>
  </div>

  <p style="color:#1c1a17">Bonjour,</p>
  <p style="color:#1c1a17">Votre paiement a bien été reçu. Voici votre clé de licence
  <strong>{plan_label}</strong> :</p>

  <div style="background:#f0f8ff;border:2px solid #2f7dc8;border-radius:6px;
              padding:16px;text-align:center;margin:20px 0">
    <p style="font-family:Consolas,monospace;font-size:20px;font-weight:bold;
              color:#2f7dc8;letter-spacing:2px;margin:0">{key}</p>
  </div>

  <p style="color:#1c1a17"><strong>Comment activer votre licence :</strong></p>
  <ol style="color:#5a5750">
    <li>Ouvrez MailGuard</li>
    <li>Cliquez sur <strong>💳 Activer Pro</strong> dans la barre du bas</li>
    <li>Sélectionnez l'offre <strong>{plan_label}</strong></li>
    <li>Collez votre clé et cliquez sur <strong>Activer</strong></li>
  </ol>

  <div style="border-top:1px solid #dbd8d1;margin-top:24px;padding-top:16px">
    <p style="color:#9c9890;font-size:12px;margin:0">
      Cette clé est liée à l'adresse email <strong>{email}</strong>.<br>
      Accès à vie · {features}<br><br>
      Un problème ? Contactez-nous : {support_email}
    </p>
  </div>
</div>
</body></html>
"""

def send_license_email(email: str, plan: str, key: str):
    p = PLANS[plan]
    max_acc = p["max_acc"]
    features = (f"{max_acc} compte(s) email · " if max_acc > 0 else "Comptes illimités · ") + \
               "Suppression illimitée · Mises à jour incluses"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🔑 Votre clé MailGuard {p['label']} — {p['price']}"
    msg["From"]    = EMAIL_FROM
    msg["To"]      = email

    html = EMAIL_TEMPLATE.format(
        plan_label=p["label"],
        key=key,
        email=email,
        features=features,
        support_email=SMTP_USER
    )
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SMTP_USER, email, msg.as_string())

    logging.info(f"Email envoyé → {email} ({plan})")

# ══ LOGGING TRANSACTIONS ═════════════════════════════════════
def log_transaction(email, plan, key, source="stripe", payment_id=""):
    data = []
    if LOG_FILE.exists():
        try: data = json.loads(LOG_FILE.read_text(encoding="utf-8"))
        except: pass
    data.append({
        "date":       datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "email":      email,
        "plan":       plan,
        "key":        key,
        "source":     source,
        "payment_id": payment_id,
    })
    LOG_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    # Aussi dans customers.json (compatible avec keygen_admin.py)
    cust = []
    if CUSTOMERS_FILE.exists():
        try: cust = json.loads(CUSTOMERS_FILE.read_text(encoding="utf-8"))
        except: pass
    cust = [c for c in cust if not (c["email"]==email and c["plan"]==plan)]
    cust.append({"email":email,"plan":plan,"key":key,
                 "date":datetime.now().strftime("%d/%m/%Y %H:%M"),"note":"auto"})
    CUSTOMERS_FILE.write_text(json.dumps(cust,indent=2,ensure_ascii=False),encoding="utf-8")

# ══ FLASK APP ════════════════════════════════════════════════
app = Flask(__name__)
stripe.api_key = STRIPE_SECRET_KEY
logging.basicConfig(level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s")

# ── Webhook Stripe ───────────────────────────────────────────
@app.route("/webhook/stripe", methods=["POST"])
def stripe_webhook():
    payload = request.get_data()
    sig     = request.headers.get("Stripe-Signature", "")
    try:
        event = stripe.Webhook.construct_event(payload, sig, STRIPE_WEBHOOK_SECRET)
    except stripe.error.SignatureVerificationError as e:
        logging.warning(f"Webhook signature invalide : {e}")
        return jsonify(error="Invalid signature"), 400
    except Exception as e:
        return jsonify(error=str(e)), 400

    if event["type"] == "checkout.session.completed":
        session  = event["data"]["object"]
        email    = (session.get("customer_details") or {}).get("email","")
        metadata = session.get("metadata") or {}
        plan     = metadata.get("plan","")

        if not email or plan not in PLANS:
            logging.error(f"Données manquantes : email={email} plan={plan}")
            return jsonify(ok=False), 400

        key = gen_key(email, plan)
        try:
            send_license_email(email, plan, key)
            log_transaction(email, plan, key, "stripe", session.get("id",""))
            logging.info(f"✓ Licence {plan} → {email} | clé: {key}")
        except Exception as e:
            logging.error(f"Erreur envoi email : {e}")
            # Ne pas retourner d'erreur à Stripe pour éviter les renvois
    return jsonify(ok=True), 200

# ── Webhook PayPal (IPN) ─────────────────────────────────────
@app.route("/webhook/paypal", methods=["POST"])
def paypal_webhook():
    import urllib.request, urllib.parse
    data = request.form.to_dict()

    # Vérification IPN PayPal
    verify_data = {"cmd": "_notify-validate", **data}
    verify_url  = "https://ipnpb.paypal.com/cgi-bin/webscr"
    req = urllib.request.Request(verify_url,
              urllib.parse.urlencode(verify_data).encode(), method="POST")
    with urllib.request.urlopen(req, timeout=10) as r:
        if r.read().decode() != "VERIFIED":
            logging.warning("PayPal IPN non vérifié")
            return "INVALID", 400

    if data.get("payment_status") != "Completed":
        return "OK", 200

    email = data.get("payer_email","")
    # Le plan est dans custom (à mettre dans le bouton PayPal : <input name="custom" value="particulier">)
    plan  = data.get("custom","")

    if not email or plan not in PLANS:
        logging.error(f"PayPal données manquantes : {email} / {plan}")
        return "ERROR", 400

    key = gen_key(email, plan)
    try:
        send_license_email(email, plan, key)
        log_transaction(email, plan, key, "paypal", data.get("txn_id",""))
        logging.info(f"✓ PayPal {plan} → {email}")
    except Exception as e:
        logging.error(f"Erreur : {e}")
    return "OK", 200

# ── Dashboard admin simple (protégé par mot de passe) ────────
ADMIN_PASSWORD = "changez-moi-en-prod"   # ← à changer

DASHBOARD_HTML = """
<!DOCTYPE html><html><head><title>MailGuard Admin</title>
<style>body{font-family:Segoe UI;background:#f5f4f0;padding:20px}
table{border-collapse:collapse;width:100%;background:#fff}
th,td{border:1px solid #dbd8d1;padding:8px 12px;text-align:left}
th{background:#2f7dc8;color:#fff}tr:hover{background:#e4f0fb}
h1{color:#2f7dc8}.badge{padding:3px 8px;border-radius:4px;font-size:12px;font-weight:bold}
.particulier{background:#e4f0fb;color:#2f7dc8}
.professionnel{background:#fef3e8;color:#7f5a00}
.entreprise{background:#e3f5ec;color:#2a8a56}
</style></head><body>
<h1>✉ MailGuard — Dashboard Admin</h1>
<p>{{ total }} clients enregistrés</p>
<table><tr><th>Date</th><th>Email</th><th>Offre</th><th>Clé</th><th>Source</th></tr>
{% for t in transactions %}
<tr><td>{{ t.date }}</td><td>{{ t.email }}</td>
<td><span class="badge {{ t.plan }}">{{ t.plan }}</span></td>
<td style="font-family:monospace">{{ t.key }}</td>
<td>{{ t.source }}</td></tr>
{% endfor %}
</table></body></html>
"""

@app.route("/admin")
def admin_dashboard():
    from flask import request as req, Response
    auth = req.authorization
    if not auth or auth.password != ADMIN_PASSWORD:
        return Response("Accès refusé", 401,
                        {"WWW-Authenticate": 'Basic realm="MailGuard Admin"'})
    data = []
    if LOG_FILE.exists():
        try: data = json.loads(LOG_FILE.read_text(encoding="utf-8"))
        except: pass
    return render_template_string(DASHBOARD_HTML,
                                  transactions=reversed(data), total=len(data))

# ── Liens de paiement Stripe (génère une session checkout) ───
@app.route("/pay/<plan>")
def create_checkout(plan):
    if plan not in PRICE_IDS:
        return "Plan inconnu", 404
    p = PLANS[plan]
    try:
        session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            line_items=[{"price": PRICE_IDS[plan], "quantity": 1}],
            mode="payment",
            metadata={"plan": plan},
            success_url=request.host_url + "success?plan=" + plan,
            cancel_url=request.host_url + "cancel",
        )
        from flask import redirect
        return redirect(session.url, 303)
    except Exception as e:
        return str(e), 500

@app.route("/success")
def success():
    plan = request.args.get("plan","")
    p = PLANS.get(plan, {})
    return f"""<html><body style="font-family:Segoe UI;text-align:center;padding:60px">
    <h1 style="color:#2a8a56">✓ Paiement confirmé !</h1>
    <p>Vous recevrez votre clé <strong>{p.get('label','')}</strong> par email dans quelques secondes.</p>
    <p style="color:#9c9890">Vérifiez vos spams si vous ne la recevez pas.</p>
    </body></html>"""

@app.route("/cancel")
def cancel():
    return """<html><body style="font-family:Segoe UI;text-align:center;padding:60px">
    <h1 style="color:#c0392b">Paiement annulé</h1>
    <p><a href="javascript:history.back()">← Retour</a></p>
    </body></html>"""

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))  # Railway injecte PORT automatiquement
    print("═" * 52)
    print("  MailGuard Payment Server")
    print(f"  Port : {port}")
    print("═" * 52)
    app.run(host="0.0.0.0", port=port, debug=False)
