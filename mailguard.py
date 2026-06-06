#!/usr/bin/env python3
"""
MailGuard v2.1 — Microsoft Outlook + Gmail
"""
import msal, requests, tkinter as tk, threading, time, json, hashlib, re as _re, io, base64, webbrowser, socket
from tkinter import ttk, scrolledtext, messagebox, simpledialog
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, date
from pathlib import Path

try:
    from PIL import Image as PILImage, ImageTk, ImageDraw
    PIL_OK = True
except ImportError:
    PIL_OK = False

try:
    import pystray
    TRAY_OK = True
except ImportError:
    TRAY_OK = False

try:
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    GMAIL_OK = True
except ImportError:
    GMAIL_OK = False

_S = requests.Session()

def _api_call(method: str, url: str, max_retries: int = 4, **kwargs):
    """
    Wrapper requests avec :
    - Timeout augmenté (30s)
    - Retry automatique sur 429 (Retry-After respecté)
    - Backoff exponentiel sur timeout/erreur réseau
    """
    kwargs.setdefault("timeout", 30)
    for attempt in range(max_retries):
        try:
            r = getattr(_S, method)(url, **kwargs)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 60))
                write(f"  ⏳ Limite API (429) — pause {wait}s puis reprise…", "warn")
                time.sleep(wait)
                continue
            if r.status_code in (500, 502, 503, 504):
                wait = 2 ** attempt
                write(f"  ⚠ Erreur serveur {r.status_code} — retry dans {wait}s…", "warn")
                time.sleep(wait)
                continue
            return r
        except requests.exceptions.Timeout:
            wait = 2 ** attempt
            if attempt < max_retries - 1:
                write(f"  ⏳ Timeout — retry {attempt+1}/{max_retries-1} dans {wait}s…", "warn")
                time.sleep(wait)
            else:
                write(f"  ✗ Timeout définitif après {max_retries} tentatives", "error")
                return None
        except requests.exceptions.ConnectionError:
            wait = 2 ** attempt
            if attempt < max_retries - 1:
                time.sleep(wait)
            else:
                return None
    return None

# ══ CONFIG MICROSOFT ══════════════════════════════════════════
CLIENT_ID = "c46a8492-4e3b-437e-86b4-929386666d3b"
AUTHORITY = "https://login.microsoftonline.com/consumers"
MS_SCOPES = ["https://graph.microsoft.com/Mail.ReadWrite",
             "https://graph.microsoft.com/User.Read"]

# ══ CONFIG GMAIL ══════════════════════════════════════════════
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

# Auto-updater (optionnel — fonctionne si updater.py est dans le même dossier)
try:
    from updater import check_for_update, CURRENT_VER as _APP_VER
    _UPDATER_OK = True
except ImportError:
    _UPDATER_OK = False

FREE_LIMIT   = 100
WEBSITE_URL  = "https://mailguard.fr"   # ← ton domaine ici

# ── Page de succès OAuth (remplace le "Authentication complete" moche) ──
_AUTH_SUCCESS_HTML = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>MailGuard — Connexion réussie</title>
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  body{background:#03070e;font-family:'Segoe UI',sans-serif;
       display:flex;align-items:center;justify-content:center;
       min-height:100vh;color:#e8f4ff}
  .card{text-align:center;padding:60px 48px;
        background:#060d1c;border:1px solid rgba(0,180,255,.2);
        border-radius:20px;max-width:480px;width:90%;
        box-shadow:0 0 60px rgba(0,180,255,.1)}
  .icon{font-size:64px;margin-bottom:20px;animation:pop .5s ease}
  @keyframes pop{from{transform:scale(0)}to{transform:scale(1)}}
  .logo{font-size:1.1rem;font-weight:800;color:#00b4ff;
        letter-spacing:1px;margin-bottom:24px}
  h1{font-size:1.6rem;font-weight:800;margin-bottom:12px;
     background:linear-gradient(135deg,#00b4ff,#7c6fff);
     -webkit-background-clip:text;-webkit-text-fill-color:transparent}
  p{color:#6b8099;font-size:.95rem;line-height:1.7;margin-bottom:8px}
  .note{font-size:.8rem;color:#3a5068;margin-top:20px;
        font-family:'Courier New',monospace}
  .bar{height:3px;background:rgba(0,180,255,.15);border-radius:2px;
       margin-top:28px;overflow:hidden}
  .fill{height:100%;background:linear-gradient(90deg,#00b4ff,#7c6fff);
        animation:load 3s linear forwards}
  @keyframes load{from{width:0}to{width:100%}}
</style>
</head>
<body>
<div class="card">
  <div class="logo">✉ MailGuard</div>
  <div class="icon">✅</div>
  <h1>Connexion réussie !</h1>
  <p>Merci d'avoir sélectionné <strong style="color:#00b4ff">MailGuard</strong>,<br>
     l'arme contre le spam et le scam.</p>
  <p>Vous pouvez fermer cet onglet<br>et retourner dans l'application.</p>
  <div class="bar"><div class="fill"></div></div>
  <p class="note">Fermeture automatique dans 3s…</p>
</div>
<script>setTimeout(()=>window.close(),3000)</script>
</body>
</html>"""

_AUTH_ERROR_HTML = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8"><title>MailGuard — Erreur</title>
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  body{background:#03070e;font-family:'Segoe UI',sans-serif;
       display:flex;align-items:center;justify-content:center;min-height:100vh;color:#e8f4ff}
  .card{text-align:center;padding:60px 48px;background:#060d1c;
        border:1px solid rgba(255,61,87,.2);border-radius:20px;
        max-width:480px;width:90%}
  .logo{color:#00b4ff;font-weight:800;margin-bottom:20px}
  h1{color:#ff3d57;font-size:1.4rem;margin-bottom:12px}
  p{color:#6b8099;font-size:.9rem}
</style>
</head>
<body>
<div class="card">
  <div class="logo">✉ MailGuard</div>
  <h1>❌ Erreur de connexion</h1>
  <p>Une erreur est survenue. Fermez cet onglet et réessayez depuis l'application.</p>
</div>
</body>
</html>"""

# ══ STRIPE — Liens de paiement (Stripe Dashboard → Payment Links) ══
# Créer 3 Payment Links sur stripe.com/dashboard → Products
# et coller les URLs ci-dessous. Rien d'autre à faire.
STRIPE_LINKS = {
    "particulier":   "https://buy.stripe.com/6oU8wPeMVaqz1vy6Tx6Vq00",
    "professionnel": "https://buy.stripe.com/7sY5kD5cl0PZa24cdR6Vq02",
    "entreprise":    "https://buy.stripe.com/4gMbJ16gp56fcac0v96Vq01",
}

# ══ LICENCE ═══════════════════════════════════════════════════
_LK = bytes([77,71,95,115,101,99,114,101,116,95,50,48,50,52,
             95,120,57,107,50,112,55,113,49,95,109,103,50,52])

def _hx(v):
    return hashlib.sha256(f"{v}:{_LK.decode()}".encode()).hexdigest()

# ── Offres disponibles ───────────────────────────────────────
PLANS = {
    "particulier":   {"label":"Particulier",   "price":"29,99 €",  "max_acc":2,  "color":"#2f7dc8"},
    "professionnel": {"label":"Professionnel", "price":"79,99 €",  "max_acc":10, "color":"#7f5a00"},
    "entreprise":    {"label":"Entreprise",    "price":"249,99 €", "max_acc":0,  "color":"#2a8a56"},
}
# max_acc = 0 → illimité

def gen_key(email, plan="particulier"):
    """[Admin] Génère une clé pour un email + plan donné."""
    r = _hx(f"{email.lower().strip()}:{plan}:KEY2024")[:32].upper()
    return f"{r[:8]}-{r[8:16]}-{r[16:24]}-{r[24:32]}"

def chk_key(email, key):
    """Vérifie la clé pour n'importe quel plan, retourne le plan ou None."""
    k = key.upper().replace("-","").strip()
    for plan in PLANS:
        if _hx(f"{email.lower().strip()}:{plan}:KEY2024")[:32].upper() == k:
            return plan
    return None

_PRO_REF = {
    # _hx("proprietaire@votre-domaine.com"): "entreprise",
    # _hx("client1@email.com"): "particulier",
}

def _is_ref_pro(email):
    return _hx(email.lower().strip()) in _PRO_REF

def _ref_plan(email):
    return _PRO_REF.get(_hx(email.lower().strip()), None)

# ══ STOCKAGE ══════════════════════════════════════════════════
_D      = Path.home() / ".mailguard"; _D.mkdir(exist_ok=True)
_FCAC   = _D / "auth.cache"
_FLIC   = _D / "license.json"
_FUSE   = _D / "usage.json"
_FCFG   = _D / "config.json"
_FGTOK  = _D / "gmail_token.json"
_FGCRD  = Path(__file__).parent / "google_credentials.json"

_LOGO_B64 = (
    "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"
)

_DCFG = {"workers":16,"max_mails":10000,"bg_interval":300,"threshold":60,"auto_empty":False}

def _lcfg():
    if _FCFG.exists():
        try: return {**_DCFG,**json.loads(_FCFG.read_text())}
        except: pass
    return dict(_DCFG)

def _scfg(c): _FCFG.write_text(json.dumps(c,indent=2))
CFG = _lcfg()

def _gu():
    if _FUSE.exists():
        try:
            d=json.loads(_FUSE.read_text())
            if d.get("d")==str(date.today()): return int(d.get("n",0))
        except: pass
    return 0

def _au(n): _FUSE.write_text(json.dumps({"d":str(date.today()),"n":_gu()+n}))

def _lico(email):
    if not _FLIC.exists(): return False
    try:
        d=json.loads(_FLIC.read_text())
        return d.get("e","").lower()==email.lower() and chk_key(email,d.get("k","")) is not None
    except: return False

def _get_plan(email):
    """Retourne le plan actif : 'free' | 'particulier' | 'professionnel' | 'entreprise'"""
    if _is_ref_pro(email): return _ref_plan(email) or "entreprise"
    if _FLIC.exists():
        try:
            d=json.loads(_FLIC.read_text())
            if d.get("e","").lower()==email.lower():
                return d.get("plan","particulier")
        except: pass
    return "free"

def _lics(e, k, plan):
    _FLIC.write_text(json.dumps({"e":e.lower(),"k":k,"plan":plan}))

def is_pro(email): return _is_ref_pro(email) or _lico(email)

def get_max_accounts(email=None):
    """Nombre max de comptes pour le plan actif (0 = illimité)."""
    e = email or user_email or ""
    plan = _get_plan(e)
    if plan == "free": return 1
    return PLANS.get(plan, {}).get("max_acc", 0)

# ══ MSAL CACHE ════════════════════════════════════════════════
try:
    if _FCAC.exists(): _FCAC.unlink()
except: pass
_tc = msal.SerializableTokenCache()

def _flush():
    if _tc.has_state_changed: _FCAC.write_text(_tc.serialize())

def _mapp(): return msal.PublicClientApplication(CLIENT_ID,authority=AUTHORITY,token_cache=_tc)
def get_accounts(): return _mapp().get_accounts()

# ══ ÉTAT GLOBAL ═══════════════════════════════════════════════
# Microsoft
ms_token     = None
ms_email     = None
# Gmail
gmail_email  = None
# Compat / actif
access_token = None
user_email   = None
pro_mode     = False
running      = False
bg_active    = True
tray_icon    = None
provider     = None   # "microsoft" | "gmail"
gmail_svc    = None   # service Google API (thread principal)
_gmail_local = threading.local()  # service par thread → exécution vraiment parallèle
deleted_ids      = []
kept_ids         = []
cat_tags         = {}
connected_emails = []   # comptes connectés dans la session

# ══ FILTRE INTELLIGENT ════════════════════════════════════════
_SPAM_KW = [
    "unsubscribe","se désabonner","désabonnement","newsletter",
    "promotion","promo","offre","soldes","deal","bon plan","réduction",
    "discount","% off","sale","marketing","no-reply","noreply",
    "do-not-reply","donotreply","publicité","advertisement",
    "limited time","flash sale","last chance","act now","offre exclusive","expire",
]
_SAFE_KW = [
    "verification","vérification","security alert","alerte sécurité",
    "password","mot de passe","two-factor","2fa","login attempt",
    "unauthorized","account suspended","paypal","stripe",
    "microsoft account","google account","facture urgent","legal",
]
_SPAM_DOMS = ["mailchimp","sendgrid","constantcontact","klaviyo","marketo","hubspot"]

# Mots-clés indiquant qu'un email pourrait être important malgré son score spam
_RISKY_KW = [
    "facture","invoice","commande","order","livraison","delivery",
    "paiement","payment","reçu","receipt","virement","transfer",
    "rendez-vous","appointment","rappel","reminder","réservation","booking",
    "contrat","contract","devis","quote","confirmation de commande",
    "votre compte","your account","action requise","action required",
    "important","urgent","échéance","deadline","résiliation","subscription cancel",
]

def _is_risky(subj: str, sender: str) -> bool:
    """True si l'email est à la fois spam ET potentiellement important."""
    s = (subj or "").lower(); f = (sender or "").lower()
    return any(k in s or k in f for k in _RISKY_KW)

def _score(subj, sender):
    s,f = (subj or "").lower(),(sender or "").lower()
    for k in _SAFE_KW:
        if k in s or k in f: return 0
    sc = 0
    for k in _SPAM_KW:
        if k in s: sc += 15
        if k in f: sc += 18
    if any(d in f for d in _SPAM_DOMS): sc += 30
    if subj and subj==subj.upper() and len(subj)>5: sc += 10
    if "@" not in sender: sc += 15
    return min(sc,100)

def _cat(subj, sender):
    s=(subj or "").lower()
    if any(k in s for k in ["urgent","facture","invoice","deadline","alerte","réunion"]): return "important"
    if any(k in s for k in ["projet","project","client","devis","rapport","contrat","meeting"]): return "pro"
    return "perso"

def passes_filter(subj, sender, rcv, fon, smart, kw, dom, age):
    if not fon: return True
    s,f=(subj or "").lower(),(sender or "").lower()
    matched=has=False
    if smart:
        has=True
        if _score(subj,sender)>=CFG.get("threshold",60): matched=True
    if kw:
        has=True
        # Vérifie dans le sujet ET dans le nom/adresse de l'expéditeur
        if any(k.strip().lower() in s or k.strip().lower() in f
               for k in kw.split(",") if k.strip()): matched=True
    if dom:
        has=True
        if any(d.strip().lower() in f for d in dom.split(",") if d.strip()): matched=True
    if not has or not matched: return False
    try:
        a=int(age or "0")
        if a>0 and rcv:
            dt=datetime.fromisoformat(rcv.replace("Z","+00:00"))
            if (datetime.now(timezone.utc)-dt).days < a: return False
    except: pass
    return True

# ══ GMAIL — AUTH ══════════════════════════════════════════════
def _gmail_login_worker():
    global access_token, user_email, pro_mode, gmail_svc, provider, gmail_email, ms_token, ms_email
    if not GMAIL_OK:
        write("Librairies Gmail manquantes. Installez : pip install google-auth-oauthlib google-api-python-client","error")
        return
    if not _FGCRD.exists():
        root.after(0, _show_gmail_setup)
        return
    write("Connexion Gmail en cours…","muted")
    try:
        creds = None
        if _FGTOK.exists():
            try: creds = Credentials.from_authorized_user_file(str(_FGTOK), GMAIL_SCOPES)
            except: creds = None
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(str(_FGCRD), GMAIL_SCOPES)
                creds = flow.run_local_server(
                port=0,
                success_message=_AUTH_SUCCESS_HTML,
                open_browser=True
            )
            _FGTOK.write_text(creds.to_json())
        gmail_svc = build("gmail","v1",credentials=creds,cache_discovery=False)
        _gmail_local.svc = gmail_svc   # cache pour le thread principal
        profile   = _get_gmail_svc().users().getProfile(userId="me").execute()
        user_email = profile.get("emailAddress","")
        pro_mode   = is_pro(user_email)
        provider   = "gmail"
        access_token = "gmail"   # sentinel non-null
        root.after(0, _on_login_ok)
        write(f"Connecté Gmail : {user_email} ✓","success")
        if pro_mode:
            write("  🔓 Licence Pro active","success")
            root.after(0,lambda p=_get_plan(user_email or ""):
                _show_welcome(f"Merci d'être passé à la version {PLANS.get(p,{}).get('label','Pro')} !",
                              "#2a8a56"))
        else:
            write(f"  Version gratuite — {_gu()}/{FREE_LIMIT} emails/jour","muted")
            root.after(0,lambda: _show_welcome(
                "Merci d'avoir sélectionné MailGuard,\nl'arme contre le spam et le scam !"))
        # Ouvrir la boîte Gmail après authentification
        root.after(1500, lambda: webbrowser.open("https://mail.google.com"))
    except Exception as e:
        write(f"Erreur Gmail : {e}","error")

def _show_welcome(msg: str, color: str = "#2f7dc8"):
    """Toast de bienvenue animé en bas à droite, fermeture auto 4s."""
    def _build():
        try:
            t = tk.Toplevel(root)
            t.overrideredirect(True)
            t.attributes("-topmost", True)
            t.configure(bg=color)
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
            t.geometry(f"380x82+{sw-400}+{sh-120}")
            tk.Label(t, text="✉  MailGuard", fg="white", bg=color,
                     font=("Segoe UI",8,"bold")).pack(anchor="w", padx=16, pady=(10,0))
            tk.Label(t, text=msg, fg="white", bg=color,
                     font=("Segoe UI",9), wraplength=340, justify="left"
                     ).pack(anchor="w", padx=16, pady=(2,10))
            t.after(4000, lambda: t.destroy() if t.winfo_exists() else None)
            t.bind("<Button-1>", lambda e: t.destroy())
        except Exception:
            pass
    root.after(300, _build)


def _get_gmail_svc():
    """Retourne un service Gmail propre au thread courant (thread-safe)."""
    svc = getattr(_gmail_local, "svc", None)
    if svc is None:
        try:
            creds = Credentials.from_authorized_user_file(str(_FGTOK), GMAIL_SCOPES)
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
                _FGTOK.write_text(creds.to_json())
            _gmail_local.svc = build("gmail","v1",credentials=creds,cache_discovery=False)
        except: return gmail_svc   # fallback sur le service principal
    return _gmail_local.svc

def _show_gmail_setup():
    msg = (
        "Fichier google_credentials.json introuvable.\n\n"
        "Pour connecter Gmail :\n"
        "1. Aller sur console.cloud.google.com\n"
        "2. Créer un projet → Activer Gmail API\n"
        "3. Identifiants → OAuth 2.0 → Application de bureau\n"
        "4. Télécharger le JSON → le renommer google_credentials.json\n"
        f"5. Le placer ici : {Path(__file__).parent}\n\n"
        "Puis cliquer à nouveau sur Se connecter avec Gmail."
    )
    messagebox.showinfo("Configuration Gmail", msg)

def _gmail_get_detail(msg_id):
    try:
        d = _get_gmail_svc().users().messages().get(
        userId="me", id=msg_id, format="metadata",
        metadataHeaders=["Subject","From","Date"]).execute()
        hdrs = {h["name"]:h["value"] for h in d.get("payload",{}).get("headers",[])}
        subj = hdrs.get("Subject","") or ""
        from_raw = hdrs.get("From","") or ""
        # On garde le nom complet "Instant Gaming <news@...>"
        # pour que le filtre puisse matcher le nom de l'expéditeur
        return {"id":msg_id,"subject":subj,"sender":from_raw}
    except: return {"id":msg_id,"subject":"","sender":""}

def _gmail_list_inbox(max_n=1000):
    svc = _get_gmail_svc()
    msgs = []
    req = svc.users().messages().list(
        userId="me", labelIds=["INBOX"], maxResults=min(max_n,500),
        fields="messages/id,nextPageToken")
    while req and len(msgs) < max_n:
        resp  = req.execute()
        batch = resp.get("messages",[])
        if not batch: break
        msgs.extend(batch)
        req = svc.users().messages().list_next(req, resp)
    return msgs   # [{id:...}, ...]

def _gmail_trash(msg_id):
    try:
        r=_get_gmail_svc().users().messages().trash(userId="me",id=msg_id).execute()
        return r.get("id",msg_id)
    except: return None

def _gmail_untrash(msg_id):
    try:
        _get_gmail_svc().users().messages().untrash(userId="me",id=msg_id).execute()
        return True
    except: return False

def _gmail_delete(msg_id):
    try:
        _get_gmail_svc().users().messages().delete(userId="me",id=msg_id).execute()
        return True
    except: return False

def _gmail_list_trash(max_n=50):
    try:
        resp=_get_gmail_svc().users().messages().list(
        userId="me",labelIds=["TRASH"],maxResults=max_n,
        fields="messages/id,nextPageToken").execute()
        return resp.get("messages",[])
    except: return []

# ══ WORKERS — SUPPRESSION ═════════════════════════════════════
def _show_risky_check(risky_items):
    """
    Affiche une fenêtre de contrôle pour les emails ambigus
    (score spam élevé MAIS contenant des mots-clés importants).
    Bloque le thread worker via threading.Event jusqu'à confirmation.
    Retourne l'ensemble des IDs que l'utilisateur veut vraiment supprimer.
    """
    confirmed_ids = set()
    evt = threading.Event()

    def _build():
        dlg = tk.Toplevel(root)
        dlg.title("⚠  Vérification avant suppression")
        dlg.geometry("600x480"); dlg.resizable(True, True)
        dlg.grab_set(); dlg.configure(bg=WHITE)
        dlg.lift()

        # En-tête
        hf = tk.Frame(dlg, bg="#fff3cd"); hf.pack(fill=tk.X)
        tk.Label(hf,
            text=f"⚠  {len(risky_items)} email(s) semblent importants malgré le filtre",
            fg="#7f5a00", bg="#fff3cd",
            font=("Segoe UI",10,"bold")).pack(padx=16, pady=10, anchor="w")
        tk.Label(hf,
            text="Décochez les emails que vous souhaitez CONSERVER avant de continuer.",
            fg="#7f5a00", bg="#fff3cd",
            font=("Segoe UI",8)).pack(padx=16, pady=(0,10), anchor="w")

        tk.Frame(dlg, bg="#f0c040", height=2).pack(fill=tk.X)

        # Liste avec cases à cocher (cochée = sera supprimé)
        list_f = tk.Frame(dlg, bg=WHITE); list_f.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        canvas  = tk.Canvas(list_f, bg=WHITE, highlightthickness=0)
        scrollb = ttk.Scrollbar(list_f, orient="vertical", command=canvas.yview)
        inner   = tk.Frame(canvas, bg=WHITE)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollb.set)
        scrollb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        check_vars = []
        for i, (mid, subj, sender, score) in enumerate(risky_items):
            row = tk.Frame(inner, bg=WHITE if i%2==0 else "#f8f8f8")
            row.pack(fill=tk.X, padx=4, pady=1)
            var = tk.BooleanVar(value=True)   # coché = supprimé par défaut
            check_vars.append((mid, var))
            tk.Checkbutton(row, variable=var, bg=row["bg"],
                           activebackground=row["bg"],
                           selectcolor=WHITE).pack(side=tk.LEFT, padx=(4,0))
            disp_subj   = subj[:52]+"…" if len(subj)>52 else subj
            disp_sender = sender[:36]+"…" if len(sender)>36 else sender
            tk.Label(row, text=disp_subj, fg="#1c1a17", bg=row["bg"],
                     font=("Segoe UI",9,"bold"), anchor="w", width=42).pack(side=tk.LEFT, padx=6)
            tk.Label(row, text=disp_sender, fg="#9c9890", bg=row["bg"],
                     font=("Segoe UI",8), anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text=f"score:{score}", fg="#c0392b", bg=row["bg"],
                     font=("Consolas",7)).pack(side=tk.RIGHT, padx=8)

        # Boutons
        tk.Frame(dlg, bg="#dbd8d1", height=1).pack(fill=tk.X)
        bf = tk.Frame(dlg, bg=WHITE); bf.pack(fill=tk.X, padx=12, pady=10)

        def _keep_all():
            for _, var in check_vars: var.set(False)

        def _delete_all():
            for _, var in check_vars: var.set(True)

        def _confirm():
            for mid, var in check_vars:
                if var.get(): confirmed_ids.add(mid)
            dlg.destroy(); evt.set()

        def _cancel():
            # Annuler = garder tous les emails risqués
            dlg.destroy(); evt.set()

        dlg.protocol("WM_DELETE_WINDOW", _cancel)

        tk.Button(bf, text="✓ Tout garder", command=_keep_all,
            bg="#eceae5", fg="#1c1a17", font=("Segoe UI",9), relief=tk.FLAT,
            padx=10, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=(0,6))
        tk.Button(bf, text="✗ Tout supprimer", command=_delete_all,
            bg="#fdecea", fg="#c0392b", font=("Segoe UI",9), relief=tk.FLAT,
            padx=10, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=(0,6))
        tk.Button(bf, text="⚡ Confirmer la sélection", command=_confirm,
            bg="#2f7dc8", fg="white", font=("Segoe UI",10,"bold"),
            relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
            activebackground="#235f9a").pack(side=tk.RIGHT)

    root.after(0, _build)
    evt.wait()          # bloque le thread worker jusqu'à la réponse
    return confirmed_ids


def _move_one_ms(mid, subj, hdr):
    disp=subj[:54]+"…" if len(subj)>54 else subj
    try:
        r=_S.post(f"https://graph.microsoft.com/v1.0/me/messages/{mid}/move",
                  headers=hdr,json={"destinationId":"deleteditems"},timeout=15)
        return r.status_code==201,(r.json().get("id",mid) if r.status_code==201 else None),disp
    except: return False,None,disp

def _move_one_gm(mid, subj):
    disp=subj[:54]+"…" if len(subj)>54 else subj
    nid=_gmail_trash(mid)
    return nid is not None, nid or mid, disp

def _clean_worker(fon, smart, kw, dom, age):
    global running
    if not access_token: return write("Non connecté","error")
    if not pro_mode:
        rem=FREE_LIMIT-_gu()
        if rem<=0:
            write(f"⚠ Limite gratuite atteinte ({FREE_LIMIT}/jour). Passez à Pro !","warn")
            root.after(0,lambda: _open_upgrade(root)); return
        max_run=min(rem,CFG["max_mails"])
    else: max_run=CFG["max_mails"]

    running=True; total=skipped=0
    write(""); write("Nettoyage démarré","header")
    if fon:
        write("  Filtres actifs :","muted")
        if smart: write("    ✓ Smart (newsletters/promos)","muted")
        if kw:    write(f"    ✓ Mots-clés : {kw}","muted")
        if dom:   write(f"    ✓ Domaine : {dom}","muted")
        if age and age!="0": write(f"    ✓ Âge ≥ {age}j","muted")
    else: write("  ⚠ Filtre désactivé — TOUS les mails supprimés","warn")
    write("")

    if provider=="gmail":
        _clean_worker_gmail(fon,smart,kw,dom,age,max_run)
    else:
        _clean_worker_ms(fon,smart,kw,dom,age,max_run)

    if CFG.get("auto_empty") and pro_mode:
        threading.Thread(target=_empty_trash_worker,daemon=True).start()
    running=False

def _clean_worker_ms(fon,smart,kw,dom,age,max_run):
    global running
    total=skipped=0
    auth={"Authorization":f"Bearer {access_token}"}
    full={**auth,"Content-Type":"application/json"}
    url=("https://graph.microsoft.com/v1.0/me/messages"
         "?$top=1000&$select=id,subject,from,receivedDateTime")
    while url and running and total<max_run:
        r=_S.get(url,headers=auth)
        if r.status_code!=200: write(f"Erreur API {r.status_code}","error"); break
        data=r.json(); emails=data.get("value",[]); url=data.get("@odata.nextLink")
        if not emails: break
        batch_del=[]; batch_keep=[]
        for m in emails:
            subj=(m.get("subject") or "").strip()
            fo=m.get("from") or {}
            addr=((fo.get("emailAddress") or {}).get("address") or "")
            rcv=m.get("receivedDateTime","")
            if passes_filter(subj,addr,rcv,fon,smart,kw,dom,age): batch_del.append(m)
            else: batch_keep.append((m.get("id"),subj,addr)); skipped+=1

        # ── Vérification interne : emails ambigus (spam + mots-clés importants)
        risky=[(m.get("id"),(m.get("subject") or "").strip(),
                ((m.get("from") or {}).get("emailAddress") or {}).get("address",""),
                _score((m.get("subject") or ""),
                       ((m.get("from") or {}).get("emailAddress") or {}).get("address","")))
               for m in batch_del
               if _is_risky((m.get("subject") or ""),
                            ((m.get("from") or {}).get("emailAddress") or {}).get("address",""))]
        if risky:
            write(f"  ⚠ {len(risky)} email(s) semblent importants — vérification…","warn")
            confirmed=_show_risky_check(risky)
            risky_ids={r[0] for r in risky}
            batch_del=[m for m in batch_del
                       if m.get("id") not in risky_ids or m.get("id") in confirmed]
            kept_back=len(risky_ids)-len(confirmed)
            if kept_back: write(f"  ✓ {kept_back} email(s) conservé(s) par précaution","success")
        for mid2,subj2,addr2 in batch_keep:
            cat=_cat(subj2,addr2); disp=subj2[:48]+"…" if len(subj2)>48 else subj2
            def _ak(s=disp,c=cat,i=mid2):
                if i not in kept_ids:
                    try:
                        kept_tree.insert(c,"end",iid=i,text=s,tags=(c,)); kept_ids.append(i)
                        n=sum(len(kept_tree.get_children(cid)) for cid in CAT_LABELS)
                        kept_count_lbl.config(text=f"— {n}")
                    except: pass
            root.after(0,_ak)
        limit=batch_del[:max_run-total]
        with ThreadPoolExecutor(max_workers=CFG["workers"]) as ex:
            futs={ex.submit(_move_one_ms,m.get("id"),(m.get("subject") or "").strip(),full):m for m in limit}
            for fut in as_completed(futs):
                if not running: break
                ok,nid,disp=fut.result()
                if ok:
                    total+=1
                    def _ad(s=disp,i=nid,t=total):
                        del_lb.insert(tk.END,s); deleted_ids.append(i)
                        del_cnt.config(text=f"— {del_lb.size()} mails")
                        progress_var.set(f"{t} supprimé(s)")
                    root.after(0,_ad)
                    if total%50==0: write(f"  {total} traités…","muted")
                else: write(f"  Erreur : {disp}","error")
    if not pro_mode: _au(total)
    write(""); write(f"✓  {total} supprimés  |  {skipped} conservés","success")
    if not pro_mode: write(f"   Quota : {_gu()}/{FREE_LIMIT} utilisés aujourd'hui","muted")
    root.after(0,lambda: progress_var.set(f"✓  {total} supprimés"))

def _clean_worker_gmail(fon,smart,kw,dom,age,max_run):
    global running
    total=skipped=0
    write("  Récupération des emails Gmail…","muted")
    raw_ids=_gmail_list_inbox(max_run*3)
    if not raw_ids: write("Boîte vide ou erreur","muted"); return
    write(f"  {len(raw_ids)} emails récupérés, analyse…","muted")
    batch_del=[]; batch_keep=[]
    with ThreadPoolExecutor(max_workers=20) as ex:
        futs={ex.submit(_gmail_get_detail,m["id"]):m for m in raw_ids[:max_run*2]}
        for fut in as_completed(futs):
            d=fut.result()
            if passes_filter(d["subject"],d["sender"],"",fon,smart,kw,dom,age):
                batch_del.append(d)
            else:
                batch_keep.append(d); skipped+=1

    # ── Vérification interne : emails ambigus
    risky = [(d["id"],d["subject"],d["sender"],_score(d["subject"],d["sender"]))
             for d in batch_del if _is_risky(d["subject"],d["sender"])]
    if risky:
        write(f"  ⚠ {len(risky)} email(s) semblent importants — vérification…","warn")
        confirmed = _show_risky_check(risky)
        risky_ids = {r[0] for r in risky}
        batch_del = [d for d in batch_del
                     if d["id"] not in risky_ids or d["id"] in confirmed]
        kept_back = len(risky_ids) - len(confirmed)
        if kept_back: write(f"  ✓ {kept_back} email(s) conservé(s) par précaution","success")
    for item in batch_keep:
        cat=_cat(item["subject"],item["sender"])
        disp=item["subject"][:48]+"…" if len(item["subject"])>48 else item["subject"]
        def _ak(s=disp,c=cat,i=item["id"]):
            if i not in kept_ids:
                try:
                    kept_tree.insert(c,"end",iid=i,text=s,tags=(c,)); kept_ids.append(i)
                    n=sum(len(kept_tree.get_children(cid)) for cid in CAT_LABELS)
                    kept_count_lbl.config(text=f"— {n}")
                except: pass
        root.after(0,_ak)
    limit=batch_del[:max_run-total]
    with ThreadPoolExecutor(max_workers=min(CFG["workers"],10)) as ex:
        futs={ex.submit(_move_one_gm,item["id"],item["subject"]):item for item in limit}
        for fut in as_completed(futs):
            if not running: break
            ok,nid,disp=fut.result()
            if ok:
                total+=1
                def _ad(s=disp,i=nid,t=total):
                    del_lb.insert(tk.END,s); deleted_ids.append(i)
                    del_cnt.config(text=f"— {del_lb.size()} mails")
                    progress_var.set(f"{t} supprimé(s)")
                root.after(0,_ad)
                if total%50==0: write(f"  {total} traités…","muted")
            else: write(f"  Erreur : {disp}","error")
    if not pro_mode: _au(total)
    write(""); write(f"✓  {total} supprimés  |  {skipped} conservés","success")
    if not pro_mode: write(f"   Quota : {_gu()}/{FREE_LIMIT} utilisés aujourd'hui","muted")
    root.after(0,lambda: progress_var.set(f"✓  {total} supprimés"))

def clean_mail():
    if not access_token: return write("Non connecté","error")
    fon=filter_var.get(); smart=smart_var.get()
    kw=kw_var.get(); dom=dom_var.get(); age=age_var.get()
    if not fon:
        if not messagebox.askyesno("⚠ Filtre désactivé",
            "TOUS vos emails seront supprim\xc3\xa9s.\nContinuer ?",icon="warning"): return
    threading.Thread(target=_clean_worker,args=(fon,smart,kw,dom,age),daemon=True).start()

def stop(): global running; running=False; write("Arrêt demandé…","warn")

# ══ WORKER — RESTAURATION ═════════════════════════════════════
def _restore_worker(items):
    done=fail=0; to_rm=[]
    if provider=="gmail":
        for idx,subj,mid in items:
            if _gmail_untrash(mid):
                to_rm.append(idx); done+=1; write(f"  Restauré  {subj}","success")
            else: fail+=1; write(f"  Échec     {subj}","error")
    else:
        hdr={"Authorization":f"Bearer {access_token}","Content-Type":"application/json"}
        for idx,subj,mid in items:
            r=_S.post(f"https://graph.microsoft.com/v1.0/me/messages/{mid}/move",
                      headers=hdr,json={"destinationId":"inbox"},timeout=15)
            if r.status_code==201:
                to_rm.append(idx); done+=1; write(f"  Restauré  {subj}","success")
            else: fail+=1; write(f"  Échec     {subj}","error")
    def _rm():
        for i in sorted(to_rm,reverse=True):
            del_lb.delete(i); deleted_ids.pop(i)
        del_cnt.config(text=f"— {del_lb.size()} mails")
        threading.Thread(target=_sync_trash,daemon=True).start()
        if done: write(f"{done} restauré(s) ✓","success")
        if fail: write(f"{fail} échec(s)","warn")
    root.after(0,_rm)

def restore_mail():
    if not access_token: return write("Non connecté","error")
    sel=del_lb.curselection()
    if not sel: return write("Sélectionnez des mails","warn")
    items=[(i,del_lb.get(i),deleted_ids[i]) for i in sel if i<len(deleted_ids)]
    threading.Thread(target=_restore_worker,args=(items,),daemon=True).start()

# ══ WORKER — VIDER CORBEILLE ══════════════════════════════════
def _empty_trash_worker():
    cnt=0; write("Vidage définitif de la corbeille…","warn")
    if provider=="gmail":
        msgs=_gmail_list_trash(500)
        for m in msgs:
            if _gmail_delete(m["id"]): cnt+=1
    else:
        hdr={"Authorization":f"Bearer {access_token}"}
        url=("https://graph.microsoft.com/v1.0/me/mailFolders/deleteditems/messages"
             "?$top=100&$select=id")
        while url:
            r=_S.get(url,headers=hdr,timeout=15)
            if r.status_code!=200: break
            data=r.json(); msgs=data.get("value",[]); url=data.get("@odata.nextLink")
            for m in msgs:
                d=_api_call("delete",f"https://graph.microsoft.com/v1.0/me/messages/{m['id']}",
                            headers=hdr)
                if d.status_code==204: cnt+=1
    def _clear():
        del_lb.delete(0,tk.END); deleted_ids.clear()
        del_cnt.config(text="— 0 mail")
    root.after(0,_clear)
    write(f"✓  {cnt} message(s) supprimé(s) définitivement","success")

def empty_trash():
    if not access_token: return write("Non connecté","error")
    if messagebox.askyesno("⚠ Suppression DÉFINITIVE", "Ces emails seront supprimés pour toujours.\nContinuer ?", icon="warning"):
        threading.Thread(target=_empty_trash_worker,daemon=True).start()

# ══ SYNCHRONISATION ═══════════════════════════════════════════
def _sync_trash():
    if not access_token: return
    if provider=="gmail":
        msgs=_gmail_list_trash(50)
        for m in msgs:
            mid=m.get("id")
            if mid not in deleted_ids:
                d=_gmail_get_detail(mid)
                disp=d["subject"][:54]+"…" if len(d["subject"])>54 else d["subject"]
                def _ins(s=disp or "(sans sujet)",i=mid):
                    del_lb.insert(0,s); deleted_ids.insert(0,i)
                    del_cnt.config(text=f"— {del_lb.size()} mails")
                root.after(0,_ins)
        root.after(0,lambda: del_cnt.config(text=f"— {del_lb.size()} mails"))
    else:
        hdr={"Authorization":f"Bearer {access_token}"}
        try:
            r=_S.get("https://graph.microsoft.com/v1.0/me/mailFolders/deleteditems",
                     headers=hdr,timeout=10)
            if r.status_code==200:
                tot=r.json().get("totalItemCount",0)
                root.after(0,lambda: del_cnt.config(text=f"— {tot} mail{'s' if tot!=1 else ''}"))
        except: pass
        try:
            url=("https://graph.microsoft.com/v1.0/me/mailFolders/deleteditems/messages"
                 "?$top=50&$select=id,subject")
            r=_api_call("get",url,headers=hdr)
            if r and r.status_code==200:
                for m in reversed(r.json().get("value",[])):
                    mid=m.get("id")
                    if mid not in deleted_ids:
                        subj=(m.get("subject") or "(sans sujet)").strip()
                        disp=subj[:54]+"…" if len(subj)>54 else subj
                        def _ins(s=disp,i=mid):
                            del_lb.insert(0,s); deleted_ids.insert(0,i)
                            del_cnt.config(text=f"— {del_lb.size()} mails")
                        root.after(0,_ins)
        except: pass

def _sync_inbox():
    if not access_token: return
    if provider=="gmail":
        raw=_gmail_list_inbox(50)
        with ThreadPoolExecutor(max_workers=10) as ex:
            details=list(ex.map(lambda m: _gmail_get_detail(m["id"]), raw))
        for d in details:
            mid=d["id"]
            cat=_cat(d["subject"],d["sender"])
            disp=d["subject"][:48]+"…" if len(d["subject"])>48 else d["subject"]
            if mid not in kept_ids:
                def _ins2(s=disp or "(sans sujet)",c=cat,i=mid):
                    if i not in kept_ids:
                        try:
                            kept_tree.insert(c,"end",iid=i,text=s,tags=(c,)); kept_ids.append(i)
                            n=sum(len(kept_tree.get_children(cid)) for cid in CAT_LABELS)
                            kept_count_lbl.config(text=f"— {n}")
                        except: pass
                root.after(0,_ins2)
    else:
        hdr={"Authorization":f"Bearer {access_token}"}
        try:
            url=("https://graph.microsoft.com/v1.0/me/messages"
                 "?$top=50&$select=id,subject,from")
            r=_api_call("get",url,headers=hdr)
            if r and r.status_code==200:
                for m in r.json().get("value",[]):
                    mid=m.get("id"); subj=(m.get("subject") or "(sans sujet)").strip()
                    fo=m.get("from") or {}
                    addr=((fo.get("emailAddress") or {}).get("address") or "")
                    cat=_cat(subj,addr); disp=subj[:48]+"…" if len(subj)>48 else subj
                    if mid not in kept_ids:
                        def _ins2(s=disp,c=cat,i=mid):
                            if i not in kept_ids:
                                try:
                                    kept_tree.insert(c,"end",iid=i,text=s,tags=(c,)); kept_ids.append(i)
                                    n=sum(len(kept_tree.get_children(cid)) for cid in CAT_LABELS)
                                    kept_count_lbl.config(text=f"— {n}")
                                except: pass
                        root.after(0,_ins2)
        except: pass

def _clear_kept():
    try:
        for cid in CAT_LABELS:
            for child in kept_tree.get_children(cid): kept_tree.delete(child)
        kept_ids.clear(); kept_count_lbl.config(text="")
    except: pass

def _full_sync():
    def _w():
        write("Synchronisation complète…","muted")
        root.after(0,lambda: del_lb.delete(0,tk.END))
        root.after(0,lambda: deleted_ids.clear())
        root.after(0,_clear_kept)
        _sync_trash(); _sync_inbox()
        write("Synchronisation terminée ✓","success")
    threading.Thread(target=_w,daemon=True).start()

# ══ CONNEXION / DÉCONNEXION ═══════════════════════════════════
def _refresh_header():
    """Met à jour l'en-tête et les boutons selon les providers connectés."""
    global pro_mode, ms_email, ms_token, gmail_email, gmail_svc
    plans_c = []
    if ms_email:    plans_c.append(_get_plan(ms_email))
    if gmail_email: plans_c.append(_get_plan(gmail_email))
    rank = {"free":0,"particulier":1,"professionnel":2,"entreprise":3}
    best = max(plans_c, key=lambda p: rank.get(p,0)) if plans_c else "free"
    pro_mode = best != "free"
    # Statut header
    parts = []
    if ms_email:    parts.append(f"📨 {ms_email[:20]}")
    if gmail_email: parts.append(f"📧 {gmail_email[:20]}")
    txt = "  |  ".join(parts) if parts else "Non connecté"
    conn_dot.config(fg=SUCCESS if parts else TEXT_M)
    conn_lbl.config(text=txt[:48], fg=SUCCESS if parts else TEXT_M)
    # Boutons login/logout dynamiques
    login_btn.config(
        text="📨 Déconnecter MS" if ms_email else "📨 Microsoft",
        command=logout_ms if ms_email else login_ms,
        bg=DANGER if ms_email else ACCENT,
        activebackground="#a93226" if ms_email else "#235f9a")
    gmail_btn.config(
        text="📧 Déconnecter Gmail" if gmail_email else "📧 Gmail",
        command=logout_gmail if gmail_email else login_gmail,
        bg=DANGER if gmail_email else GMAIL_RED,
        activebackground="#a93226" if gmail_email else "#c0392b")
    clean_btn.config(state=tk.NORMAL if (ms_email or gmail_email) else tk.DISABLED)
    # Barre comptes
    if ms_email or gmail_email:
        ab_disc.pack_forget()
        ab_conn.pack(fill=tk.X, padx=14, pady=7)
    else:
        ab_conn.pack_forget()
        ab_disc.pack(fill=tk.X, padx=14, pady=7)
    # Bouton Pro
    if pro_mode:
        upgrade_btn.config(text="⚙ Paramètres Pro",
            command=lambda: _open_pro(root), bg="#2a8a56", activebackground="#1e6b40")
    else:
        upgrade_btn.config(text="✨ Passer à Pro",
            command=lambda: _open_upgrade(root), bg=ORANGE, activebackground="#c0651a")


def _on_login_ok():
    if user_email and user_email not in connected_emails:
        connected_emails.append(user_email)
    _update_acct_menu()
    _refresh_header()
    def _startup_sync():
        _sync_trash(); _sync_inbox()
    threading.Thread(target=_startup_sync, daemon=True).start()

def _login_worker_ms(hint=None):
    global access_token,user_email,pro_mode,provider,ms_token,ms_email,gmail_email
    write("Ouverture du sélecteur de compte Microsoft…","muted")
    try:
        app=_mapp()
        kw={"prompt":"select_account","max_age":0,
            "success_template":_AUTH_SUCCESS_HTML,
            "error_template":_AUTH_ERROR_HTML}
        if hint: kw["login_hint"]=hint
        res=app.acquire_token_interactive(MS_SCOPES,**kw)
        if "access_token" in res:
            ms_token=res["access_token"]
            access_token=ms_token   # compat
            r=requests.get("https://graph.microsoft.com/v1.0/me?$select=userPrincipalName",
                           headers={"Authorization":f"Bearer {ms_token}"},timeout=15)
            ms_email=r.json().get("userPrincipalName","") if r.status_code==200 else (hint or "")
            user_email=ms_email; provider="microsoft" if not gmail_email else "both"; _flush()
            root.after(0,_on_login_ok)
            write(f"Connecté Microsoft : {user_email} ✓","success")
            if pro_mode:
                write("  🔓 Licence Pro active","success")
                root.after(0,lambda p=_get_plan(user_email or ""):
                    _show_welcome(f"Merci d'être passé à la version {PLANS.get(p,{}).get('label','Pro')} !",
                                  "#2a8a56"))
            else:
                write(f"  Version gratuite — {_gu()}/{FREE_LIMIT} emails/jour","muted")
                root.after(0,lambda: _show_welcome(
                    "Merci d'avoir sélectionné MailGuard,\nl'arme contre le spam et le scam !"))
            # Ouvrir la boîte mail Microsoft après authentification
            root.after(1500, lambda: webbrowser.open("https://outlook.live.com"))
        else: write("Échec : "+str(res.get("error_description","?")),"error")
    except Exception as e: write(f"Erreur : {e}","error")

def login_ms(hint=None):
    threading.Thread(target=_login_worker_ms,args=(hint,),daemon=True).start()

def login_gmail():
    threading.Thread(target=_gmail_login_worker,daemon=True).start()

def add_account():
    # Vérifier la limite de comptes selon le plan
    max_acc = get_max_accounts()
    if max_acc > 0 and len(connected_emails) >= max_acc:
        plan = _get_plan(user_email or "")
        plan_name = PLANS.get(plan,{}).get("label","actuel")
        messagebox.showwarning(
            "Limite atteinte",
            f"Votre plan {plan_name} autorise {max_acc} compte(s) maximum.\n\n"
            f"Passez à un plan supérieur pour connecter davantage de comptes.",
            parent=root)
        return
    hint=simpledialog.askstring("Ajouter un compte","Email (optionnel) :",parent=root)
    if hint is not None:
        if provider=="gmail" or (hint and "gmail" in hint.lower()):
            login_gmail()
        else: login_ms(hint.strip() or None)

def switch_account(username):
    if provider=="gmail": login_gmail()
    else: login_ms(hint=username)

def logout():
    global access_token,user_email,pro_mode,gmail_svc,provider
    try:
        if provider=="microsoft":
            app=_mapp()
            for acct in app.get_accounts():
                if acct.get("username","").lower()==(user_email or "").lower():
                    app.remove_account(acct)
            _flush()
        elif provider=="gmail" and _FGTOK.exists():
            _FGTOK.unlink()
    except: pass
    if user_email and user_email in connected_emails:
        connected_emails.remove(user_email)
    access_token=None; user_email=None; pro_mode=False; gmail_svc=None; provider=None
    _gmail_local.svc = None
    conn_dot.config(fg=TEXT_M); conn_lbl.config(text="Non connecté",fg=TEXT_M)
    login_btn.config(text="Se connecter",command=login_ms,bg=ACCENT,activebackground="#235f9a")
    clean_btn.config(state=tk.DISABLED)
    ab_conn.pack_forget(); ab_disc.pack(fill=tk.X,padx=14,pady=7)
    del_lb.delete(0,tk.END); deleted_ids.clear(); del_cnt.config(text="— 0 mail")
    _clear_kept()
    upgrade_btn.config(text="✨ Passer à Pro",
                       command=lambda: _open_upgrade(root),
                       bg=ORANGE, activebackground="#c0651a")
    write("Déconnecté ✓","warn")

def _update_acct_menu():
    if provider=="gmail":
        acct_combo["values"]=[user_email] if user_email else []
        if user_email: acct_var.set(user_email)
    else:
        accts=get_accounts(); names=[a.get("username","?") for a in accts]
        acct_combo["values"]=names
        if user_email in names: acct_var.set(user_email)
        elif names: acct_var.set(names[0])

# ══ BACKGROUND WORKER ═════════════════════════════════════════
def _bg_worker():
    processed=set()
    while True:
        time.sleep(CFG.get("bg_interval",300))
        if not bg_active or not access_token or running: continue
        if not pro_mode and _gu()>=FREE_LIMIT: continue
        try:
            moved=0
            if provider=="gmail":
                raw=_gmail_list_inbox(50)
                spam=[]
                for m in raw:
                    if m["id"] in processed: continue
                    d=_gmail_get_detail(m["id"])
                    if _score(d["subject"],d["sender"])>=CFG.get("threshold",60): spam.append(d)
                    processed.add(m["id"])
                allowed=CFG["max_mails"] if pro_mode else min(FREE_LIMIT-_gu(),len(spam))
                with ThreadPoolExecutor(max_workers=min(CFG["workers"],8)) as ex:
                    futs={ex.submit(_move_one_gm,d["id"],d["subject"]):d for d in spam[:max(0,allowed)]}
                    for fut in as_completed(futs):
                        ok,nid,disp=fut.result()
                        if ok:
                            moved+=1
                            if not pro_mode: _au(1)
                            def _u(s=disp,i=nid):
                                del_lb.insert(0,s); deleted_ids.insert(0,i)
                                del_cnt.config(text=f"— {del_lb.size()} mails")
                            root.after(0,_u)
            else:
                hdr={"Authorization":f"Bearer {access_token}"}
                full={**hdr,"Content-Type":"application/json"}
                url=("https://graph.microsoft.com/v1.0/me/messages"
                     "?$top=50&$select=id,subject,from,receivedDateTime")
                r=_S.get(url,headers=hdr,timeout=15)
                if r.status_code!=200: continue
                spam=[]
                for m in r.json().get("value",[]):
                    mid=m.get("id")
                    if mid in processed: continue
                    subj=(m.get("subject") or "").strip()
                    fo=m.get("from") or {}
                    addr=((fo.get("emailAddress") or {}).get("address") or "")
                    if _score(subj,addr)>=CFG.get("threshold",60): spam.append(m)
                    processed.add(mid)
                allowed=CFG["max_mails"] if pro_mode else min(FREE_LIMIT-_gu(),len(spam))
                with ThreadPoolExecutor(max_workers=min(CFG["workers"],8)) as ex:
                    futs={ex.submit(_move_one_ms,m.get("id"),(m.get("subject") or "").strip(),full):m
                          for m in spam[:max(0,allowed)]}
                    for fut in as_completed(futs):
                        ok,nid,disp=fut.result()
                        if ok:
                            moved+=1
                            if not pro_mode: _au(1)
                            def _u(s=disp,i=nid):
                                del_lb.insert(0,s); deleted_ids.insert(0,i)
                                del_cnt.config(text=f"— {del_lb.size()} mails")
                            root.after(0,_u)
            if moved:
                write(f"[Auto] {moved} spam(s) bloqué(s) ✓","success")
                if TRAY_OK and tray_icon:
                    try: tray_icon.notify(f"{moved} spam(s) bloqué(s)","MailGuard")
                    except: pass
        except: pass

# ══ ROOT + DESIGN ═════════════════════════════════════════════
root = tk.Tk()
root.title("MailGuard"); root.geometry("1140x740")
root.configure(bg="#f5f4f0"); root.minsize(920,580)

BG="#f5f4f0"; PANEL="#ffffff"; SIDEBAR="#eceae5"; BORDER="#dbd8d1"
ACCENT="#2f7dc8"; ACCENT_L="#e4f0fb"; SUCCESS="#2a8a56"
DANGER="#c0392b"; DANGER_L="#fdecea"; WARNING="#a85f00"
ORANGE="#e67e22"; TEXT="#1c1a17"; TEXT_M="#9c9890"; WHITE="#ffffff"
GMAIL_RED="#ea4335"
CAT_COLORS={"important":(WARNING,"#fef3e8"),"pro":(ACCENT,ACCENT_L),"perso":(SUCCESS,"#e3f5ec")}
CAT_LABELS={"important":"⭐ Importants","pro":"💼 Pro","perso":"👤 Perso"}
FS=("Segoe UI",9); FT=("Segoe UI",14,"bold"); FB=("Segoe UI",10,"bold"); FL=("Consolas",9)

def write(msg,tag=None):
    def _w():
        try:
            if tag: log_text.insert(tk.END,msg+"\n",tag)
            else:   log_text.insert(tk.END,msg+"\n")
            log_text.see(tk.END)
        except: pass
    root.after(0,_w)

# ══ UI — EN-TÊTE ══════════════════════════════════════════════
hdr_f=tk.Frame(root,bg=WHITE); hdr_f.pack(fill=tk.X)
_logo_ref=None
logo_f=tk.Frame(hdr_f,bg=WHITE); logo_f.pack(side=tk.LEFT,padx=8,pady=5)
if PIL_OK:
    try:
        _pil=PILImage.open(io.BytesIO(base64.b64decode(_LOGO_B64)))
        mh=62; ratio=mh/_pil.height
        _pil=_pil.resize((int(_pil.width*ratio),mh),PILImage.LANCZOS)
        _logo_ref=ImageTk.PhotoImage(_pil)
        tk.Label(logo_f,image=_logo_ref,bg=WHITE,bd=0).pack(side=tk.LEFT)
    except: tk.Label(logo_f,text="✉ MailGuard",fg=ACCENT,bg=WHITE,font=FT).pack(side=tk.LEFT)
else: tk.Label(logo_f,text="✉ MailGuard",fg=ACCENT,bg=WHITE,font=FT).pack(side=tk.LEFT)
st_f=tk.Frame(hdr_f,bg=WHITE); st_f.pack(side=tk.RIGHT,padx=14)
conn_dot=tk.Label(st_f,text="●",fg=TEXT_M,bg=WHITE,font=("Segoe UI",9))
conn_dot.pack(side=tk.LEFT,padx=(0,4))
conn_lbl=tk.Label(st_f,text="Non connecté",fg=TEXT_M,bg=WHITE,font=FS)
conn_lbl.pack(side=tk.LEFT)
tk.Frame(root,bg=BORDER,height=1).pack(fill=tk.X)

# ══ UI — BARRE COMPTES ════════════════════════════════════════
ab=tk.Frame(root,bg=SIDEBAR); ab.pack(fill=tk.X)
ab_disc=tk.Frame(ab,bg=SIDEBAR); ab_disc.pack(fill=tk.X,padx=14,pady=7)
tk.Label(ab_disc,text="👆  Cliquez sur Se connecter pour choisir votre compte",
         fg=TEXT_M,bg=SIDEBAR,font=FS).pack(side=tk.LEFT)
ab_conn=tk.Frame(ab,bg=SIDEBAR)
tk.Label(ab_conn,text="Compte :",fg=TEXT,bg=SIDEBAR,font=("Segoe UI",9,"bold")
         ).pack(side=tk.LEFT,padx=(0,6))
acct_var=tk.StringVar()
acct_combo=ttk.Combobox(ab_conn,textvariable=acct_var,width=32,state="readonly",font=FS)
acct_combo.pack(side=tk.LEFT,padx=(0,8))
acct_combo.bind("<<ComboboxSelected>>",lambda e: switch_account(acct_var.get()))
tk.Button(ab_conn,text="+ Ajouter",command=add_account,bg=ACCENT,fg=WHITE,font=FS,
          relief=tk.FLAT,padx=8,pady=3,cursor="hand2",activebackground="#235f9a",
          activeforeground=WHITE).pack(side=tk.LEFT,padx=(0,6))
tk.Button(ab_conn,text="↺ Sync",command=_full_sync,bg=SIDEBAR,fg=TEXT_M,font=FS,
          relief=tk.FLAT,padx=8,pady=3,cursor="hand2",activebackground=ACCENT_L,
          activeforeground=ACCENT,borderwidth=1,highlightthickness=1,
          highlightbackground=BORDER).pack(side=tk.LEFT)
tk.Frame(root,bg=BORDER,height=1).pack(fill=tk.X)

# ══ UI — FILTRES ══════════════════════════════════════════════
filter_var=tk.BooleanVar(value=True); smart_var=tk.BooleanVar(value=True)
kw_var=tk.StringVar(value=""); dom_var=tk.StringVar(value=""); age_var=tk.StringVar(value="0")
fb=tk.Frame(root,bg=ACCENT_L); fb.pack(fill=tk.X)
fb_in=tk.Frame(fb,bg=ACCENT_L); fb_in.pack(fill=tk.X,padx=10,pady=6)
tk.Label(fb_in,text="🔍",bg=ACCENT_L,font=("Segoe UI",11)).pack(side=tk.LEFT)
tk.Label(fb_in,text="Filtre :",fg=TEXT,bg=ACCENT_L,font=FB).pack(side=tk.LEFT,padx=(3,6))
tk.Checkbutton(fb_in,text="Actif",variable=filter_var,bg=ACCENT_L,fg=TEXT,font=FS,
               activebackground=ACCENT_L,selectcolor=WHITE).pack(side=tk.LEFT,padx=(0,10))
tk.Frame(fb_in,bg=BORDER,width=1,height=18).pack(side=tk.LEFT,padx=5)
tk.Checkbutton(fb_in,text="Smart (newsletters/promos)",variable=smart_var,bg=ACCENT_L,fg=TEXT,
               font=FS,activebackground=ACCENT_L,selectcolor=WHITE).pack(side=tk.LEFT,padx=(0,10))
tk.Frame(fb_in,bg=BORDER,width=1,height=18).pack(side=tk.LEFT,padx=5)
tk.Label(fb_in,text="Sujet :",fg=TEXT_M,bg=ACCENT_L,font=FS).pack(side=tk.LEFT,padx=(0,3))
tk.Entry(fb_in,textvariable=kw_var,width=14,font=FS,relief=tk.FLAT,bg=WHITE,
         highlightthickness=1,highlightbackground=BORDER).pack(side=tk.LEFT,padx=(0,8),ipady=2)
tk.Label(fb_in,text="Domaine :",fg=TEXT_M,bg=ACCENT_L,font=FS).pack(side=tk.LEFT,padx=(0,3))
tk.Entry(fb_in,textvariable=dom_var,width=12,font=FS,relief=tk.FLAT,bg=WHITE,
         highlightthickness=1,highlightbackground=BORDER).pack(side=tk.LEFT,padx=(0,8),ipady=2)
tk.Label(fb_in,text="Âge ≥",fg=TEXT_M,bg=ACCENT_L,font=FS).pack(side=tk.LEFT,padx=(0,3))
tk.Entry(fb_in,textvariable=age_var,width=4,font=FS,relief=tk.FLAT,bg=WHITE,
         highlightthickness=1,highlightbackground=BORDER).pack(side=tk.LEFT,padx=(0,2),ipady=2)
tk.Label(fb_in,text="j",fg=TEXT_M,bg=ACCENT_L,font=FS).pack(side=tk.LEFT)
tk.Frame(root,bg=BORDER,height=1).pack(fill=tk.X)

# ══ UI — CONTENU PRINCIPAL ════════════════════════════════════
main_f=tk.Frame(root,bg=BG); main_f.pack(fill=tk.BOTH,expand=True)
lp=tk.Frame(main_f,bg=SIDEBAR,width=270); lp.pack(side=tk.LEFT,fill=tk.Y); lp.pack_propagate(False)
lp_h=tk.Frame(lp,bg=SIDEBAR); lp_h.pack(fill=tk.X,padx=10,pady=(10,4))
tk.Label(lp_h,text="📥 Conservés",fg=TEXT,bg=SIDEBAR,font=FB).pack(side=tk.LEFT)
kept_count_lbl=tk.Label(lp_h,text="",fg=TEXT_M,bg=SIDEBAR,font=FS); kept_count_lbl.pack(side=tk.LEFT,padx=4)
tk.Frame(lp,bg=BORDER,height=1).pack(fill=tk.X)
tf=tk.Frame(lp,bg=PANEL); tf.pack(fill=tk.BOTH,expand=True,padx=5,pady=5)
t_vsb=ttk.Scrollbar(tf); t_vsb.pack(side=tk.RIGHT,fill=tk.Y)
kept_tree=ttk.Treeview(tf,yscrollcommand=t_vsb.set,selectmode="extended",show="tree")
kept_tree.pack(side=tk.LEFT,fill=tk.BOTH,expand=True); t_vsb.config(command=kept_tree.yview)
kept_tree.column("#0",width=238,stretch=True)
kept_tree.tag_configure("important",foreground=WARNING)
kept_tree.tag_configure("pro",      foreground=ACCENT)
kept_tree.tag_configure("perso",    foreground=SUCCESS)
kept_tree.tag_configure("cat",      font=("Segoe UI",9,"bold"))
cat_tags={"important":"important","pro":"pro","perso":"perso"}
for cid,clabel in CAT_LABELS.items():
    kept_tree.insert("","end",iid=cid,text=clabel,open=True,tags=("cat",))
tk.Frame(lp,bg=BORDER,height=1).pack(fill=tk.X)
tk.Button(lp,text="↺ Actualiser",command=lambda: threading.Thread(target=_sync_inbox,daemon=True).start(),
    bg=SIDEBAR,fg=ACCENT,font=FB,relief=tk.FLAT,pady=7,cursor="hand2",
    borderwidth=0,activebackground=ACCENT_L,activeforeground=ACCENT).pack(fill=tk.X,padx=5,pady=5)
tk.Frame(main_f,bg=BORDER,width=1).pack(side=tk.LEFT,fill=tk.Y)

cp=tk.Frame(main_f,bg=SIDEBAR,width=255); cp.pack(side=tk.LEFT,fill=tk.Y); cp.pack_propagate(False)
cp_h=tk.Frame(cp,bg=SIDEBAR); cp_h.pack(fill=tk.X,padx=10,pady=(10,4))
tk.Label(cp_h,text="🗑 Corbeille",fg=TEXT,bg=SIDEBAR,font=FB).pack(side=tk.LEFT)
del_cnt=tk.Label(cp_h,text="— 0 mail",fg=TEXT_M,bg=SIDEBAR,font=FS); del_cnt.pack(side=tk.LEFT,padx=4)
tk.Label(cp,text="Ctrl+clic pour multi-sélection",fg=TEXT_M,bg=SIDEBAR,font=("Segoe UI",7)).pack(padx=10,anchor="w")
tk.Frame(cp,bg=BORDER,height=1).pack(fill=tk.X)
dw=tk.Frame(cp,bg=PANEL,highlightthickness=1,highlightbackground=BORDER)
dw.pack(fill=tk.BOTH,expand=True,padx=5,pady=5)
d_vsb=ttk.Scrollbar(dw); d_vsb.pack(side=tk.RIGHT,fill=tk.Y)
del_lb=tk.Listbox(dw,selectmode=tk.EXTENDED,bg=PANEL,fg=TEXT,selectbackground=DANGER_L,
    selectforeground=DANGER,font=FS,borderwidth=0,highlightthickness=0,
    activestyle="none",yscrollcommand=d_vsb.set,relief=tk.FLAT,cursor="hand2")
del_lb.pack(side=tk.LEFT,fill=tk.BOTH,expand=True); d_vsb.config(command=del_lb.yview)
tk.Frame(cp,bg=BORDER,height=1).pack(fill=tk.X)
restore_btn=tk.Button(cp,text="↩ Restaurer la sélection",command=restore_mail,
    bg=SIDEBAR,fg=TEXT_M,font=FB,relief=tk.FLAT,pady=7,cursor="hand2",
    borderwidth=0,activebackground=ACCENT_L,activeforeground=ACCENT,
    state=tk.DISABLED,disabledforeground=TEXT_M)
restore_btn.pack(fill=tk.X,padx=5,pady=(4,2))
def _on_del_sel(e=None):
    n=len(del_lb.curselection())
    restore_btn.config(state=tk.NORMAL if n else tk.DISABLED,fg=ACCENT if n else TEXT_M)
del_lb.bind("<<ListboxSelect>>",_on_del_sel)
tk.Button(cp,text="🗑 Vider définitivement",command=empty_trash,
    bg=DANGER_L,fg=DANGER,font=FB,relief=tk.FLAT,pady=7,cursor="hand2",
    borderwidth=0,activebackground="#f5c6c2",activeforeground=DANGER).pack(fill=tk.X,padx=5,pady=(2,5))
tk.Frame(main_f,bg=BORDER,width=1).pack(side=tk.LEFT,fill=tk.Y)

rp=tk.Frame(main_f,bg=BG); rp.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
rp_h=tk.Frame(rp,bg=BG); rp_h.pack(fill=tk.X,padx=12,pady=(12,6))
tk.Label(rp_h,text="📋 Activité",fg=TEXT,bg=BG,font=FB).pack(side=tk.LEFT)
tk.Button(rp_h,text="Effacer",command=lambda: log_text.delete("1.0",tk.END),
    bg=BG,fg=TEXT_M,font=FS,relief=tk.FLAT,cursor="hand2",
    borderwidth=0,activebackground=BG,activeforeground=TEXT).pack(side=tk.RIGHT)
tk.Frame(rp,bg=BORDER,height=1).pack(fill=tk.X)
log_text=scrolledtext.ScrolledText(rp,bg=PANEL,fg=TEXT,insertbackground=ACCENT,
    font=FL,borderwidth=0,highlightthickness=0,relief=tk.FLAT,padx=12,pady=10,wrap=tk.WORD)
log_text.pack(fill=tk.BOTH,expand=True,padx=4,pady=4)
log_text.tag_configure("success",foreground=SUCCESS)
log_text.tag_configure("error",  foreground=DANGER)
log_text.tag_configure("warn",   foreground=WARNING)
log_text.tag_configure("header", foreground=ACCENT,font=("Segoe UI",10,"bold"))
log_text.tag_configure("muted",  foreground=TEXT_M)

# ══ UI — BARRE DU BAS ════════════════════════════════════════
tk.Frame(root,bg=BORDER,height=1).pack(fill=tk.X)
bot=tk.Frame(root,bg=WHITE,height=62); bot.pack(fill=tk.X); bot.pack_propagate(False)
br=tk.Frame(bot,bg=WHITE); br.pack(side=tk.LEFT,padx=12,pady=10)

def _fb(text,cmd,bg_c,hov,state=tk.NORMAL):
    b=tk.Button(br,text=text,command=cmd,bg=bg_c,fg=WHITE,font=FB,
                relief=tk.FLAT,padx=13,pady=6,cursor="hand2",borderwidth=0,
                activebackground=hov,activeforeground=WHITE,state=state,
                disabledforeground="#aaaaaa")
    b.pack(side=tk.LEFT,padx=(0,5)); return b

def _ob(text,cmd,col):
    b=tk.Button(br,text=text,command=cmd,bg=WHITE,fg=col,font=FB,
                relief=tk.FLAT,padx=11,pady=5,cursor="hand2",
                borderwidth=1,highlightthickness=1,highlightbackground=col,
                activebackground=DANGER_L,activeforeground=col)
    b.pack(side=tk.LEFT,padx=(0,5)); return b

login_btn   = _fb("📨 Microsoft",  login_ms,                     ACCENT,     "#235f9a")
gmail_btn   = _fb("📧 Gmail",      login_gmail,                  GMAIL_RED,  "#c0392b")
clean_btn   = _fb("⚡ Nettoyer",   clean_mail,                   SUCCESS,    "#1e6b40", tk.DISABLED)
_ob("⏹ Stop",                      stop,                         DANGER)
# Un seul bouton Pro — change selon le statut (passer à pro / paramètres pro)
upgrade_btn = _fb("✨ Passer à Pro", lambda: _open_upgrade(root), ORANGE,    "#c0651a")

progress_var=tk.StringVar(value="")
tk.Label(bot,textvariable=progress_var,fg=TEXT_M,bg=WHITE,font=FS).pack(side=tk.RIGHT,padx=14)

# ══ DIALOGUES ═════════════════════════════════════════════════
def _open_pro(parent):
    if not pro_mode: _open_upgrade(parent); return
    dlg=tk.Toplevel(parent); dlg.title("⚙ Paramètres Pro")
    dlg.geometry("420x360"); dlg.resizable(False,False)
    dlg.grab_set(); dlg.configure(bg=WHITE)
    tk.Label(dlg,text="⚙ Paramètres Pro",fg=ACCENT,bg=WHITE,font=FB).pack(pady=(16,8))
    frm=tk.Frame(dlg,bg=WHITE); frm.pack(padx=20,fill=tk.X)
    def _row(lbl,val,lo,hi):
        f=tk.Frame(frm,bg=WHITE); f.pack(fill=tk.X,pady=4)
        tk.Label(f,text=lbl,fg=TEXT,bg=WHITE,font=FS,width=32,anchor="w").pack(side=tk.LEFT)
        v=tk.IntVar(value=val)
        tk.Spinbox(f,from_=lo,to=hi,textvariable=v,width=8,font=FS).pack(side=tk.LEFT)
        return v
    v_w=_row("Connexions parallèles (workers) :",CFG["workers"],1,100)
    v_m=_row("Max emails par nettoyage :",CFG["max_mails"],100,100000)
    v_i=_row("Intervalle scan auto (secondes) :",CFG["bg_interval"],60,3600)
    v_t=_row("Seuil de spam (0–100) :",CFG["threshold"],10,100)
    ae_v=tk.BooleanVar(value=CFG.get("auto_empty",False))
    ae_f=tk.Frame(frm,bg=WHITE); ae_f.pack(fill=tk.X,pady=4)
    tk.Checkbutton(ae_f,text="Vider corbeille automatiquement après nettoyage",
                   variable=ae_v,bg=WHITE,font=FS,activebackground=WHITE,
                   selectcolor=PANEL).pack(anchor="w")
    def _save():
        CFG["workers"]=v_w.get(); CFG["max_mails"]=v_m.get()
        CFG["bg_interval"]=v_i.get(); CFG["threshold"]=v_t.get()
        CFG["auto_empty"]=ae_v.get(); _scfg(CFG)
        write("Paramètres Pro sauvegardés ✓","success"); dlg.destroy()
    tk.Button(dlg,text="Sauvegarder",command=_save,bg=SUCCESS,fg=WHITE,font=FB,
              relief=tk.FLAT,padx=16,pady=7,cursor="hand2",borderwidth=0,
              activebackground="#1e6b40").pack(pady=(12,4))
    tk.Button(dlg,text="Annuler",command=dlg.destroy,bg=SIDEBAR,fg=TEXT,font=FS,
              relief=tk.FLAT,padx=16,pady=5,cursor="hand2",borderwidth=0).pack()

def _open_upgrade(parent):
    """
    Ouvre le site de paiement dans le navigateur.
    Après achat, le client reçoit sa clé par email
    et l'entre dans le champ ci-dessous.
    """
    dlg = tk.Toplevel(parent)
    dlg.title("💳 MailGuard Pro")
    dlg.geometry("480x420"); dlg.resizable(False,False)
    dlg.grab_set(); dlg.configure(bg=WHITE)
    dlg.lift()

    # ── En-tête ──────────────────────────────────────────────
    hdr_f = tk.Frame(dlg, bg=ACCENT); hdr_f.pack(fill=tk.X)
    tk.Label(hdr_f, text="✉  MailGuard Pro",
             fg=WHITE, bg=ACCENT, font=("Segoe UI",13,"bold")
             ).pack(side=tk.LEFT, padx=18, pady=14)
    tk.Label(hdr_f, text="🔒 SSL · Stripe",
             fg="#c8e0f5", bg=ACCENT, font=("Segoe UI",9)
             ).pack(side=tk.RIGHT, padx=18)

    # ── Étapes ───────────────────────────────────────────────
    steps_f = tk.Frame(dlg, bg="#e8f4fd"); steps_f.pack(fill=tk.X)
    for step in ["1️⃣  Choisir une offre sur le site",
                 "2️⃣  Payer en ligne (carte, sans compte)",
                 "3️⃣  Recevoir la clé par email",
                 "4️⃣  L'entrer ci-dessous"]:
        tk.Label(steps_f, text=step, fg=ACCENT, bg="#e8f4fd",
                 font=("Segoe UI",8)).pack(anchor="w", padx=18, pady=3)
    tk.Frame(steps_f, bg="#e8f4fd", height=6).pack()

    tk.Frame(dlg, bg=BORDER, height=1).pack(fill=tk.X)

    # ── Bouton site de paiement ───────────────────────────────
    mid = tk.Frame(dlg, bg=WHITE); mid.pack(padx=24, pady=20, fill=tk.X)

    def _open_site():
        # Ouvre index.html local en priorité, site web en fallback
        local = Path(__file__).parent / "html" / "index.html"
        if local.exists():
            target = local.as_uri() + "#pricing"
        else:
            target = f"{WEBSITE_URL}/#pricing"
            write("index.html introuvable → redirection vers le site","warn")
        webbrowser.open(target)
        site_btn.config(text="✓  Page de paiement ouverte",
                        bg=SUCCESS, state=tk.DISABLED)
        key_entry.config(state=tk.NORMAL, highlightbackground=SUCCESS)
        key_entry.focus()
        info_lbl.config(text="Après paiement, copiez la clé reçue par email ↓",
                        fg=SUCCESS)

    site_btn = tk.Button(mid,
        text=f"🌐  Accéder à la page de paiement",
        command=_open_site,
        bg=ACCENT, fg=WHITE, font=("Segoe UI",11,"bold"),
        relief=tk.FLAT, pady=13, cursor="hand2",
        activebackground="#235f9a", activeforeground=WHITE)
    site_btn.pack(fill=tk.X, pady=(0,12))

    info_lbl = tk.Label(mid,
        text="Cliquez ci-dessus → choisissez votre offre → payez par carte",
        fg=TEXT_M, bg=WHITE, font=("Segoe UI",8,"italic"))
    info_lbl.pack(anchor="w", pady=(0,16))

    tk.Frame(mid, bg=BORDER, height=1).pack(fill=tk.X, pady=(0,16))

    # ── Activation clé ────────────────────────────────────────
    tk.Label(mid, text="Clé de licence reçue par email :",
             fg=TEXT, bg=WHITE, font=("Segoe UI",9,"bold")).pack(anchor="w")

    kv = tk.StringVar()
    key_entry = tk.Entry(mid, textvariable=kv,
                         font=("Consolas",11), relief=tk.FLAT,
                         highlightthickness=2, highlightbackground=BORDER,
                         state=tk.DISABLED, disabledbackground="#f8f8f8")
    key_entry.pack(fill=tk.X, pady=(6,10), ipady=7)

    def _activate():
        global pro_mode
        if not user_email:
            messagebox.showwarning("", "Connectez-vous d'abord.", parent=dlg)
            return
        k = kv.get().strip()
        if not k:
            messagebox.showwarning("", "Entrez votre clé de licence.", parent=dlg)
            return
        matched = chk_key(user_email, k)
        if matched:
            _lics(user_email, k, matched)
            pro_mode = True
            p = PLANS[matched]
            max_txt = f"{p['max_acc']} compte(s)" if p["max_acc"]>0 else "illimité"
            messagebox.showinfo("🎉 Activé !",
                f"Bienvenue dans MailGuard {p['label']} !\n\n"
                f"✓ Comptes email : {max_txt}\n"
                f"✓ Suppression illimitée\n✓ Accès à vie", parent=dlg)
            write(f"🔓 Licence {p['label']} activée ✓","success")
            dlg.destroy()
        else:
            messagebox.showerror("Clé invalide",
                f"Cette clé ne correspond pas à {user_email}.\n\n"
                "Vérifiez que vous utilisez la même adresse email\n"
                "qu'au moment du paiement.", parent=dlg)

    act_btn = tk.Button(mid, text="🔓  Activer la licence",
        command=_activate,
        bg=SUCCESS, fg=WHITE, font=("Segoe UI",10,"bold"),
        relief=tk.FLAT, pady=9, cursor="hand2",
        activebackground="#1e6b40", activeforeground=WHITE)
    act_btn.pack(fill=tk.X)

    tk.Label(dlg,
        text="🔒 Paiement chiffré · Aucune donnée bancaire stockée · Remboursement 14j",
        fg=TEXT_M, bg=WHITE, font=("Segoe UI",8)).pack(pady=6)



# ══ SYSTEM TRAY ═══════════════════════════════════════════════
def _build_tray():
    global tray_icon
    if not TRAY_OK: return
    try:
        if PIL_OK:
            img=PILImage.open(io.BytesIO(base64.b64decode(_LOGO_B64))).resize((64,64),PILImage.LANCZOS).convert("RGBA")
        else: return
        def _show(): root.after(0,root.deiconify)
        def _sync(): threading.Thread(target=_full_sync,daemon=True).start()
        def _quit():
            global bg_active; bg_active=False
            try:
                if tray_icon: tray_icon.stop()
            except: pass
            root.after(0,root.quit)
        menu=pystray.Menu(
            pystray.MenuItem("Ouvrir MailGuard",_show,default=True),
            pystray.MenuItem("Synchroniser",_sync),
            pystray.MenuItem("Quitter",_quit))
        tray_icon=pystray.Icon("MailGuard",img,"MailGuard",menu=menu)
        threading.Thread(target=tray_icon.run,daemon=True).start()
    except: pass

def _on_close():
    if TRAY_OK: root.withdraw();
    if TRAY_OK and tray_icon is None: _build_tray()
    if not TRAY_OK: root.quit()

root.protocol("WM_DELETE_WINDOW",_on_close)

def _auto_refresh():
    if access_token and not running:
        threading.Thread(target=_sync_trash,daemon=True).start()
    root.after(30000,_auto_refresh)

root.after(30000,_auto_refresh)

# ══ DÉMARRAGE ═════════════════════════════════════════════════
write("Bienvenue dans MailGuard v2.1 👋","header")
write("")
write("  📨 Microsoft Outlook  — cliquez sur  « 📨 Microsoft »","muted")
write("  📧 Gmail              — cliquez sur  « 📧 Gmail »","muted")
if not GMAIL_OK:
    write("","muted")
    write("  ⚠ Support Gmail désactivé. Installez :","warn")
    write("    pip install google-auth-oauthlib google-api-python-client","warn")
write("")

threading.Thread(target=_bg_worker,daemon=True).start()

# Vérification silencieuse des mises à jour au démarrage
if _UPDATER_OK:
    root.after(3000, lambda: check_for_update(root, silent=True))

root.mainloop()
