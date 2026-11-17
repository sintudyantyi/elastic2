# ═══════════════════════════════════════════════════════════════
# ELASTIC BAND TRADING PLATFORM — SINGLE FILE
# Real OANDA balance, pluggable banking, 5 withdrawal rails.
# ═══════════════════════════════════════════════════════════════

from flask import Flask, request, jsonify, Response
import sqlite3, os, uuid, requests
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
from abc import ABC, abstractmethod
from typing import Dict

app = Flask(__name__)
DB = "dashboard.db"

# ── CONFIG ─────────────────────────────────────────────────────
ZAR_PER_USD = float(os.environ.get("ZAR_PER_USD", 18.50))
ZAR_PER_BTC = float(os.environ.get("ZAR_PER_BTC", 1200000))

CFG = {
    "OANDA_TOKEN": os.environ.get("OANDA_TOKEN", ""),
    "OANDA_ACCOUNT": os.environ.get("OANDA_ACCOUNT", ""),
    "OANDA_MODE": os.environ.get("OANDA_MODE", "practice"),
    "PAYPAL_MODE": os.environ.get("PAYPAL_MODE", "sandbox"),
    "PAYPAL_CLIENT_ID": os.environ.get("PAYPAL_CLIENT_ID", ""),
    "PAYPAL_CLIENT_SECRET": os.environ.get("PAYPAL_CLIENT_SECRET", ""),
    "PAYPAL_EMAIL": os.environ.get("PAYPAL_EMAIL", ""),
    "PAYPAL_RECEIVER_EMAIL": os.environ.get("PAYPAL_RECEIVER_EMAIL", ""),
    "PAYPAL_ME_USERNAME": os.environ.get("PAYPAL_ME_USERNAME", ""),
    "BANK_ACCOUNT": os.environ.get("BANK_ACCOUNT", ""),
    "BTC_WALLET": os.environ.get("BTC_WALLET", ""),
    "LUNO_RECEIVER_ACCOUNT": os.environ.get("LUNO_RECEIVER_ACCOUNT", ""),
    "SIGNAL_TOKEN": os.environ.get("SIGNAL_TOKEN", ""),
}

# ═══════════════════════════════════════════════════════════════
# BANKING ABSTRACTION LAYER (INLINE)
# ═══════════════════════════════════════════════════════════════

class BankProvider(ABC):
    @abstractmethod
    def get_provider_name(self) -> str: ...
    @abstractmethod
    def is_configured(self) -> bool: ...
    @abstractmethod
    def validate_account(self, account_number: str) -> Dict: ...
    @abstractmethod
    def prepare_withdrawal(self, amount: float, currency: str, destination: Dict) -> Dict: ...
    @abstractmethod
    def get_withdrawal_status(self, reference: str) -> Dict: ...


class ManualEFTAdapter(BankProvider):
    def get_provider_name(self): return "Manual EFT"
    def is_configured(self): return bool(os.environ.get("BANK_ACCOUNT"))
    def validate_account(self, account_number): return {"valid": True}
    def prepare_withdrawal(self, amount, currency, destination):
        ref = f"EFT_{uuid.uuid4().hex[:8].upper()}"
        return {"status": "ready", "reference": ref,
                "instructions": [f"Amount: {currency} {amount:.2f}",
                                 f"To: {destination.get('account_number','N/A')}",
                                 f"Reference: {ref}",
                                 "Complete via your bank app."],
                "redirect_url": None, "reason": None}
    def get_withdrawal_status(self, reference):
        return {"status": "pending"}


class OzowAdapter(BankProvider):
    def get_provider_name(self): return "Ozow"
    def is_configured(self): return bool(os.environ.get("OZOW_API_KEY"))
    def validate_account(self, account_number): return {"valid": True}
    def prepare_withdrawal(self, amount, currency, destination):
        return {"status": "ready", "reference": "OZOW_REF",
                "instructions": ["Redirect to Ozow payment page"],
                "redirect_url": "https://pay.ozow.com/...", "reason": None}
    def get_withdrawal_status(self, reference):
        return {"status": "pending"}


class PayFastAdapter(BankProvider):
    def get_provider_name(self): return "PayFast"
    def is_configured(self): return bool(os.environ.get("PAYFAST_MERCHANT_ID"))
    def validate_account(self, account_number): return {"valid": True}
    def prepare_withdrawal(self, amount, currency, destination):
        return {"status": "ready", "reference": "PF_REF",
                "instructions": ["Redirect to PayFast"],
                "redirect_url": "https://www.payfast.co.za/...", "reason": None}
    def get_withdrawal_status(self, reference):
        return {"status": "pending"}


class PeachAdapter(BankProvider):
    def get_provider_name(self): return "Peach Payments"
    def is_configured(self): return bool(os.environ.get("PEACH_ENTITY_ID"))
    def validate_account(self, account_number): return {"valid": True}
    def prepare_withdrawal(self, amount, currency, destination):
        return {"status": "ready", "reference": "PEACH_REF",
                "instructions": ["Redirect to Peach"],
                "redirect_url": "https://...", "reason": None}
    def get_withdrawal_status(self, reference):
        return {"status": "pending"}


class StitchAdapter(BankProvider):
    def get_provider_name(self): return "Stitch"
    def is_configured(self): return bool(os.environ.get("STITCH_CLIENT_ID"))
    def validate_account(self, account_number): return {"valid": True}
    def prepare_withdrawal(self, amount, currency, destination):
        return {"status": "ready", "reference": "STITCH_REF",
                "instructions": ["Redirect to Stitch"],
                "redirect_url": "https://...", "reason": None}
    def get_withdrawal_status(self, reference):
        return {"status": "pending"}


# ── BANK PROVIDER FACTORY ──────────────────────────────────────
BANK_PROVIDERS = {
    "manual": ManualEFTAdapter,
    "ozow": OzowAdapter,
    "payfast": PayFastAdapter,
    "peach": PeachAdapter,
    "stitch": StitchAdapter,
}

def get_bank_provider() -> BankProvider:
    name = os.environ.get("BANK_PROVIDER", "manual").lower()
    if name not in BANK_PROVIDERS:
        raise ValueError(f"Unknown bank provider: {name}")
    return BANK_PROVIDERS[name]()

def get_all_providers() -> list:
    return [{"id": k, "name": v().get_provider_name(),
             "configured": v().is_configured()} for k, v in BANK_PROVIDERS.items()]

# ═══════════════════════════════════════════════════════════════
# OANDA HELPERS
# ═══════════════════════════════════════════════════════════════
def _oanda_url():
    return ("https://api-fxtrade.oanda.com/v3" if CFG["OANDA_MODE"] == "live"
            else "https://api-fxpractice.oanda.com/v3")

def _oanda_headers():
    return {"Authorization": f"Bearer {CFG['OANDA_TOKEN']}",
            "Content-Type": "application/json"}

def oanda_balance():
    if not CFG["OANDA_TOKEN"] or not CFG["OANDA_ACCOUNT"]:
        return {"error": "OANDA not configured"}
    try:
        r = requests.get(f"{_oanda_url()}/accounts/{CFG['OANDA_ACCOUNT']}/summary",
                         headers=_oanda_headers(), timeout=30)
        if r.status_code != 200:
            return {"error": f"HTTP {r.status_code}"}
        d = r.json()["account"]
        return {"balance": float(d["balance"]), "currency": d["currency"],
                "nav": float(d["NAV"]), "unrealized_pl": float(d["unrealizedPL"])}
    except Exception as e:
        return {"error": str(e)}

def oanda_health():
    if not CFG["OANDA_TOKEN"] or not CFG["OANDA_ACCOUNT"]:
        return {"connected": False, "reason": "missing credentials"}
    bal = oanda_balance()
    if "error" in bal:
        return {"connected": False, "reason": bal["error"]}
    return {"connected": True, "mode": CFG["OANDA_MODE"],
            "account": CFG["OANDA_ACCOUNT"],
            "currency": bal["currency"], "balance": bal["balance"]}

def prepare_oanda_withdrawal(amount):
    bal = oanda_balance()
    if "error" in bal:
        return {"status": "error", "reason": bal["error"]}
    if amount > bal["balance"]:
        return {"status": "error",
                "reason": f"Insufficient. Available: {bal['balance']:.2f} {bal['currency']}"}
    portal = ("https://www.oanda.com/dashboard/deposit-withdraw/withdraw"
              if CFG["OANDA_MODE"] == "live"
              else "https://www.oanda.com/dashboard/deposit-withdraw/withdraw?practice=1")
    return {"status": "ready", "amount": amount, "currency": bal["currency"],
            "portal_url": portal,
            "instructions": [
                "1. Click 'Open OANDA Portal'",
                f"2. Enter amount: {amount:.2f} {bal['currency']}",
                "3. Select your verified bank account",
                "4. Confirm withdrawal",
                "5. Money arrives in 1-2 business days",
                "6. Return here and click 'Complete'"
            ]}

# ═══════════════════════════════════════════════════════════════
# PAYPAL HELPERS
# ═══════════════════════════════════════════════════════════════
def _pp_url():
    return ("https://api-m.paypal.com" if CFG["PAYPAL_MODE"] == "live"
            else "https://api-m.sandbox.paypal.com")

def pp_token():
    cid, sec = CFG["PAYPAL_CLIENT_ID"], CFG["PAYPAL_CLIENT_SECRET"]
    if not cid or not sec:
        return None, "PayPal credentials missing"
    try:
        r = requests.post(f"{_pp_url()}/v1/oauth2/token",
                          auth=(cid, sec),
                          data={"grant_type": "client_credentials"},
                          timeout=30)
        if r.status_code != 200:
            return None, f"Auth {r.status_code}"
        return r.json()["access_token"], None
    except Exception as e:
        return None, str(e)

def pp_health():
    t, e = pp_token()
    return (not e), (e or "connected")

def pp_payout(email, amount, note="Dashboard"):
    t, e = pp_token()
    if e:
        return {"status": "error", "reason": e}
    bid = f"DASH_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}"
    try:
        r = requests.post(f"{_pp_url()}/v1/payments/payouts",
                          headers={"Content-Type": "application/json",
                                   "Authorization": f"Bearer {t}"},
                          json={"sender_batch_header": {"sender_batch_id": bid},
                                "items": [{"recipient_type": "EMAIL",
                                           "amount": {"value": f"{amount:.2f}",
                                                      "currency": "USD"},
                                           "receiver": email, "note": note}]},
                          timeout=30)
        if r.status_code == 201:
            return {"status": "success",
                    "payout_id": r.json()["batch_header"]["payout_batch_id"]}
        return {"status": "error", "reason": r.text[:200]}
    except Exception as e:
        return {"status": "error", "reason": str(e)}

def pp_xclick(email, amount):
    return (f"https://www.paypal.com/cgi-bin/webscr?cmd=_xclick"
            f"&business={email}&amount={amount:.2f}&currency_code=USD&item_name=Withdrawal")

# ═══════════════════════════════════════════════════════════════
# EXCEL
# ═══════════════════════════════════════════════════════════════
EXCEL_FILE = "withdrawals.xlsx"

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "All Withdrawals"
        ws.append(["Date", "ID", "Method", "Amount (ZAR)", "Amount (USD)",
                   "Status", "Destination"])
        ws_pp = wb.create_sheet("PayPal")
        ws_pp.append(["Date", "ID", "Amount (ZAR)", "Amount (USD)",
                      "Status", "Email", "Note"])
        hf = Font(bold=True, color="FFFFFF")
        hfill = PatternFill(start_color="0A2540", end_color="0A2540",
                            fill_type="solid")
        for s in [ws, ws_pp]:
            for c in s[1]:
                c.font = hf
                c.fill = hfill
        wb.save(EXCEL_FILE)

init_excel()

def log_excel(wid, zar, usd, method, status, dest, note=""):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        wb["All Withdrawals"].append([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            wid, method, round(zar, 2), round(usd, 2), status, dest
        ])
        if method == "paypal":
            wb["PayPal"].append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                wid, round(zar, 2), round(usd, 2), status, dest, note
            ])
        wb.save(EXCEL_FILE)
    except Exception as e:
        print(f"Excel error: {e}")

# ═══════════════════════════════════════════════════════════════
# DATABASE
# ═══════════════════════════════════════════════════════════════
def init_db():
    c = sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS trades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT, signal TEXT, price REAL,
        sl REAL DEFAULT 0, tp REAL DEFAULT 0, units INTEGER DEFAULT 10,
        status TEXT DEFAULT 'open', pnl_usd REAL DEFAULT 0,
        pnl_zar REAL DEFAULT 0, close_price REAL DEFAULT 0,
        source TEXT DEFAULT 'manual', confidence INTEGER DEFAULT 0)""")
    c.execute("""CREATE TABLE IF NOT EXISTS withdrawals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT, amount_zar REAL, method TEXT,
        status TEXT DEFAULT 'pending', note TEXT DEFAULT '',
        pay_link TEXT DEFAULT '', reference TEXT DEFAULT '')""")
    c.commit()
    c.close()

init_db()

def q(sql, args=(), one=False):
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    cur = c.execute(sql, args)
    rows = cur.fetchall()
    c.close()
    return (rows[0] if rows else None) if one else rows

def ex(sql, args=()):
    c = sqlite3.connect(DB)
    c.execute(sql, args)
    c.commit()
    c.close()

# ═══════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route('/health')
def health():
    pp_ok, pp_msg = pp_health()
    oa = oanda_health()
    provider = get_bank_provider()
    return jsonify({"status": "ok",
                    "paypal": {"connected": pp_ok, "info": pp_msg},
                    "oanda": oa,
                    "bank": {"provider": provider.get_provider_name(),
                             "configured": provider.is_configured()},
                    "time": datetime.utcnow().isoformat()})

@app.route('/api/summary')
def summary():
    oa = oanda_balance()
    oa_connected = "error" not in oa
    real_zar = (oa["balance"] if oa_connected and oa["currency"] == "ZAR"
                else (oa["balance"] * ZAR_PER_USD if oa_connected else 0.0))

    total_pnl_zar = q("SELECT COALESCE(SUM(pnl_zar),0) AS s FROM trades WHERE status='closed'", one=True)["s"]
    open_trades   = q("SELECT COUNT(*) AS c FROM trades WHERE status='open'", one=True)["c"]
    closed_trades = q("SELECT COUNT(*) AS c FROM trades WHERE status='closed'", one=True)["c"]
    wins          = q("SELECT COUNT(*) AS c FROM trades WHERE status='closed' AND pnl_zar>0", one=True)["c"]
    win_rate      = round((wins / closed_trades * 100), 1) if closed_trades else 0

    wd   = q("SELECT COALESCE(SUM(amount_zar),0) AS s FROM withdrawals WHERE status='completed'", one=True)["s"]
    pend = q("SELECT COALESCE(SUM(amount_zar),0) AS s FROM withdrawals WHERE status='pending'", one=True)["s"]

    return jsonify({
        "balance_zar":       round(real_zar, 2),
        "total_pnl_zar":     round(total_pnl_zar, 2),
        "withdrawn_zar":     round(wd, 2),
        "pending_zar":       round(pend, 2),
        "available_zar":     round(max(real_zar - wd - pend, 0), 2),
        "open_trades":       open_trades,
        "win_rate":          win_rate,
        "crypto_equiv_btc":  round(real_zar / ZAR_PER_BTC, 8),
        "oanda_connected":   oa_connected,
        "oanda_balance":     oa.get("balance") if oa_connected else None,
        "oanda_currency":    oa.get("currency") if oa_connected else None,
        "oanda_mode":        CFG["OANDA_MODE"],
        "oanda_error":       oa.get("error") if not oa_connected else None,
        "zar_per_usd":       ZAR_PER_USD,
    })

@app.route('/api/oanda/balance')
def oanda_bal():
    return jsonify(oanda_balance())

@app.route('/signal', methods=['POST'])
def signal():
    data = request.json or {}
    if CFG["SIGNAL_TOKEN"] and data.get("token") != CFG["SIGNAL_TOKEN"]:
        return jsonify({"status": "rejected", "reason": "invalid token"}), 401
    if data.get("signal") not in ("LONG", "SHORT"):
        return jsonify({"status": "rejected", "reason": "invalid signal"}), 400
    price = float(data.get("price", 0))
    if price <= 0:
        return jsonify({"status": "rejected", "reason": "invalid price"}), 400
    ex("""INSERT INTO trades (timestamp, signal, price, sl, tp, units, source, confidence)
          VALUES (?,?,?,?,?,?,?,?)""",
       (datetime.utcnow().isoformat(), data["signal"], price,
        float(data.get("sl", 0)), float(data.get("tp", 0)),
        int(data.get("units", 10)), "indicator", int(data.get("confidence", 0))))
    tid = q("SELECT last_insert_rowid() AS id", one=True)["id"]
    return jsonify({"status": "logged", "trade_id": tid})

@app.route('/api/withdraw', methods=['POST'])
def withdraw():
    data = request.json or {}
    zar = float(data.get("amount_zar", 0))
    method = data.get("method", "").lower()

    if zar <= 0:
        return jsonify({"status": "error", "reason": "invalid amount"}), 400
    if method not in ("oanda", "paypal", "btc", "luno", "bank"):
        return jsonify({"status": "error", "reason": "invalid method"}), 400

    s = summary().get_json()
    if zar > s["available_zar"]:
        return jsonify({"status": "rejected",
                        "reason": f"Insufficient. Available: R{s['available_zar']:.2f}"}), 400

    usd = round(zar / ZAR_PER_USD, 2)
    dest = ""
    note = ""
    link = ""
    status = "pending"
    reference = ""

    # OANDA
    if method == "oanda":
        bal = oanda_balance()
        if "error" in bal:
            return jsonify({"status": "error", "reason": f"OANDA: {bal['error']}"}), 500
        amt = usd if bal["currency"] == "USD" else zar
        r = prepare_oanda_withdrawal(amt)
        if r.get("status") == "error":
            return jsonify({"status": "rejected", "reason": r["reason"]}), 400
        dest = f"OANDA → verified bank ({CFG['OANDA_ACCOUNT']})"
        note = f"OANDA withdrawal READY: {amt:.2f} {bal['currency']}"
        link = r["portal_url"]

    # PAYPAL
    elif method == "paypal":
        recv = CFG["PAYPAL_RECEIVER_EMAIL"]
        if not recv or "@" not in recv:
            return jsonify({"status": "error",
                            "reason": "PAYPAL_RECEIVER_EMAIL not set"}), 400
        if recv.lower() == CFG["PAYPAL_EMAIL"].lower():
            return jsonify({"status": "error",
                            "reason": "Receiver = sender (self-payment blocked)"}), 400
        dest = recv
        ok, _ = pp_health()
        if ok:
            p = pp_payout(recv, usd)
            if p.get("status") == "success":
                status = "completed"
                link = p["payout_id"]
                note = f"PayPal Payout ${usd:.2f}"
            else:
                link = pp_xclick(recv, usd)
                note = f"Fallback link for {recv}"
        else:
            link = pp_xclick(recv, usd)
            note = f"Manual link for {recv}"

    # BTC
    elif method == "btc":
        recv = CFG["BTC_WALLET"]
        if not recv:
            return jsonify({"status": "error", "reason": "BTC_WALLET not set"}), 400
        btc = zar / ZAR_PER_BTC
        dest = recv
        note = f"BTC: {btc:.8f} to {recv}"

    # LUNO
    elif method == "luno":
        recv = CFG["LUNO_RECEIVER_ACCOUNT"] or "SET_LUNO_ENV"
        btc = zar / ZAR_PER_BTC
        dest = recv
        note = f"Luno: {btc:.8f} BTC to {recv}"

    # BANK (pluggable)
    elif method == "bank":
        provider = get_bank_provider()
        if not provider.is_configured():
            return jsonify({"status": "error",
                            "reason": f"{provider.get_provider_name()} not configured"}), 400
        r = provider.prepare_withdrawal(
            amount=zar,
            currency="ZAR",
            destination={"account_number": CFG["BANK_ACCOUNT"]}
        )
        if r.get("status") == "error":
            return jsonify({"status": "rejected", "reason": r.get("reason")}), 400
        dest = f"{provider.get_provider_name()} → {CFG['BANK_ACCOUNT']}"
        note = f"Withdrawal via {provider.get_provider_name()}"
        link = r.get("redirect_url") or ""
        reference = r.get("reference", "")
        status = "pending"

    ex("""INSERT INTO withdrawals
          (timestamp, amount_zar, method, status, note, pay_link, reference)
          VALUES (?,?,?,?,?,?,?)""",
       (datetime.utcnow().isoformat(), zar, method, status, note, link, reference))
    wid = q("SELECT last_insert_rowid() AS id", one=True)["id"]
    log_excel(wid, zar, usd, method, status, dest, note)

    return jsonify({"status": status, "withdrawal_id": wid,
                    "amount_zar": round(zar, 2), "amount_usd": usd,
                    "method": method, "note": note,
                    "destination": dest, "pay_link": link,
                    "reference": reference})

@app.route('/api/withdrawals')
def withdrawals():
    return jsonify([dict(r) for r in q(
        "SELECT * FROM withdrawals ORDER BY id DESC LIMIT 100")])

@app.route('/api/withdraw/<int:wid>/complete', methods=['POST'])
def complete_wd(wid):
    ex("UPDATE withdrawals SET status='completed' WHERE id=?", (wid,))
    return jsonify({"status": "completed"})

@app.route('/api/withdraw/<int:wid>/cancel', methods=['POST'])
def cancel_wd(wid):
    ex("UPDATE withdrawals SET status='cancelled' WHERE id=?", (wid,))
    return jsonify({"status": "cancelled"})

@app.route('/api/bank/providers')
def bank_providers():
    return jsonify({"providers": get_all_providers(),
                    "active": get_bank_provider().get_provider_name()})

@app.route('/download/excel')
def download_excel():
    if not os.path.exists(EXCEL_FILE):
        return jsonify({"error": "no file"}), 404
    with open(EXCEL_FILE, "rb") as f:
        data = f.read()
    return Response(data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=withdrawals.xlsx"})

@app.route('/api/paypal_me_link')
def pp_me():
    u = CFG["PAYPAL_ME_USERNAME"].strip().lstrip("@")
    if not u:
        return jsonify({"status": "unconfigured"})
    return jsonify({"status": "ok", "link": f"https://paypal.me/{u}"})

@app.route('/api/reset', methods=['POST'])
def reset():
    ex("DELETE FROM trades")
    ex("DELETE FROM withdrawals")
    return jsonify({"status": "reset"})

# ═══════════════════════════════════════════════════════════════
# DASHBOARD HTML
# ═══════════════════════════════════════════════════════════════

DASHBOARD_HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>💰 Elastic Band Trading Dashboard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0a1929;color:#e6edf3;font-family:Inter,system-ui,sans-serif;padding:24px;min-height:100vh}
header{display:flex;justify-content:space-between;align-items:center;margin-bottom:28px;padding-bottom:16px;border-bottom:2px solid #1e4976;flex-wrap:wrap;gap:12px}
h1{color:#FFD700;font-size:1.5rem}
h2{color:#00A8A8;font-size:1.1rem;margin-bottom:14px}
.status{padding:6px 14px;border-radius:20px;background:#164b2c;color:#32CD32;font-size:.85rem;font-weight:600}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-bottom:24px}
.card{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:18px}
.card.highlight{border-color:#FFD700;box-shadow:0 0 20px rgba(255,215,0,.15)}
.card.crypto{border-color:#00b5e2;box-shadow:0 0 20px rgba(0,181,226,.15)}
.card.oanda{border-color:#32CD32;box-shadow:0 0 20px rgba(50,205,50,.15)}
.label{color:#8899aa;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.metric{font-size:1.35rem;font-weight:700;color:#00A8A8}
.metric.small{font-size:1rem;font-weight:600;color:#e6edf3}
.positive{color:#32CD32!important}
.negative{color:#FF4444!important}
.gold{color:#FFD700!important}
.btc{color:#00b5e2!important}
.actions{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:24px}
@media(max-width:900px){.actions{grid-template-columns:1fr}body{padding:16px}}
.panel{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:20px}
.panel.receive{border-color:#00b5e2;box-shadow:0 0 15px rgba(0,181,226,.1)}
.hint{color:#8899aa;font-size:.8rem;margin-bottom:14px}
.row{display:flex;flex-wrap:wrap;gap:10px;align-items:center}
input,select,button{padding:10px 14px;border-radius:8px;border:1px solid #1e4976;background:#0a1929;color:#e6edf3;font-size:.9rem;font-family:inherit}
input{width:140px}
button{cursor:pointer;font-weight:600;transition:.2s}
button:hover{opacity:.85}
.btn-buy{background:#164b2c;color:#32CD32;border-color:#32CD32}
.btn-sell{background:#4b1616;color:#FF4444;border-color:#FF4444}
.btn-small{padding:6px 10px;font-size:.75rem;background:#00A8A8;color:#0a1929;border:none;margin-left:4px}
.btn-danger{background:#4b1616;color:#FF4444;border-color:#FF4444}
.btn-paypal{background:#003087;color:#fff;border:none}
.btn-luno{background:#00b5e2;color:#000;border:none}
.btn-bank{background:#FFD700;color:#0a1929;border:none}
.btn-btc{background:#f7931a;color:#000;border:none}
.btn-oanda{background:#c8102e;color:#fff;border:none}
.btn-copy{background:#00b5e2;color:#000;border:none}
.btn-excel{background:#1D6F42;color:#fff;border:none;padding:10px 16px;border-radius:8px;font-weight:600;cursor:pointer;text-decoration:none;display:inline-block}
.msg{margin-top:12px;font-weight:600;font-size:.85rem;word-break:break-word}
.msg a{color:#FFD700;font-weight:700}
.table-section{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:20px;margin-bottom:20px;overflow-x:auto}
table{width:100%;border-collapse:collapse;min-width:700px}
th{text-align:left;padding:10px;color:#00A8A8;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #1e4976}
td{padding:9px;border-bottom:1px solid #1e4976;font-size:.83rem}
tr:hover{background:#143557}
footer{margin-top:32px;padding-top:16px;border-top:2px solid #1e4976;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.footer-note{color:#8899aa;font-size:.8rem}
.withdraw-buttons{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
.receive-box{display:flex;gap:10px;flex-wrap:wrap;margin-top:10px}
.receive-box input{flex:1;min-width:220px;font-family:monospace;font-size:.85rem}
</style>
</head>
<body>

<header>
  <h1>💰 Elastic Band Trading Dashboard <span style="font-size:.75rem;color:#8899aa">(ZAR)</span></h1>
  <div class="status" id="status">● Loading...</div>
</header>

<section class="grid">
  <div class="card highlight"><div class="label">Balance (ZAR)</div><div class="metric gold" id="balance">R0.00</div></div>
  <div class="card oanda"><div class="label">OANDA Real</div><div class="metric positive" id="oanda-balance">—</div></div>
  <div class="card crypto"><div class="label">Crypto Equiv (BTC)</div><div class="metric btc" id="crypto-eq">0.00000000</div></div>
  <div class="card"><div class="label">Total P&L (ZAR)</div><div class="metric" id="pnl-zar">R0.00</div></div>
  <div class="card"><div class="label">Win Rate</div><div class="metric" id="winrate">0%</div></div>
  <div class="card"><div class="label">Open</div><div class="metric" id="open">0</div></div>
  <div class="card"><div class="label">Withdrawn (ZAR)</div><div class="metric" id="withdrawn">R0.00</div></div>
  <div class="card highlight"><div class="label">Available (ZAR)</div><div class="metric gold" id="available">R0.00</div></div>
</section>

<section class="actions">
  <div class="panel">
    <h2>🧪 Simulate Trade</h2>
    <p class="hint">XAUUSD price · units · P&L auto-converted to ZAR</p>
    <div class="row">
      <input type="number" id="sim-price" step="0.01" value="2065.00"/>
      <input type="number" id="sim-units" step="1" value="10"/>
    </div>
    <div class="row" style="margin-top:10px">
      <button onclick="simulateTrade('LONG')" class="btn-buy">BUY (LONG)</button>
      <button onclick="simulateTrade('SHORT')" class="btn-sell">SELL (SHORT)</button>
    </div>
  </div>

  <div class="panel">
    <h2>💸 Withdraw</h2>
    <p class="hint">OANDA = real broker withdrawal · others = manual</p>
    <div class="row">
      <input type="number" id="wd-amount" step="0.01" placeholder="Amount in R"/>
    </div>
    <div class="withdraw-buttons">
      <button onclick="requestWithdrawal('oanda')" class="btn-oanda">🏦 OANDA</button>
      <button onclick="requestWithdrawal('paypal')" class="btn-paypal">💳 PayPal</button>
      <button onclick="requestWithdrawal('btc')" class="btn-btc">₿ BTC</button>
      <button onclick="requestWithdrawal('luno')" class="btn-luno">💠 Luno</button>
      <button onclick="requestWithdrawal('bank')" class="btn-bank">🏦 Bank</button>
    </div>
    <div class="msg" id="wd-msg"></div>
    <div style="margin-top:14px">
      <a href="/download/excel" class="btn-excel">📊 Download Excel</a>
      <a href="/setup" class="btn-excel" style="background:#003087;margin-left:8px">⚙️ Setup Guide</a>
    </div>
  </div>
</section>

<section class="actions">
  <div class="panel receive" style="grid-column:1 / -1;">
    <h2>💰 Receive Payment (From Clients)</h2>
    <p class="hint">Share this link so clients can pay you.</p>
    <div class="receive-box">
      <input type="text" id="paypal-me-link" readonly placeholder="Loading..."/>
      <button onclick="copyLink()" class="btn-copy">📋 Copy</button>
      <a id="paypal-me-open" href="#" target="_blank" class="btn-excel" style="background:#003087;">→ Open</a>
    </div>
    <div class="msg" id="me-msg"></div>
  </div>
</section>

<section class="table-section">
  <h2>📈 Open Trades</h2>
  <table id="open-trades">
    <thead><tr><th>ID</th><th>Time</th><th>Signal</th><th>Price</th><th>SL</th><th>TP</th><th>Units</th><th>Action</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<section class="table-section">
  <h2>📊 Closed Trades</h2>
  <table id="closed-trades">
    <thead><tr><th>ID</th><th>Time</th><th>Signal</th><th>Price</th><th>Close</th><th>P&L USD</th><th>P&L ZAR</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<section class="table-section">
  <h2>💸 Withdrawal History</h2>
  <table id="wd-history">
    <thead><tr><th>ID</th><th>Time</th><th>Amount (ZAR)</th><th>Method</th><th>Status</th><th>Action</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<footer>
  <button onclick="resetAll()" class="btn-danger">🔄 Reset</button>
  <span class="footer-note" id="mode-note">OANDA Connected</span>
</footer>

<script>
const fmtZAR = n => 'R' + Number(n||0).toLocaleString('en-ZA',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtUSD = n => '$' + Number(n||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtBTC = n => Number(n||0).toFixed(8);

function setText(id,v){document.getElementById(id).textContent=v}
function setColor(id,txt,val){
  const el=document.getElementById(id);
  el.textContent=txt;
  el.classList.remove('positive','negative');
  if(val>0)el.classList.add('positive');
  if(val<0)el.classList.add('negative');
}
async function j(url, opts){return await (await fetch(url, opts||{})).json()}

async function refreshAll(){
  const s = await j('/api/summary');
  setText('balance', fmtZAR(s.balance_zar));
  setText('crypto-eq', fmtBTC(s.crypto_equiv_btc) + ' BTC');
  setColor('pnl-zar', fmtZAR(s.total_pnl_zar), s.total_pnl_zar);
  setText('winrate', s.win_rate + '%');
  setText('open', s.open_trades);
  setText('withdrawn', fmtZAR(s.withdrawn_zar));
  setText('available', fmtZAR(s.available_zar));
  if(s.oanda_connected && s.oanda_balance !== null){
    const ob = document.getElementById('oanda-balance');
    ob.textContent = s.oanda_balance + ' ' + s.oanda_currency;
    ob.className = 'metric positive';
    setText('mode-note', `OANDA ${s.oanda_mode.toUpperCase()} · Real balance: ${s.oanda_balance} ${s.oanda_currency}`);
  } else {
    document.getElementById('oanda-balance').textContent = 'Not connected';
    document.getElementById('oanda-balance').className = 'metric negative';
    setText('mode-note', `OANDA error: ${s.oanda_error || 'check /setup'}`);
  }
  document.getElementById('status').textContent = '● Live';

  const trades = await j('/api/trades');
  const ob = document.querySelector('#open-trades tbody');
  const cb = document.querySelector('#closed-trades tbody');
  ob.innerHTML = ''; cb.innerHTML = '';
  for(const t of trades){
    if(t.status==='open'){
      ob.innerHTML += `<tr>
        <td>${t.id}</td><td>${t.timestamp.slice(0,16).replace('T',' ')}</td>
        <td class="${t.signal==='LONG'?'positive':'negative'}">${t.signal}</td>
        <td>$${t.price.toFixed(2)}</td><td>$${t.sl.toFixed(2)}</td>
        <td>$${t.tp.toFixed(2)}</td><td>${t.units}</td>
        <td><input type="number" id="close-${t.id}" step="0.01" placeholder="Close" style="width:80px;padding:6px"/>
            <button class="btn-small" onclick="closeTrade(${t.id})">Close</button></td>
      </tr>`;
    } else {
      cb.innerHTML += `<tr>
        <td>${t.id}</td><td>${t.timestamp.slice(0,16).replace('T',' ')}</td>
        <td class="${t.signal==='LONG'?'positive':'negative'}">${t.signal}</td>
        <td>$${t.price.toFixed(2)}</td><td>$${t.close_price.toFixed(2)}</td>
        <td class="${t.pnl_usd>=0?'positive':'negative'}">$${t.pnl_usd.toFixed(2)}</td>
        <td class="${t.pnl_zar>=0?'positive':'negative'}">${fmtZAR(t.pnl_zar)}</td>
      </tr>`;
    }
  }

  const wds = await j('/api/withdrawals');
  const wb = document.querySelector('#wd-history tbody');
  wb.innerHTML = '';
  for(const w of wds){
    let actions = '';
    if(w.status==='pending'){
      if(w.pay_link && w.pay_link.startsWith('http')){
        actions += `<a href="${w.pay_link}" target="_blank" class="btn-small" style="text-decoration:none;background:#FFD700;color:#0a1929">Open</a> `;
      }
      actions += `<button class="btn-small" onclick="completeWd(${w.id})">✓</button> `;
      actions += `<button class="btn-small btn-danger" onclick="cancelWd(${w.id})">✕</button>`;
    } else actions = w.status==='completed' ? '✅' : '❌';
    wb.innerHTML += `<tr>
      <td>${w.id}</td><td>${w.timestamp.slice(0,16).replace('T',' ')}</td>
      <td>${fmtZAR(w.amount_zar)}</td><td>${w.method}</td><td>${w.status}</td><td>${actions}</td>
    </tr>`;
  }
}

async function loadPP(){
  const r = await j('/api/paypal_me_link');
  const input = document.getElementById('paypal-me-link');
  const openBtn = document.getElementById('paypal-me-open');
  if(r.status === 'ok'){
    input.value = r.link;
    openBtn.href = r.link;
    openBtn.style.display = 'inline-block';
  } else {
    input.value = r.message || 'Not configured';
    openBtn.style.display = 'none';
  }
}

function copyLink(){
  const input = document.getElementById('paypal-me-link');
  const msg = document.getElementById('me-msg');
  if(!input.value || input.value.includes('not') || input.value.startsWith('Set ')){
    msg.textContent = '⚠️ Configure PAYPAL_ME_USERNAME on Render';
    msg.className = 'msg negative'; return;
  }
  input.select(); document.execCommand('copy');
  msg.textContent = '✅ Copied!'; msg.className = 'msg positive';
  setTimeout(() => { msg.textContent = ''; }, 2500);
}

async function simulateTrade(signal){
  const price = parseFloat(document.getElementById('sim-price').value);
  const units = parseInt(document.getElementById('sim-units').value);
  if(!price||price<=0){alert('Enter valid price');return}
  await j('/api/manual_trade', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({signal, price, units})
  });
  refreshAll();
}

async function closeTrade(id){
  const price = parseFloat(document.getElementById('close-'+id).value);
  if(!price||price<=0){alert('Enter close price');return}
  const r = await j('/api/close_trade/'+id, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({close_price: price})
  });
  alert(`Closed. P&L: $${r.pnl_usd.toFixed(2)} / ${fmtZAR(r.pnl_zar)}`);
  refreshAll();
}

async function requestWithdrawal(method){
  const amount = parseFloat(document.getElementById('wd-amount').value);
  const msg = document.getElementById('wd-msg');
  if(!amount||amount<=0){msg.textContent='❌ Enter valid amount';msg.className='msg negative';return}
  msg.textContent = '⏳ Processing...'; msg.className = 'msg';
  const r = await j('/api/withdraw', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({amount_zar: amount, method})
  });
  if(r.status==='rejected' || r.status==='error'){
    msg.textContent='❌ '+(r.reason||'failed'); msg.className='msg negative';
  } else if(r.pay_link && r.pay_link.startsWith('http')){
    msg.innerHTML = `<div style="background:#164b2c;padding:14px;border-radius:8px;margin-top:8px">
      <strong style="color:#32CD32">✅ ${r.method.toUpperCase()} WITHDRAWAL READY</strong><br>
      ${r.note}<br>
      <a href="${r.pay_link}" target="_blank" style="background:#FFD700;color:#0a1929;padding:8px 16px;border-radius:6px;text-decoration:none;font-weight:700;display:inline-block;margin-top:8px">→ Open ${r.method.toUpperCase()} Portal</a></div>`;
    msg.className='msg';
  } else {
    msg.textContent=`✅ #${r.withdrawal_id} · ${fmtZAR(r.amount_zar)} — ${r.note}`;
    msg.className='msg positive';
  }
  refreshAll();
}

async function completeWd(id){await fetch('/api/withdraw/'+id+'/complete',{method:'POST'});refreshAll()}
async function cancelWd(id){await fetch('/api/withdraw/'+id+'/cancel',{method:'POST'});refreshAll()}
async function resetAll(){if(!confirm('Reset all data?'))return;await fetch('/api/reset',{method:'POST'});refreshAll()}

refreshAll();
loadPP();
setInterval(refreshAll, 15000);
</script>
</body>
</html>
"""

# ═══════════════════════════════════════════════════════════════
# SETUP WIZARD HTML
# ═══════════════════════════════════════════════════════════════

SETUP_HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>⚙️ OANDA Setup Wizard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0a1929;color:#e6edf3;font-family:Inter,system-ui,sans-serif;padding:24px;min-height:100vh;line-height:1.6}
.container{max-width:900px;margin:0 auto}
header{display:flex;justify-content:space-between;align-items:center;margin-bottom:32px;padding-bottom:16px;border-bottom:2px solid #1e4976;flex-wrap:wrap;gap:12px}
h1{color:#FFD700;font-size:1.8rem}
h2{color:#00A8A8;font-size:1.3rem;margin:32px 0 16px;padding-left:12px;border-left:4px solid #FFD700}
h3{color:#e6edf3;font-size:1.1rem;margin:0 0 12px}
a{color:#FFD700;text-decoration:none;font-weight:600}
a:hover{text-decoration:underline}
.step{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:24px;margin-bottom:20px;position:relative}
.step-number{position:absolute;top:-14px;left:24px;background:#FFD700;color:#0a1929;width:40px;height:40px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:1.2rem}
.step-content{margin-top:12px}
.code-box{background:#0a1929;border:1px solid #1e4976;border-radius:8px;padding:14px;margin:12px 0;font-family:monospace;font-size:.85rem;overflow-x:auto;position:relative}
.copy-btn{position:absolute;top:8px;right:8px;background:#00A8A8;color:#0a1929;border:none;padding:4px 10px;border-radius:4px;font-size:.75rem;font-weight:600;cursor:pointer}
.badge{display:inline-block;padding:4px 10px;border-radius:12px;font-size:.75rem;font-weight:700;margin-left:8px}
.badge-green{background:#164b2c;color:#32CD32}
.badge-red{background:#4b1616;color:#FF4444}
.badge-yellow{background:#4b3d16;color:#FFD700}
.note{background:#4b3d16;border-left:4px solid #FFD700;padding:14px 18px;border-radius:8px;margin:16px 0;color:#FFD700}
.warning{background:#4b1616;border-left:4px solid #FF4444;padding:14px 18px;border-radius:8px;margin:16px 0;color:#FF4444}
.success{background:#164b2c;border-left:4px solid #32CD32;padding:14px 18px;border-radius:8px;margin:16px 0;color:#32CD32}
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media(max-width:700px){.grid-2{grid-template-columns:1fr}}
.field{background:#0a1929;border:1px solid #1e4976;border-radius:8px;padding:12px;margin:8px 0}
.field-label{color:#8899aa;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
.field-value{color:#e6edf3;font-family:monospace;font-size:.9rem;word-break:break-all}
.status-panel{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:20px;margin:24px 0}
.status-item{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #1e4976}
.status-item:last-child{border-bottom:none}
.cta-button{display:inline-block;background:#FFD700;color:#0a1929;padding:12px 24px;border-radius:8px;font-weight:700;margin:8px 8px 8px 0;text-decoration:none}
.cta-button:hover{background:#ffed4e;text-decoration:none}
.cta-button.secondary{background:#1e4976;color:#e6edf3}
nav{display:flex;gap:16px;flex-wrap:wrap}
nav a{padding:8px 16px;background:#112d4e;border-radius:8px;color:#e6edf3;font-weight:500}
nav a:hover{background:#1e4976;text-decoration:none}
</style>
</head>
<body>
<div class="container">

<header>
  <h1>⚙️ OANDA Setup Wizard</h1>
  <nav><a href="/">← Dashboard</a><a href="/health">Health</a></nav>
</header>

<div class="note">
  <strong>⏱️ Total time: 10-15 minutes</strong><br>
  Follow these steps exactly. Each one must be completed before the next.
</div>

<div class="step">
  <div class="step-number">1</div>
  <div class="step-content">
    <h3>Create / Log into OANDA Account</h3>
    <p>If you don't have one, create a free account. If you do, log in.</p>
    <a href="https://www.oanda.com/register/" target="_blank" class="cta-button">→ Create OANDA Account</a>
    <a href="https://www.oanda.com/dashboard/" target="_blank" class="cta-button secondary">→ Log In</a>
    <div class="note"><strong>💡 Tip:</strong> Start with a <strong>practice account</strong> to test safely.</div>
  </div>
</div>

<div class="step">
  <div class="step-number">2</div>
  <div class="step-content">
    <h3>Generate API Access Token</h3>
    <p>The token lets the dashboard read your real balance and place trades.</p>
    <ol style="padding-left:20px;margin:12px 0">
      <li>Log into OANDA</li>
      <li>Click <strong>"My Account"</strong></li>
      <li>Click <strong>"Manage API Access"</strong></li>
      <li>Click <strong>"Generate"</strong></li>
      <li>Copy the token immediately</li>
    </ol>
    <a href="https://www.oanda.com/dashboard/account/api" target="_blank" class="cta-button">→ Open API Management</a>
    <div class="warning"><strong>⚠️ NEVER share this token.</strong> Only paste into Render environment variables.</div>
  </div>
</div>

<div class="step">
  <div class="step-number">3</div>
  <div class="step-content">
    <h3>Link Your SA Bank Account</h3>
    <p>OANDA sends withdrawals to your <strong>verified</strong> bank account only.</p>
    <ol style="padding-left:20px;margin:12px 0">
      <li>Click <strong>"Deposit / Withdraw"</strong></li>
      <li>Click <strong>"Add Withdrawal Method"</strong></li>
      <li>Enter your bank details (FNB/Standard/ABSA/Capitec)</li>
      <li>Submit for verification</li>
      <li>Wait 1-2 business days</li>
    </ol>
    <a href="https://www.oanda.com/dashboard/deposit-withdraw/" target="_blank" class="cta-button">→ Add Bank Account</a>
  </div>
</div>

<div class="step">
  <div class="step-number">4</div>
  <div class="step-content">
    <h3>Choose Practice vs Live Mode</h3>
    <div class="grid-2">
      <div class="field" style="border-color:#32CD32">
        <div class="field-label" style="color:#32CD32">✅ PRACTICE</div>
        <p style="margin-top:8px">Fake money. Real API. Safe testing.</p>
        <div class="code-box">OANDA_MODE=practice</div>
      </div>
      <div class="field" style="border-color:#FF4444">
        <div class="field-label" style="color:#FF4444">💰 LIVE</div>
        <p style="margin-top:8px">Real money. Real withdrawals.</p>
        <div class="code-box">OANDA_MODE=live</div>
      </div>
    </div>
  </div>
</div>

<div class="step">
  <div class="step-number">5</div>
  <div class="step-content">
    <h3>Add to Render Environment Variables</h3>
    <p>Go to Render dashboard and add these variables:</p>
    <a href="https://dashboard.render.com/" target="_blank" class="cta-button">→ Open Render Dashboard</a>
    <div class="code-box" style="margin-top:16px">
      <button class="copy-btn" onclick="copyToClipboard(this)">Copy</button>
OANDA_TOKEN = eyJhbGciOiJIUzI1NiIs...
OANDA_ACCOUNT = 101-004-1234567-001
OANDA_MODE = practice
    </div>
    <div class="note"><strong>💡 Replace placeholders</strong> with your actual values from Step 2.</div>
  </div>
</div>

<div class="step">
  <div class="step-number">6</div>
  <div class="step-content">
    <h3>Verify Connection</h3>
    <a href="/health" class="cta-button">→ Check /health</a>
    <a href="/api/oanda/balance" class="cta-button secondary">→ Check Balance</a>
    <div class="success"><strong>✅ Success:</strong> /health shows <code>"oanda":{"connected":true}</code></div>
  </div>
</div>

<div class="step">
  <div class="step-number">7</div>
  <div class="step-content">
    <h3>Test the Dashboard</h3>
    <p>Dashboard should now show your REAL OANDA balance.</p>
    <a href="/" class="cta-button">→ Open Dashboard</a>
  </div>
</div>

<h2>🔍 Live Status</h2>
<div class="status-panel">
  <div class="status-item"><span>OANDA API</span><span id="api-status" class="badge badge-yellow">Checking...</span></div>
  <div class="status-item"><span>Mode</span><span id="mode-status" class="badge badge-yellow">Unknown</span></div>
  <div class="status-item"><span>Account</span><span id="account-status" class="badge badge-yellow">Not set</span></div>
  <div class="status-item"><span>Balance</span><span id="balance-status" class="badge badge-yellow">Loading...</span></div>
  <button onclick="refreshStatus()" class="cta-button" style="margin-top:16px">🔄 Refresh</button>
</div>

<footer style="margin-top:48px;padding-top:24px;border-top:2px solid #1e4976;text-align:center;color:#8899aa;font-size:.85rem">
  <p>Setup Wizard · Elastic Band Dashboard</p>
  <p style="margin-top:8px">🇿🇦 WINK LAPHO 😏</p>
</footer>

</div>

<script>
async function refreshStatus(){
  try{
    const r = await fetch('/health');
    const data = await r.json();
    const oa = data.oanda || {};
    document.getElementById('api-status').textContent = oa.connected ? '✅ Connected' : '❌ Not Connected';
    document.getElementById('api-status').className = 'badge ' + (oa.connected ? 'badge-green' : 'badge-red');
    document.getElementById('mode-status').textContent = oa.mode || 'Unknown';
    document.getElementById('mode-status').className = 'badge badge-green';
    document.getElementById('account-status').textContent = oa.account || 'Not set';
    document.getElementById('account-status').className = 'badge ' + (oa.account ? 'badge-green' : 'badge-red');
    const balRes = await fetch('/api/oanda/balance');
    const balData = await balRes.json();
    if(balData.balance !== undefined){
      document.getElementById('balance-status').textContent = balData.balance + ' ' + balData.currency;
      document.getElementById('balance-status').className = 'badge badge-green';
    } else {
      document.getElementById('balance-status').textContent = balData.error || 'Error';
      document.getElementById('balance-status').className = 'badge badge-red';
    }
  } catch(e){
    document.getElementById('api-status').textContent = '❌ Error';
    document.getElementById('api-status').className = 'badge badge-red';
  }
}
function copyToClipboard(btn){
  const text = btn.parentElement.textContent.replace('Copy', '').trim();
  navigator.clipboard.writeText(text).then(() => {
    btn.textContent = '✅ Copied';
    setTimeout(() => btn.textContent = 'Copy', 2000);
  });
}
refreshStatus();
setInterval(refreshStatus, 15000);
</script>
</body>
</html>
"""

# ═══════════════════════════════════════════════════════════════
# ROUTES (HTML)
# ═══════════════════════════════════════════════════════════════
@app.route('/setup')
def setup_page():
    return Response(SETUP_HTML, mimetype='text/html')

@app.route('/')
def index():
    return Response(DASHBOARD_HTML, mimetype='text/html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))