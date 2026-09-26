from flask import Flask, jsonify, request, send_from_directory, session, redirect, url_for
import json, os, secrets
from datetime import datetime, timedelta
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

# ── LEGACY BOOTSTRAP CREDENTIALS ──────────────────────
# Used only once, to create the first admin account in users.json
# if that file doesn't exist yet. After that, all logins go
# through users.json — change your password from the Users tab.
BOOTSTRAP_USERNAME = "Sky Fr8"
BOOTSTRAP_PASSWORD = "Skyfr8@2022"

DATA_FILE    = os.path.join(os.path.dirname(__file__), "data", "shipments.json")
FLIGHTS_FILE = os.path.join(os.path.dirname(__file__), "data", "flights.json")
BIN_FILE     = os.path.join(os.path.dirname(__file__), "data", "bin.json")
USERS_FILE   = os.path.join(os.path.dirname(__file__), "data", "users.json")
ACTLOG_FILE  = os.path.join(os.path.dirname(__file__), "data", "activity_log.json")
ACTLOG_MAX_ENTRIES = 5000

def load_json(path):
    if not os.path.exists(path):
        return []
    with open(path, "r") as f:
        return json.load(f)

def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, indent=2)

def is_logged_in():
    return session.get('logged_in') is True

def is_admin():
    return session.get('logged_in') is True and session.get('role') == 'admin'

def load_shipments(): return load_json(DATA_FILE)
def save_shipments(d): save_json(DATA_FILE, d)
def load_flights(): return load_json(FLIGHTS_FILE)
def save_flights(d): save_json(FLIGHTS_FILE, d)
def load_bin(): return load_json(BIN_FILE)
def save_bin(d): save_json(BIN_FILE, d)

def load_users(): return load_json(USERS_FILE)
def save_users(d): save_json(USERS_FILE, d)

def load_activity_log(): return load_json(ACTLOG_FILE)
def save_activity_log(d): save_json(ACTLOG_FILE, d)

def log_activity(action, detail=""):
    """Append an entry to the activity log. Uses the logged-in
    session's username, or 'system' if there isn't one."""
    entries = load_activity_log()
    entries.append({
        "timestamp": datetime.now().isoformat(),
        "username": session.get("username", "system"),
        "action": action,
        "detail": detail
    })
    if len(entries) > ACTLOG_MAX_ENTRIES:
        entries = entries[-ACTLOG_MAX_ENTRIES:]
    save_activity_log(entries)

def ensure_users_bootstrapped():
    """If users.json doesn't exist yet, create it with a single admin
    account using the old hardcoded credentials, so existing logins
    keep working after this update."""
    if os.path.exists(USERS_FILE):
        return
    bootstrap_admin = {
        "username": BOOTSTRAP_USERNAME,
        "password_hash": generate_password_hash(BOOTSTRAP_PASSWORD),
        "role": "admin",
        "active": True,
        "created": datetime.now().isoformat()
    }
    save_users([bootstrap_admin])

def find_user(username):
    return next((u for u in load_users() if u["username"].lower() == username.lower()), None)

def public_user(u):
    return {"username": u["username"], "role": u["role"], "active": u.get("active", True), "created": u.get("created")}

ensure_users_bootstrapped()

@app.route("/")
def index():
    if not is_logged_in():
        return redirect('/login')
    return send_from_directory("templates", "index.html")

@app.route("/login", methods=["GET"])
def login_page():
    return send_from_directory("templates", "login.html")

@app.route("/login", methods=["POST"])
def login():
    data = request.get_json()
    username = (data.get("username") or "").strip()
    password = data.get("password") or ""
    user = find_user(username)
    if not user or not check_password_hash(user["password_hash"], password):
        return jsonify({"error": "Invalid username or password."}), 401
    if not user.get("active", True):
        return jsonify({"error": "This account has been deactivated."}), 403
    session['logged_in'] = True
    session['username']  = user["username"]
    session['role']      = user["role"]
    log_activity("LOGIN")
    return jsonify({"success": True})

@app.route("/logout")
def logout():
    if is_logged_in():
        log_activity("LOGOUT")
    session.clear()
    return redirect('/login')

@app.route("/api/session", methods=["GET"])
def get_session():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    return jsonify({"username": session.get("username"), "role": session.get("role")})

@app.route("/track")
def track():
    return send_from_directory("templates", "track.html")

# ── SHIPMENTS ──────────────────────────────────────────
@app.route("/api/shipments", methods=["GET"])
def get_shipments():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    return jsonify(load_shipments())

@app.route("/api/shipments", methods=["POST"])
def add_shipment():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    shipments = load_shipments()
    new = request.get_json()
    awb = new.get("awb", "").strip()
    if not awb:
        return jsonify({"error": "AWB number is required."}), 400
    if any(s["awb"].lower() == awb.lower() for s in shipments):
        return jsonify({"error": "AWB already exists."}), 409
    flight = new.get("flight", "")
    date   = new.get("eta", "")
    match  = next((f for f in load_flights() if f["flight"].lower()==flight.lower() and f["date"]==date), None)
    if match and match.get("asycuda") and not new.get("asycuda"):
        new["asycuda"] = match["asycuda"]
    shipments.append(new)
    save_shipments(shipments)
    log_activity("ADD_SHIPMENT", f"AWB {awb}")
    return jsonify({"success": True, "awb": awb}), 201

@app.route("/api/shipments/<awb>", methods=["PUT"])
def update_shipment(awb):
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    shipments = load_shipments()
    data = request.get_json()
    for i, s in enumerate(shipments):
        if s["awb"].lower() == awb.lower():
            shipments[i].update(data)
            save_shipments(shipments)
            log_activity("EDIT_SHIPMENT", f"AWB {awb}")
            return jsonify({"success": True})
    return jsonify({"error": "AWB not found."}), 404

@app.route("/api/shipments/<awb>/bin", methods=["POST"])
def move_to_bin(awb):
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    shipments = load_shipments()
    record = next((s for s in shipments if s["awb"].lower()==awb.lower()), None)
    if not record:
        return jsonify({"error": "AWB not found."}), 404
    bin_items = load_bin()
    record["deleted_at"] = datetime.now().isoformat()
    bin_items.append(record)
    save_bin(bin_items)
    save_shipments([s for s in shipments if s["awb"].lower()!=awb.lower()])
    log_activity("DELETE_SHIPMENT", f"AWB {awb}")
    return jsonify({"success": True})

@app.route("/api/shipments/bulk-asycuda", methods=["POST"])
def bulk_asycuda():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    data      = request.get_json()
    flight    = data.get("flight", "").strip()
    date      = data.get("date", "").strip()
    asycuda   = data.get("asycuda", "").strip()
    shipments = load_shipments()
    updated   = 0
    for s in shipments:
        if s.get("flight","").lower()==flight.lower() and s.get("eta","")==date:
            s["asycuda"] = asycuda
            updated += 1
    save_shipments(shipments)
    log_activity("BULK_ASYCUDA", f"Flight {flight} {date}: {updated} AWB(s) → {asycuda}")
    return jsonify({"success": True, "updated": updated})

# ── RECYCLE BIN ──────────────────────────────────────
@app.route("/api/bin", methods=["GET"])
def get_bin():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    bin_items = load_bin()
    cutoff = datetime.now() - timedelta(days=30)
    bin_items = [b for b in bin_items if datetime.fromisoformat(b.get("deleted_at","2000-01-01")) > cutoff]
    save_bin(bin_items)
    return jsonify(bin_items)

@app.route("/api/bin/<awb>/restore", methods=["POST"])
def restore_from_bin(awb):
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    bin_items = load_bin()
    record = next((b for b in bin_items if b["awb"].lower()==awb.lower()), None)
    if not record:
        return jsonify({"error": "Record not found in bin."}), 404
    record.pop("deleted_at", None)
    shipments = load_shipments()
    shipments.append(record)
    save_shipments(shipments)
    save_bin([b for b in bin_items if b["awb"].lower()!=awb.lower()])
    log_activity("RESTORE_SHIPMENT", f"AWB {awb}")
    return jsonify({"success": True})

@app.route("/api/bin/<awb>", methods=["DELETE"])
def permanent_delete(awb):
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    bin_items = load_bin()
    new_bin = [b for b in bin_items if b["awb"].lower()!=awb.lower()]
    if len(new_bin)==len(bin_items):
        return jsonify({"error": "Record not found in bin."}), 404
    save_bin(new_bin)
    log_activity("PERM_DELETE", f"AWB {awb}")
    return jsonify({"success": True})

@app.route("/api/bin", methods=["DELETE"])
def empty_bin():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    save_bin([])
    log_activity("EMPTY_BIN")
    return jsonify({"success": True})

# ── FLIGHTS ───────────────────────────────────────────
@app.route("/api/flights", methods=["GET"])
def get_flights():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    return jsonify(load_flights())

@app.route("/api/flights", methods=["POST"])
def save_flight():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    flights = load_flights()
    new = request.get_json()
    flight = new.get("flight","").strip()
    date   = new.get("date","").strip()
    if not flight or not date:
        return jsonify({"error": "Flight number and date are required."}), 400
    for i, f in enumerate(flights):
        if f["flight"].lower()==flight.lower() and f["date"]==date:
            flights[i].update(new)
            save_flights(flights)
            log_activity("EDIT_FLIGHT", f"{flight} ({date})")
            return jsonify({"success": True, "updated": True})
    flights.append(new)
    save_flights(flights)
    log_activity("ADD_FLIGHT", f"{flight} ({date})")
    return jsonify({"success": True, "updated": False}), 201

@app.route("/api/flights/<flight>/<date>", methods=["DELETE"])
def delete_flight(flight, date):
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    flights = load_flights()
    new_list = [f for f in flights if not (f["flight"].lower()==flight.lower() and f["date"]==date)]
    if len(new_list)==len(flights):
        return jsonify({"error": "Flight not found."}), 404
    save_flights(new_list)
    log_activity("DELETE_FLIGHT", f"{flight} ({date})")
    return jsonify({"success": True})

# ── USERS (admin only) ────────────────────────────────
@app.route("/api/users", methods=["GET"])
def get_users():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    if not is_admin(): return jsonify({"error": "Admin access required."}), 403
    return jsonify([public_user(u) for u in load_users()])

@app.route("/api/users", methods=["POST"])
def create_user():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    if not is_admin(): return jsonify({"error": "Admin access required."}), 403
    data     = request.get_json()
    username = (data.get("username") or "").strip()
    password = data.get("password") or ""
    role     = data.get("role") or "staff"
    if not username or not password:
        return jsonify({"error": "Username and password are required."}), 400
    if role not in ("admin", "staff"):
        return jsonify({"error": "Role must be 'admin' or 'staff'."}), 400
    if find_user(username):
        return jsonify({"error": "A user with that username already exists."}), 409
    users = load_users()
    users.append({
        "username": username,
        "password_hash": generate_password_hash(password),
        "role": role,
        "active": True,
        "created": datetime.now().isoformat()
    })
    save_users(users)
    log_activity("CREATE_USER", f"{username} ({role})")
    return jsonify({"success": True}), 201

@app.route("/api/users/<username>", methods=["PUT"])
def update_user(username):
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    if not is_admin(): return jsonify({"error": "Admin access required."}), 403
    data  = request.get_json()
    users = load_users()
    idx   = next((i for i,u in enumerate(users) if u["username"].lower()==username.lower()), None)
    if idx is None:
        return jsonify({"error": "User not found."}), 404

    target = users[idx]
    new_role   = data.get("role", target["role"])
    new_active = data.get("active", target.get("active", True))

    # Don't allow removing the last active admin's admin rights or access
    admins = [u for u in users if u["role"]=="admin" and u.get("active",True)]
    is_only_admin = target["role"]=="admin" and target.get("active",True) and len(admins)<=1
    if is_only_admin and (new_role != "admin" or not new_active):
        return jsonify({"error": "Can't remove the last active admin."}), 400

    target["role"]   = new_role
    target["active"] = new_active
    if data.get("password"):
        target["password_hash"] = generate_password_hash(data["password"])
    save_users(users)
    log_activity("EDIT_USER", f"{username} → role={new_role}, active={new_active}")
    return jsonify({"success": True})

@app.route("/api/users/<username>", methods=["DELETE"])
def delete_user(username):
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    if not is_admin(): return jsonify({"error": "Admin access required."}), 403
    if username.lower() == (session.get("username") or "").lower():
        return jsonify({"error": "You can't delete your own account while logged in."}), 400
    users  = load_users()
    target = next((u for u in users if u["username"].lower()==username.lower()), None)
    if not target:
        return jsonify({"error": "User not found."}), 404
    admins = [u for u in users if u["role"]=="admin" and u.get("active",True)]
    if target["role"]=="admin" and target.get("active",True) and len(admins)<=1:
        return jsonify({"error": "Can't delete the last active admin."}), 400
    save_users([u for u in users if u["username"].lower()!=username.lower()])
    log_activity("DELETE_USER", username)
    return jsonify({"success": True})

# ── ACTIVITY LOG (admin only) ─────────────────────────
@app.route("/api/activity-log", methods=["GET"])
def get_activity_log():
    if not is_logged_in(): return jsonify({"error": "Unauthorised"}), 401
    if not is_admin(): return jsonify({"error": "Admin access required."}), 403
    limit = request.args.get("limit", default=500, type=int)
    entries = load_activity_log()
    entries = list(reversed(entries))  # newest first
    if limit:
        entries = entries[:limit]
    return jsonify(entries)

if __name__ == "__main__":
    ensure_users_bootstrapped()
    print("\n" + "="*52)
    print("  AirCargo WMS is running!")
    print("="*52)
    print("  Staff portal:    http://localhost:5000")
    print("  Customer track:  http://localhost:5000/track")
    print("")
    print("  To access from other devices on the network,")
    print("  replace 'localhost' with this computer's IP.")
    print("  (Run 'ipconfig' in Command Prompt to find it)")
    print("="*52 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False)
