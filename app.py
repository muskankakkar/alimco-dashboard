import sys, os, io, json, re, urllib.request, urllib.error
from datetime import datetime

# Windows UTF-8 fix
if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

from flask import Flask, jsonify, send_file
from flask_cors import CORS
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

app  = Flask(__name__)
CORS(app)

FILE_NAME  = "exports dashboard.xlsx"
SHEET_NAME = "Leads Data"
HEADER_ROW = 2

GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
GEMINI_MODEL   = "gemini-2.0-flash"

COUNTRY_COORDS = {
    "Zambia":{"lat":-13.13,"lng":27.85},"Nigeria":{"lat":9.08,"lng":8.68},
    "Bangladesh":{"lat":23.68,"lng":90.35},"Togo":{"lat":8.62,"lng":0.82},
    "Iraq":{"lat":33.22,"lng":43.68},"Mongolia":{"lat":46.86,"lng":103.85},
    "Turkmenistan":{"lat":38.97,"lng":59.56},"South Africa":{"lat":-30.56,"lng":22.94},
    "Tanzania":{"lat":-6.37,"lng":34.89},"India":{"lat":20.59,"lng":78.96},
    "UAE":{"lat":23.42,"lng":53.85},"Egypt":{"lat":26.82,"lng":30.80},
    "Kenya":{"lat":-0.02,"lng":37.91},"Ethiopia":{"lat":9.15,"lng":40.49},
    "Ghana":{"lat":7.95,"lng":-1.02},"Uganda":{"lat":1.37,"lng":32.29},
    "Nepal":{"lat":28.39,"lng":84.12},"Sri Lanka":{"lat":7.87,"lng":80.77},
    "Philippines":{"lat":12.88,"lng":121.77},"Indonesia":{"lat":-0.79,"lng":113.92},
    "Malaysia":{"lat":4.21,"lng":101.98},"Thailand":{"lat":15.87,"lng":100.99},
    "Vietnam":{"lat":14.06,"lng":108.28},"Pakistan":{"lat":30.38,"lng":69.35},
    "Saudi Arabia":{"lat":23.89,"lng":45.08},"Oman":{"lat":21.51,"lng":55.92},
    "Jordan":{"lat":30.59,"lng":36.24},"Morocco":{"lat":31.79,"lng":-7.09},
    "Senegal":{"lat":14.50,"lng":-14.45},"Cameroon":{"lat":3.85,"lng":11.50},
    "Mozambique":{"lat":-18.67,"lng":35.53},"Zimbabwe":{"lat":-19.02,"lng":29.15},
    "Rwanda":{"lat":-1.94,"lng":29.87},"Angola":{"lat":-11.20,"lng":17.87},
    "Zambia":{"lat":-13.13,"lng":27.85},"Malawi":{"lat":-13.25,"lng":34.30},
    "Botswana":{"lat":-22.33,"lng":24.68},"Namibia":{"lat":-22.96,"lng":18.49},
    "Germany":{"lat":51.17,"lng":10.45},"France":{"lat":46.23,"lng":2.21},
    "Japan":{"lat":36.20,"lng":138.25},"Singapore":{"lat":1.35,"lng":103.82},
    "Australia":{"lat":-25.27,"lng":133.78},
}

DEMO_DATA = [
    {"sno":1,"company":"Afrimed Pharmaceutical Ltd.","contact_name":"Mr. Hyder Khan","last_contacted":"2026-03-01","order_number":"100 BTE TD 0E 17036","order_type":"Product Supply","value_inr":680256,"value_usd":8503.2,"lead_source":"Email","lead_status":"Interested","deal_stage":"Negotiating","email":"afrimedpharmaceuticals@gmail.com","country":"Zambia","city":"LUSAKA","pin_code":"10101","contact_number":"+260 979 328887","notes":"Interested in Hearing Aids","last_updated":"2026-05-19"},
    {"sno":2,"company":"SJS Life Sciences Ltd.","contact_name":"Jitendra Pandey","last_contacted":"2026-05-12","order_number":"101 BTE TD 0E 17036","order_type":"Product Supply","value_inr":680256,"value_usd":8503.2,"lead_source":"Email","lead_status":"Negotiating","deal_stage":"Negotiating","email":"jkpandey3s@gmail.com","country":"Nigeria","city":"LAGOS","pin_code":"100271","contact_number":"+2348128444444","notes":"Interested in Hearing Aids","last_updated":"2026-05-19"},
    {"sno":3,"company":"Medikit Corporation","contact_name":"Md. Mamunor Rashid","last_contacted":"2026-04-07","order_number":"MULTIPLE ASSISTIVE DEVICES","order_type":"Product Supply","value_inr":62458200,"value_usd":780727.5,"lead_source":"Email","lead_status":"Not Interested","deal_stage":"Closed - Lost","email":"s.johnson@texmart.us","country":"Bangladesh","city":"Dhaka","pin_code":"1206","contact_number":"+2348128444444","notes":"Responded to email","last_updated":"2026-05-19"},
    {"sno":4,"company":"Kamsitoc Global Trading Sarl","contact_name":"Comrade Kwaku Jojo","last_contacted":"2026-01-05","order_number":"MULTIPLE ASSISTIVE DEVICES","order_type":"Product Supply","value_inr":45000,"value_usd":562.5,"lead_source":"Email","lead_status":"Interested","deal_stage":"Negotiating","email":"kwakujojo.7@gmail.com","country":"Togo","city":"Lome","pin_code":"42041","contact_number":"+22898191726","notes":"export pricing required","last_updated":"2026-05-19"},
    {"sno":5,"company":"JMB Entity","contact_name":"Dr. Rasar Rahim","last_contacted":"","order_number":"","order_type":"Product Supply","value_inr":0,"value_usd":0,"lead_source":"ITEC","lead_status":"Reply to be Given","deal_stage":"Proposal","email":"rasarrebwar@gmail.com","country":"Iraq","city":"","pin_code":"","contact_number":"","notes":"export pricing required","last_updated":""},
    {"sno":6,"company":"Mongolian Burn Association","contact_name":"Dr. Batorgil Enkhbayar","last_contacted":"2026-02-26","order_number":"phased collaboration","order_type":"Phased Collaboration","value_inr":0,"value_usd":0,"lead_source":"Email","lead_status":"Reply to be Given","deal_stage":"Proposal","email":"btrglcx98@gmail.com","country":"Mongolia","city":"","pin_code":"","contact_number":"","notes":"email to be replied","last_updated":""},
    {"sno":7,"company":"JMB Pharmaceuticals Pvt. Ltd.","contact_name":"Krishna Mehta","last_contacted":"2026-03-12","order_number":"orthotics and prosthetics","order_type":"Product Supply","value_inr":0,"value_usd":0,"lead_source":"Email","lead_status":"Reply to be Given","deal_stage":"Proposal","email":"wessely@jmbpharma.com","country":"Turkmenistan","city":"","pin_code":"","contact_number":"+91-9082004363","notes":"email to be replied","last_updated":""},
    {"sno":8,"company":"Tripalogix Pty Ltd","contact_name":"Karabo Motsamai Vuso","last_contacted":"2026-02-23","order_number":"MULTIPLE ASSISTIVE DEVICES","order_type":"Product Supply","value_inr":0,"value_usd":0,"lead_source":"Email","lead_status":"Reply to be Given","deal_stage":"Proposal","email":"tripalogix@gmail.com","country":"South Africa","city":"","pin_code":"","contact_number":"0615414005","notes":"email to be replied","last_updated":""},
    {"sno":9,"company":"Kamal Group","contact_name":"Mr. Sameer Santosh","last_contacted":"2026-05-18","order_number":"","order_type":"Phased Collaboration","value_inr":0,"value_usd":0,"lead_source":"Visit","lead_status":"Interested","deal_stage":"Proposal","email":"md@kamal-group.co.tz","country":"Tanzania","city":"Dar es Salaam","pin_code":"10392","contact_number":"+255 767 624 571","notes":"reply awaited","last_updated":""},
    {"sno":10,"company":"METL","contact_name":"Mr. Gulam Dewji","last_contacted":"2026-05-18","order_number":"","order_type":"Phased Collaboration","value_inr":0,"value_usd":0,"lead_source":"Visit","lead_status":"Interested","deal_stage":"Proposal","email":"gd@metl.net","country":"Tanzania","city":"Dar es Salaam","pin_code":"20660","contact_number":"+255 754 600 000","notes":"reply awaited","last_updated":""},
]

AI_INSIGHTS = ""
AI_ALERTS   = []
ALL_LEADS   = []
AI_READY    = False
FILE_MTIME  = ""


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)


def load_excel():
    global ALL_LEADS, FILE_MTIME
    try:
        path = FILE_NAME
        if not os.path.exists(path):
            for alt in ["exports_dashbaord.xlsx", "exports_dashboard.xlsx"]:
                if os.path.exists(alt):
                    path = alt
                    break
        if not os.path.exists(path):
            log(f"WARNING: {FILE_NAME} not found. Using demo data.")
            ALL_LEADS = list(DEMO_DATA)
            FILE_MTIME = "Demo data"
            return
        df = pd.read_excel(path, sheet_name=SHEET_NAME, engine="openpyxl", header=HEADER_ROW)
        df.columns = df.columns.str.strip()
        df = df.dropna(subset=["SNO."]).copy()
        df["value_inr"] = pd.to_numeric(df.get("Order Value (₹ INR)", 0), errors="coerce").fillna(0)
        df["value_usd"] = pd.to_numeric(df.get("Order Value ($ USD)", 0), errors="coerce").fillna(0)
        str_cols = {
            "SNO.":"sno","Company Name":"company","Contact Person":"contact_name",
            "Last Contacted":"last_contacted","Order #":"order_number",
            "Order Type":"order_type","Lead Source":"lead_source",
            "Lead Status":"lead_status","Deal Stage":"deal_stage",
            "Email Address":"email","Country":"country","City":"city",
            "Pin Code":"pin_code","Contact Number":"contact_number",
            "Notes":"notes","Last Updated":"last_updated",
        }
        leads = []
        for _, row in df.iterrows():
            lead = {"value_inr": float(row["value_inr"]), "value_usd": float(row["value_usd"])}
            for col, key in str_cols.items():
                val = row.get(col, "")
                if pd.isna(val) or str(val).strip() in ("NaT","nan","None"):
                    lead[key] = ""
                else:
                    lead[key] = str(val).strip().title() if key == "country" else str(val).strip()
            try:
                lead["sno"] = int(float(lead["sno"]))
            except Exception:
                lead["sno"] = 0
            leads.append(lead)
        ALL_LEADS = leads
        FILE_MTIME = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
        log(f"Loaded {len(ALL_LEADS)} leads from {path}")
    except Exception as e:
        log(f"ERROR loading Excel: {e}. Using demo data.")
        ALL_LEADS = list(DEMO_DATA)
        FILE_MTIME = "Demo data (error)"


def gemini_chat(system_msg, user_msg, max_tokens=900):
    api_key = os.getenv("GEMINI_API_KEY", "")
    if not api_key:
        raise ValueError("GEMINI_API_KEY not set in .env")
    url = GEMINI_API_URL.format(model=GEMINI_MODEL, key=api_key)
    combined = f"{system_msg}\n\n{user_msg}"
    payload = json.dumps({
        "contents": [{"parts": [{"text": combined}]}],
        "generationConfig": {"maxOutputTokens": max_tokens, "temperature": 0.4},
    }).encode("utf-8")
    req = urllib.request.Request(url, data=payload,
        headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return data["candidates"][0]["content"]["parts"][0]["text"].strip()


def parse_alerts(raw):
    try:
        cleaned = re.sub(r"```(?:json)?|```", "", raw).strip()
        alerts = json.loads(cleaned)
        if not isinstance(alerts, list):
            alerts = []
        required = {"company","country","city","urgency","alert","action","due_in_days"}
        alerts = [a for a in alerts if isinstance(a, dict) and required.issubset(a.keys())]
        order = {"CRITICAL":0,"HIGH":1,"MEDIUM":2}
        alerts.sort(key=lambda x: order.get(x.get("urgency","MEDIUM"), 2))
        return alerts
    except Exception as e:
        log(f"Alert parse error: {e}")
        return []


def run_ai(leads):
    global AI_INSIGHTS, AI_ALERTS, AI_READY
    api_key = os.getenv("GEMINI_API_KEY", "")
    if not api_key:
        AI_INSIGHTS = "AI unavailable — add GEMINI_API_KEY to .env (free at aistudio.google.com)"
        AI_ALERTS = []
        AI_READY = False
        log("GEMINI_API_KEY not set. Skipping AI.")
        return

    condensed = [{k:v for k,v in lead.items() if k in
        ("company","country","city","lead_status","deal_stage","order_type",
         "value_inr","value_usd","lead_source","last_contacted","notes")}
        for lead in leads]

    briefing_system = (
        "You are a senior export strategy consultant. Write concise data-driven "
        "executive briefings. Always cite specific numbers. Max 320 words. "
        "Use bold headings and numbered points."
    )
    briefing_user = (
        f"Analyse this export pipeline and write a briefing.\n\nDATA:\n"
        f"{json.dumps(condensed, ensure_ascii=False)}\n\n"
        "Write these 6 sections:\n"
        "**1. PIPELINE OVERVIEW** - total leads, INR/USD value, countries\n"
        "**2. TOP 3 MARKETS** - by order value with INR amounts\n"
        "**3. DEAL STAGE HEALTH** - where leads are clustering\n"
        "**4. LEAD SOURCE PERFORMANCE** - which sources generate most value\n"
        "**5. STRATEGIC RECOMMENDATIONS** - two specific actions\n"
        "**6. RISK WATCH** - one concrete risk in the data"
    )
    try:
        log("AI Call A — Executive Briefing via Gemini...")
        AI_INSIGHTS = gemini_chat(briefing_system, briefing_user, max_tokens=900)
        log("AI Call A complete.")
    except Exception as e:
        log(f"AI Call A error: {e}")
        AI_INSIGHTS = f"Briefing unavailable: {e}"

    notes_data = [{"company":l.get("company",""),"country":l.get("country",""),
                   "city":l.get("city",""),"notes":l.get("notes","")}
                  for l in leads if l.get("notes","").strip()]

    alerts_system = (
        "You are an export operations analyst. "
        "CRITICAL: Respond with ONLY a valid JSON array, no other text, no code fences. "
        "If nothing urgent, return: []"
    )
    alerts_user = (
        f"Scan these lead notes and return urgent items as JSON array.\n"
        f"Each object needs: company, country, city, urgency (CRITICAL/HIGH/MEDIUM), "
        f"alert (1 sentence), action (1 sentence), due_in_days (int or null).\n\n"
        f"LEADS:\n{json.dumps(notes_data, ensure_ascii=False)}\n\n"
        f"Return ONLY the JSON array."
    )
    try:
        log("AI Call B — Urgent Alerts via Gemini...")
        raw = gemini_chat(alerts_system, alerts_user, max_tokens=1000)
        AI_ALERTS = parse_alerts(raw)
        log(f"AI Call B complete. {len(AI_ALERTS)} alerts found.")
        AI_READY = True
    except Exception as e:
        log(f"AI Call B error: {e}")
        AI_ALERTS = []
        AI_READY = False


def build_meta():
    total_inr = sum(l.get("value_inr", 0) for l in ALL_LEADS)
    total_usd = sum(l.get("value_usd", 0) for l in ALL_LEADS)
    countries  = len(set(l.get("country","") for l in ALL_LEADS if l.get("country")))
    return {"file":FILE_NAME,"sheet":SHEET_NAME,"last_modified":FILE_MTIME,
            "total_leads":len(ALL_LEADS),"total_pipeline_inr":round(total_inr,2),
            "total_pipeline_usd":round(total_usd,2),"countries_count":countries,
            "ai_ready":AI_READY}


def init_data():
    load_excel()
    run_ai(ALL_LEADS)


@app.route("/")
def index():
    return send_file("dashboard.html")


@app.route("/api/leads")
def api_leads():
    return jsonify({"leads":ALL_LEADS,"ai_insights":AI_INSIGHTS,
                    "alerts":AI_ALERTS,"meta":build_meta()})


@app.route("/api/alerts")
def api_alerts():
    return jsonify({"alerts":AI_ALERTS})


@app.route("/api/map")
def api_map():
    agg = {}
    for lead in ALL_LEADS:
        c = lead.get("country","").strip()
        if not c:
            continue
        if c not in agg:
            agg[c] = {"name":c,"lead_count":0,"pipeline_inr":0.0,
                      "pipeline_usd":0.0,"hot_leads":0,"top_city":"","top_company":""}
        a = agg[c]
        a["lead_count"] += 1
        a["pipeline_inr"] += lead.get("value_inr",0)
        a["pipeline_usd"] += lead.get("value_usd",0)
        if any(x in lead.get("lead_status","").lower() for x in ("hot","interest","negot")):
            a["hot_leads"] += 1
        if lead.get("city") and not a["top_city"]:
            a["top_city"] = lead["city"]
        if lead.get("company") and not a["top_company"]:
            a["top_company"] = lead["company"]
    result = []
    for name, a in agg.items():
        coords = COUNTRY_COORDS.get(name)
        if coords:
            a["lat"] = coords["lat"]
            a["lng"] = coords["lng"]
            a["pipeline_inr"] = round(a["pipeline_inr"], 2)
            a["pipeline_usd"] = round(a["pipeline_usd"], 2)
            result.append(a)
    return jsonify({"countries":result})


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    init_data()
    return jsonify({"leads":ALL_LEADS,"ai_insights":AI_INSIGHTS,
                    "alerts":AI_ALERTS,"meta":build_meta()})


@app.route("/api/health")
def api_health():
    return jsonify({"status":"ok","leads_loaded":len(ALL_LEADS),
                    "ai_ready":AI_READY,"sheet":SHEET_NAME,
                    "ai_provider":"Google Gemini","model":GEMINI_MODEL})


# ── STARTUP ──────────────────────────────────────────────────────────────────
log("Starting ALIMCO Export Intelligence Dashboard...")
init_data()
log("Startup complete. Flask is ready.")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    log(f"Running on http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
