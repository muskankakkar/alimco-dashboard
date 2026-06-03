"""
ALIMCO Export Dashboard — Google Sheets Live Sync
==================================================
Set SHEET_URL environment variable in Railway to your
Google Sheets published CSV link.
Dashboard auto-refreshes every 30 seconds from the sheet.
"""

import os, io, csv, time
from datetime import datetime
from flask import Flask, jsonify, send_from_directory
from flask_cors import CORS
import urllib.request as urlreq

app  = Flask(__name__, static_folder=".")
CORS(app)

SHEET_URL = os.environ.get("SHEET_URL", "")
CACHE_TTL = 30
_cache    = {"data": None, "ts": 0}

COORDS = {
    "zambia":       (-15.4166,  28.2833),
    "nigeria":      (  6.5244,   3.3792),
    "bangladesh":   ( 23.8103,  90.4125),
    "togo":         (  6.1374,   1.2123),
    "iraq":         ( 33.3128,  44.3615),
    "mongolia":     ( 47.8864, 106.9057),
    "turkmenistan": ( 37.9601,  58.3261),
    "south africa": (-26.2041,  28.0473),
    "tanzania":     ( -6.7924,  39.2083),
    "kenya":        ( -1.2921,  36.8219),
    "ethiopia":     (  9.0054,  38.7636),
    "egypt":        ( 30.0444,  31.2357),
    "ghana":        (  5.6037,  -0.1870),
    "uae":          ( 25.2048,  55.2708),
    "saudi arabia": ( 24.7136,  46.6753),
    "thailand":     ( 13.7563, 100.5018),
    "philippines":  ( 14.5995, 120.9842),
    "vietnam":      ( 21.0278, 105.8342),
    "nepal":        ( 27.7172,  85.3240),
    "sri lanka":    (  6.9271,  79.8612),
    "iran":         ( 35.6892,  51.3890),
}

FALLBACK = [
    {"sno":1,"company":"Afrimed Pharmaceutical Ltd.","contact":"Mr. Hyder Khan","country":"Zambia","status":"Interested","stage":"Negotiating","type":"Product Supply","source":"Email","value_inr":680256,"value_usd":8503.2,"notes":"Interested in Hearing Aids","lat":-15.4166,"lng":28.2833},
    {"sno":2,"company":"SJS Life Sciences Ltd.","contact":"Jitendra Pandey","country":"Nigeria","status":"Negotiating","stage":"Negotiating","type":"Product Supply","source":"Email","value_inr":680256,"value_usd":8503.2,"notes":"Interested in Hearing Aids","lat":6.5244,"lng":3.3792},
    {"sno":3,"company":"Medikit Corporation","contact":"Md. Mamunor Rashid","country":"Bangladesh","status":"Not Interested","stage":"Closed - Lost","type":"Product Supply","source":"Email","value_inr":62458200,"value_usd":780727.5,"notes":"Responded to email","lat":23.8103,"lng":90.4125},
    {"sno":4,"company":"Kamsitoc Global Trading Sarl","contact":"Comrade Kwaku Jojo","country":"Togo","status":"Interested","stage":"Negotiating","type":"Product Supply","source":"Email","value_inr":45000,"value_usd":562.5,"notes":"Export pricing required","lat":6.1374,"lng":1.2123},
    {"sno":5,"company":"(Unnamed Lead)","contact":"Dr. Rasar Rahim","country":"Iraq","status":"Reply to be Given","stage":"Proposal","type":"Product Supply","source":"ITEC","value_inr":0,"value_usd":0,"notes":"Export pricing required","lat":33.3128,"lng":44.3615},
    {"sno":6,"company":"Mongolian Burn Association","contact":"Dr. Batorgil Enkhbayar","country":"Mongolia","status":"Reply to be Given","stage":"Proposal","type":"Phased Collaboration","source":"Email","value_inr":0,"value_usd":0,"notes":"Email to be replied","lat":47.8864,"lng":106.9057},
    {"sno":7,"company":"JMB Pharmaceuticals Pvt. Ltd.","contact":"Krishna Mehta","country":"Turkmenistan","status":"Reply to be Given","stage":"Proposal","type":"Product Supply","source":"Email","value_inr":0,"value_usd":0,"notes":"Email to be replied","lat":37.9601,"lng":58.3261},
    {"sno":8,"company":"Tripalogix Pty Ltd","contact":"Karabo Motsamai Vuso","country":"South Africa","status":"Reply to be Given","stage":"Proposal","type":"Product Supply","source":"Email","value_inr":0,"value_usd":0,"notes":"Email to be replied","lat":-26.2041,"lng":28.0473},
    {"sno":9,"company":"Kamal Group","contact":"Mr. Sameer Santosh","country":"Tanzania","status":"Interested","stage":"Proposal","type":"Phased Collaboration","source":"Visit","value_inr":0,"value_usd":0,"notes":"Reply awaited","lat":-6.7924,"lng":39.2083},
    {"sno":10,"company":"METL","contact":"Mr. Gulam Dewji","country":"Tanzania","status":"Interested","stage":"Proposal","type":"Phased Collaboration","source":"Visit","value_inr":0,"value_usd":0,"notes":"Reply awaited","lat":-6.7924,"lng":39.2083},
]

def fetch_sheet():
    global _cache
    if not SHEET_URL:
        return None, "SHEET_URL not set"
    if _cache["data"] and time.time() - _cache["ts"] < CACHE_TTL:
        return _cache["data"], None
    try:
        req = urlreq.Request(SHEET_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urlreq.urlopen(req, timeout=10) as r:
            raw = r.read().decode("utf-8")
    except Exception as e:
        return None, str(e)

    rows = list(csv.reader(io.StringIO(raw)))
    if len(rows) < 2:
        return None, "Sheet empty"

    hi = next((i for i,row in enumerate(rows)
               if any("company" in c.lower() for c in row)
               and any("country" in c.lower() for c in row)), None)
    if hi is None:
        return None, "Header not found"

    hdrs = [c.lower().strip() for c in rows[hi]]

    def ci(*kws):
        for kw in kws:
            for i,h in enumerate(hdrs):
                if kw in h: return i
        return -1

    idx = dict(
        company=ci("company name","company"), contact=ci("contact person","contact"),
        country=ci("country"), status=ci("lead status","status"),
        stage=ci("deal stage","stage"), type=ci("order type","type"),
        source=ci("lead source","source"), val_inr=ci("₹ inr","inr","value (₹","order value"),
        val_usd=ci("$ usd","usd"), notes=ci("notes"), date=ci("last contacted","date"), city=ci("city"),
    )

    def g(row,k):
        i=idx.get(k,-1)
        return row[i].strip() if 0<=i<len(row) else ""
    def gf(row,k):
        try: return float(g(row,k).replace(",","").replace("₹","").replace("$",""))
        except: return 0.0

    leads=[]
    for sno,row in enumerate(rows[hi+1:],1):
        if len(row)<3: continue
        co=g(row,"company"); ct=g(row,"country")
        if not co and not ct: continue
        coord=COORDS.get(ct.lower().strip(),(None,None))
        leads.append(dict(sno=sno,company=co or "(Unnamed)",contact=g(row,"contact"),
            country=ct.title() if ct else "—",city=g(row,"city"),status=g(row,"status"),
            stage=g(row,"stage"),type=g(row,"type"),source=g(row,"source"),
            value_inr=gf(row,"val_inr"),value_usd=gf(row,"val_usd"),
            notes=g(row,"notes"),date=g(row,"date"),lat=coord[0],lng=coord[1]))

    result={"leads":leads,"meta":{"total":len(leads),"source":"google_sheets",
        "last_modified":datetime.now().strftime("%d %b %Y, %I:%M:%S %p"),
        "file":"Google Sheets (live sync)"}}
    _cache={"data":result,"ts":time.time()}
    return result, None

@app.route("/api/leads")
def api_leads():
    data, err = fetch_sheet()
    if err:
        return jsonify({"leads":FALLBACK,"meta":{"total":len(FALLBACK),"source":"fallback",
            "last_modified":datetime.now().strftime("%d %b %Y, %I:%M %p"),
            "file":"Embedded data — add SHEET_URL in Railway vars","warning":err}})
    return jsonify(data)

@app.route("/api/status")
def status():
    return jsonify({"sheet_configured":bool(SHEET_URL),
        "cache_age":round(time.time()-_cache["ts"]) if _cache["ts"] else None})

@app.route("/")
def index():
    return send_from_directory(".","dashboard.html")

if __name__ == "__main__":
    port=int(os.environ.get("PORT",5000))
    print(f"\n{'='*50}\n  ALIMCO Export Dashboard")
    print(f"  SHEET_URL: {'✅ set' if SHEET_URL else '❌ NOT SET'}")
    print(f"  Port: {port}\n{'='*50}\n")
    app.run(host="0.0.0.0",port=port,debug=False)
