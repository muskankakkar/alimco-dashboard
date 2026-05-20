# Run this file instead: python test_run.py
# It writes all errors to error_log.txt so nothing is lost

import sys, os, traceback

log = open("error_log.txt", "w", encoding="utf-8")
log.write("=== Starting ===\n")
log.flush()

try:
    log.write("Step 1: basic imports OK\n"); log.flush()
    import json, re, urllib.request, urllib.error
    from datetime import datetime
    log.write("Step 2: stdlib imports OK\n"); log.flush()

    from flask import Flask, jsonify, send_file, request
    log.write("Step 3: Flask import OK\n"); log.flush()

    from flask_cors import CORS
    log.write("Step 4: flask_cors OK\n"); log.flush()

    import pandas as pd
    log.write("Step 5: pandas OK\n"); log.flush()

    from dotenv import load_dotenv
    log.write("Step 6: dotenv OK\n"); log.flush()

    load_dotenv()
    log.write("Step 7: .env loaded OK\n"); log.flush()

    log.write("Step 8: testing Excel read...\n"); log.flush()
    df = pd.read_excel("exports dashboard.xlsx", sheet_name="Leads Data", engine="openpyxl", header=2)
    log.write(f"Step 8 OK: {len(df)} rows found\n"); log.flush()

    log.write("Step 9: testing Gemini key...\n"); log.flush()
    key = os.getenv("GEMINI_API_KEY", "")
    log.write(f"Step 9: GEMINI_API_KEY = {'SET' if key else 'NOT SET'}\n"); log.flush()

    log.write("\n=== ALL STEPS PASSED ===\n")
    log.write("The problem is in app.py startup. Run: python app.py\n")

except Exception as e:
    log.write(f"\n=== FAILED ===\n")
    log.write(traceback.format_exc())

log.close()
print("Done - check error_log.txt in this folder")
