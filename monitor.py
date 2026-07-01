import os
import smtplib
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import requests
import xlsxwriter  # noqa: F401 – used via xlsxwriter engine

# ── CONFIG (all values come from GitHub Secrets / env vars) ──
FROM     = os.environ["EMAIL_FROM"]
TO       = os.environ["EMAIL_TO"]          # comma-separated
CC       = os.environ.get("EMAIL_CC", "")  # optional
APP_PASS = os.environ["EMAIL_APP_PASS"]

TIMEOUT    = 30    # seconds to wait per request
RETRIES    = 2     # number of extra attempts if site is unreachable
RETRY_WAIT = 5     # seconds to wait between retries
SLOW_MS    = 3000  # response time threshold for "Slow" warning (ms)
MAX_WORKERS = 10   # parallel site checks at once

SITES = [
    ("Qatar Steel",        "https://www.qatarsteel.com.qa/"),
    ("Qcoat",              "https://www.qcoat.com.qa/"),
    ("Diyafah",            "https://www.diyafah.com/"),
    ("Sriarunodayam",      "https://sriarunodayam.org/"),
    ("Pixint",             "https://www.pixint.com/"),
    ("Rarefly",            "https://www.rarefly.com/"),
    ("Shadesandlights",    "https://www.shadesandlights.com/"),
    ("Villagescreens",     "https://www.villagescreens.com/"),
    ("Maneshmadhavan",     "https://www.maneshmadhavan.com/"),
    ("Knhdas",             "https://knhdas.com/"),
    ("SSDassociates",      "https://www.ssdassociates.in/"),
    ("Pixintprojects",     "https://pixintprojects.com/"),
    ("Funinwisconsin",     "https://www.funinwisconsin.com/"),
    ("Photoscanplus",      "https://photoscanplus.com/"),
    ("Lanternglobal",      "http://lanternglobal.com/"),
    ("Stellarixsolutions", "https://stellarixsolutions.com/"),
    ("Kindernest",         "https://kindernest.in/"),
    ("Rickgraffpaints",    "https://www.rickgraffpaints.com/"),
    ("Greenleafchina",     "https://greenleafchina.com/"),
    ("Spicenet",           "https://www.spicenet.info/"),
]

# ── 1. CHECK A SINGLE SITE (with retries) ────────────────────
def check_site(index_name_url):
    idx, name, url = index_name_url
    attempt, status, error, response_ms = 0, "Unreachable", "-", 0

    for attempt in range(1 + RETRIES):
        try:
            t0   = time.monotonic()
            resp = requests.get(url, timeout=TIMEOUT, allow_redirects=True)
            response_ms = round((time.monotonic() - t0) * 1000)

            if resp.status_code < 400:
                status = "Slow ⚠️" if response_ms > SLOW_MS else "Reachable"
                error  = "-"
                break                          # success — no more retries
            else:
                error = f"HTTP {resp.status_code}"

        except requests.exceptions.Timeout:
            error = "Timed out"
        except requests.exceptions.ConnectionError:
            error = "Connection error"
        except requests.exceptions.RequestException as e:
            error = str(e)[:60]

        # Only retry if not the last attempt
        if attempt < RETRIES:
            print(f"  ↻ Retry {attempt + 1}/{RETRIES} for {name} ...")
            time.sleep(RETRY_WAIT)

    # If still not reachable after all retries
    if status == "Unreachable":
        icon = "❌"
    elif status.startswith("Slow"):
        icon = "⚠️"
    else:
        icon = "✅"

    print(f"[{icon}] {name:25s} {status:15s} {response_ms:>6} ms  {error}")
    return (idx, name, url, status, response_ms, error)


# ── 2. RUN ALL CHECKS IN PARALLEL ───────────────────────────
print(f"🔍 Checking {len(SITES)} sites  (timeout={TIMEOUT}s, retries={RETRIES})\n")
tasks   = [(i + 1, name, url) for i, (name, url) in enumerate(SITES)]
results = {}

with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
    futures = {pool.submit(check_site, t): t for t in tasks}
    for future in as_completed(futures):
        row = future.result()
        results[row[0]] = row   # key by S.No so we can sort

rows = [results[k] for k in sorted(results)]

# ── 3. BUILD XLSX ────────────────────────────────────────────
today     = date.today().isoformat()
xlsx_file = f"site-report-{today}.xlsx"
df = pd.DataFrame(rows, columns=["S.No", "Website Name", "URL", "Status", "Response (ms)", "Error"])

with pd.ExcelWriter(xlsx_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Report", index=False, header=False, startrow=1)
    wb, ws = writer.book, writer.sheets["Report"]

    # ── Formats
    hdr    = wb.add_format({"bold": True, "fg_color": "#D7E4BC", "border": 1, "valign": "top", "text_wrap": True})
    green  = wb.add_format({"bg_color": "#C6EFCE", "border": 1})
    red    = wb.add_format({"bg_color": "#FFC7CE", "border": 1})
    orange = wb.add_format({"bg_color": "#FFEB9C", "border": 1})   # Slow warning
    center = wb.add_format({"align": "center", "border": 1})

    for c, col in enumerate(df.columns):
        ws.write(0, c, col, hdr)

    sc = df.columns.get_loc("Status")
    ws.conditional_format(1, sc, len(df), sc,
        {"type": "text", "criteria": "containing", "value": "Reachable", "format": green})
    ws.conditional_format(1, sc, len(df), sc,
        {"type": "text", "criteria": "containing", "value": "Unreachable", "format": red})
    ws.conditional_format(1, sc, len(df), sc,
        {"type": "text", "criteria": "containing", "value": "Slow", "format": orange})

    # Highlight slow response times in orange
    rc = df.columns.get_loc("Response (ms)")
    ws.conditional_format(1, rc, len(df), rc,
        {"type": "cell", "criteria": ">=", "value": SLOW_MS, "format": orange})

    for col, w in [("A:A",6), ("B:B",22), ("C:C",42), ("D:D",14), ("E:E",14), ("F:F",50)]:
        ws.set_column(col, w)

    ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    ws.freeze_panes(1, 0)   # freeze header row

    # ── Summary sheet
    total      = len(df)
    up_count   = len(df[df["Status"] == "Reachable"])
    slow_count = len(df[df["Status"].str.startswith("Slow", na=False)])
    down_count = len(df[df["Status"] == "Unreachable"])

    ws_sum = wb.add_worksheet("Summary")
    bold   = wb.add_format({"bold": True, "font_size": 13})
    ws_sum.write("A1", "Summary",         bold)
    ws_sum.write("A2", "Total Sites",     wb.add_format({"bold": True}))
    ws_sum.write("B2", total)
    ws_sum.write("A3", "✅ Reachable",    wb.add_format({"bold": True, "bg_color": "#C6EFCE"}))
    ws_sum.write("B3", up_count)
    ws_sum.write("A4", "⚠️ Slow (>3s)",  wb.add_format({"bold": True, "bg_color": "#FFEB9C"}))
    ws_sum.write("B4", slow_count)
    ws_sum.write("A5", "❌ Unreachable",  wb.add_format({"bold": True, "bg_color": "#FFC7CE"}))
    ws_sum.write("B5", down_count)
    ws_sum.write("A6", "Report Date")
    ws_sum.write("B6", today)
    ws_sum.set_column("A:A", 22)
    ws_sum.set_column("B:B", 12)

print(f"\n📊 Report saved: {xlsx_file}")
print(f"   ✅ {up_count} Reachable  ⚠️ {slow_count} Slow  ❌ {down_count} Down")

# ── 4. SEND EMAIL ────────────────────────────────────────────
to_list = [t.strip() for t in TO.split(",") if t.strip()]
cc_list = [c.strip() for c in CC.split(",") if c.strip()]

down_names = df[df["Status"] == "Unreachable"]["Website Name"].tolist()

subject = f"{'🚨 ALERT: ' if down_names else ''}Site Reachability Report – {today}"

# ── Email body: only Reachable / Not Reachable — no slow mention ──
if not down_names:
    body = """
<html><body style='font-family:Arial,sans-serif;font-size:14px'>
<p>Hello Team,</p>
<p style='color:green;font-size:16px'><b>✅ All websites are reachable.</b></p>
<p>Please find the detailed report in the attached Excel file.</p>
<br><p>Best Regards,<br><b>Ajaykumar R</b><br>System Admin</p>
</body></html>"""
else:
    items = "".join(
        f"<li style='color:red;margin:4px 0'><b>{s}</b> — Not Reachable</li>"
        for s in down_names
    )
    body = f"""
<html><body style='font-family:Arial,sans-serif;font-size:14px'>
<p>Hello Team,</p>
<p>The following site(s) are <b style='color:red'>not reachable</b>:</p>
<ul style='line-height:1.8'>{items}</ul>
<p style='color:#555'>Please check the attached Excel report for full details.</p>
<br><p>Best Regards,<br><b>Ajaykumar R</b><br>System Admin</p>
</body></html>"""

msg = MIMEMultipart()
msg["From"]    = FROM
msg["To"]      = TO
msg["Cc"]      = CC
msg["Subject"] = subject
msg.attach(MIMEText(body, "html"))

with open(xlsx_file, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={xlsx_file}")
    msg.attach(part)

ctx = ssl.create_default_context()
with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls(context=ctx)
    server.login(FROM, APP_PASS)
    server.sendmail(FROM, to_list + cc_list, msg.as_string())

print(f"📧 Email sent to: {', '.join(to_list + cc_list)}")
