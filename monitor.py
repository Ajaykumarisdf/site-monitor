import os
import smtplib
import ssl
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

SITES = [
    ("Qatar Steel",        "https://www.qatarsteel.com.qa/"),
    ("Webdev Qatar",       "http://webdev.qatarsteel.com.qa/"),
    ("Qcoat",              "https://www.qcoat.com.qa/"),
    ("Webdev Qcoat",       "http://webdev.qcoat.com.qa/"),
    ("Diyafah",            "https://www.diyafah.com/"),
    ("Sriarunodayam",      "https://www.sriarunodayam.org/"),
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
    ("Ecliptasolutions",   "https://www.ecliptasolutions.com/"),
    ("Kindernest",         "https://www.kindernest.in"),
    ("Rickgraffpaints",    "https://www.rickgraffpaints.com/"),
    ("Greenleafchina",     "https://www.greenleafchina.com/"),
    ("Spicenet",           "https://www.spicenet.info/"),
    ("Cadoganglobal",      "https://cadoganglobal.com/"),
]

# ── 1. CHECK SITES ───────────────────────────────────────────
rows = []
for i, (name, url) in enumerate(SITES, start=1):
    try:
        resp = requests.get(url, timeout=15, allow_redirects=True)
        if resp.status_code < 400:
            status, error = "Reachable", "-"
        else:
            status, error = "Unreachable", f"HTTP {resp.status_code}"
    except requests.exceptions.ConnectionError:
        status, error = "Unreachable", "Connection error"
    except requests.exceptions.Timeout:
        status, error = "Unreachable", "Timed out"
    except requests.exceptions.RequestException as e:
        status, error = "Unreachable", str(e)[:60]

    rows.append((i, name, url, status, error))
    print(f"[{'✅' if status == 'Reachable' else '❌'}] {name:25s} {status}")

# ── 2. BUILD XLSX ────────────────────────────────────────────
today     = date.today().isoformat()
xlsx_file = f"site-report-{today}.xlsx"
df        = pd.DataFrame(rows, columns=["S.No", "Website Name", "URL", "Status", "Error"])

with pd.ExcelWriter(xlsx_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Report", index=False, header=False, startrow=1)
    wb, ws = writer.book, writer.sheets["Report"]

    hdr   = wb.add_format({"bold": True, "fg_color": "#D7E4BC", "border": 1, "valign": "top"})
    green = wb.add_format({"bg_color": "#C6EFCE", "border": 1})
    red   = wb.add_format({"bg_color": "#FFC7CE", "border": 1})

    for c, col in enumerate(df.columns):
        ws.write(0, c, col, hdr)

    sc = df.columns.get_loc("Status")
    ws.conditional_format(1, sc, len(df), sc,
        {"type": "text", "criteria": "containing", "value": "Reachable",   "format": green})
    ws.conditional_format(1, sc, len(df), sc,
        {"type": "text", "criteria": "containing", "value": "Unreachable", "format": red})

    for col, w in [("A:A",6),("B:B",25),("C:C",45),("D:D",15),("E:E",50)]:
        ws.set_column(col, w)
    ws.autofilter(0, 0, len(df), len(df.columns) - 1)

print(f"\n📊 Report saved: {xlsx_file}")

# ── 3. SEND EMAIL ────────────────────────────────────────────
to_list = [t.strip() for t in TO.split(",") if t.strip()]
cc_list = [c.strip() for c in CC.split(",") if c.strip()]

down    = df[df["Status"] == "Unreachable"]["Website Name"].tolist()
subject = f"{'ALERT: ' if down else ''}Site Reachability Report – {today}"

if down:
    items = "".join(f"<li style='color:red'><b>{s}</b></li>" for s in down)
    body  = f"""<html><body>
<p>Hello Team,</p>
<p>The following sites are currently <b style='color:red'>unreachable</b>:</p>
<ul>{items}</ul>
<p>Full report is attached.</p>
<br><p>Best Regards,<br><b>Ajaykumar R</b><br>System Admin</p>
</body></html>"""
else:
    body = """<html><body>
<p>Hello Team,</p>
<p style='color:green'><b>✅ All sites are reachable!</b></p>
<p>Detailed report is attached.</p>
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
