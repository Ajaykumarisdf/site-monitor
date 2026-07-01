# 🌐 Site Reachability Monitor

Automatically checks 24 websites daily and emails a colour-coded Excel report.  
Powered by **GitHub Actions** — no server, no cron, no maintenance needed.

---

## 📁 Project Structure

```
site-monitor/
├── monitor.py              # Main script (check → XLSX → email)
├── requirements.txt        # Python dependencies
└── .github/
    └── workflows/
        └── monitor.yml     # GitHub Actions workflow (daily cron)
```

---

## ⚙️ One-Time Setup

### Step 1 — Create a GitHub Repository

```bash
cd site-monitor
git init
git add .
git commit -m "Initial commit"
# Create a repo on github.com, then:
git remote add origin https://github.com/YOUR_USERNAME/site-monitor.git
git push -u origin main
```

### Step 2 — Add GitHub Secrets

Go to your repo → **Settings → Secrets and variables → Actions → New repository secret**

| Secret Name      | Value                          |
|------------------|--------------------------------|
| `EMAIL_FROM`     | `ajaykumar@pixint.com`         |
| `EMAIL_TO`       | `ajaykumarpmk@gmail.com`       |
| `EMAIL_CC`       | `ajaikumarpmk@gmail.com,ajaykumartrainee@gmail.com` |
| `EMAIL_APP_PASS` | `bigp ruwz urzy zmgy`          |

> **Why secrets?** They are encrypted and never visible in logs. Your password is safe.

### Step 3 — That's it! 🎉

The workflow runs **every day at 9:00 AM IST** automatically.

---

## ▶️ Run Manually

Go to your repo → **Actions → Site Reachability Monitor → Run workflow**

---

## 📊 Report

- An Excel report (`.xlsx`) is emailed daily as an attachment
- The same file is also saved under **Actions → your run → Artifacts** (kept 30 days)
- Green rows = Reachable ✅ | Red rows = Unreachable ❌

---

## 🔧 Customise Schedule

Edit `.github/workflows/monitor.yml`:

```yaml
- cron: '30 3 * * *'   # 3:30 AM UTC = 9:00 AM IST
```

Use [crontab.guru](https://crontab.guru) to pick your preferred time.
