<div align="center">

# 🎯 Rynex LeadForge

### Automated Lead Generation & Context-Aware Outreach Engine

*Finds the right companies → works out **which service they need and why** → sends an intro email written for that exact context → follows up on schedule → stops the moment someone replies.*

![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows_|_Linux-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-Private_Use-black?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Active-brightgreen?style=for-the-badge)

</div>

---

## 👾 What It Does

```python
class RynexLeadForge:
    finds        = "companies matching your target industries"
    infers       = "which service they need (SOC / VAPT / GRC) — and WHY"
    evidences    = "job postings, news, certifications, site signals"
    writes       = "intro + follow-up emails tailored to that evidence"
    stops_on     = "reply", "bounce", "unsubscribe"
    honest       = True   # never guesses emails, never scrapes LinkedIn
```

No generic blasts. Every lead gets a **reason-backed** email — built from a real signal (a hiring post, a news mention, a compliance badge on their site) — not a template with the company name swapped in.

---

## ⚙️ How a Lead Gets Qualified

<div align="center">

| Hint | Service Inferred | Source | Dated? |
|:---|:---:|:---:|:---:|
| Hiring security roles (SOC analyst, pentester, compliance officer) | SOC / VAPT / GRC | Careers page, job boards | ✅ when published |
| Security incident in the news | SOC | News search | ✅ |
| Compliance / certification announcement | GRC | News search | ✅ |
| Launch, funding, or expansion news | VAPT | News search | ✅ |
| Site mentions ISO 27001, PCI DSS, SOC 2, GDPR, HIPAA | GRC | Company website | — |
| Payments / API / customer-data web apps | VAPT | Company website | — |
| Nothing else found | Inferred from industry | — | — |

</div>

The strongest, most recent hint decides the angle. Follow-ups keep that same context — no whiplash between emails.

---

## 🖥️ Commands

<div align="center">

| Command | What it does |
|:---|:---|
| `init` | First-time setup, generates `config.json` |
| `doctor` | Checks internet search, file saving, email settings |
| `find` | Discovers leads + hints → saves to `data/leads.csv` / `.xlsx` / `.db` |
| `enrich` | Backfills hints, employee counts, decision-makers on older leads |
| `preview` | Renders the intro + 3 follow-ups a lead would receive |
| `send [--dry-run]` | Sends everything currently due |
| `run` / `daemon` | One full cycle / repeats every 6 hours |
| `replies` / `stats` / `export` / `import` / `unsubscribe` / `test-email` | Utilities |

</div>

---

## 📊 Every Lead Record Includes

<div align="center">

```
Company · Website · Industry · Location
Service Needed · Hint / Evidence · Hint Date · Source URL
Employees (current) + source
Decision-Maker: name · title · email · LinkedIn
Company Email · Phone · LinkedIn
Send-To Email · Score · Status · Emails Sent · Follow-up Dates
```

</div>

---

## 🚀 Quick Start

```bash
pip install -r requirements.txt

python rynex_leadgen.py init          # skip if config.json already exists
python rynex_leadgen.py doctor        # sanity-check search, storage, SMTP
```

Edit `config.json` → set `company.sender_name`, `company.postal_address`, and `smtp.password` (a Gmail/Workspace **App Password**, not your login password).

---

## 📦 Standalone Builds

No Python install needed for end users — grab a single-file executable:

<div align="center">

![Windows](https://img.shields.io/badge/rynex__leadgen.exe-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![Linux](https://img.shields.io/badge/rynex__leadgen-Linux-FCC624?style=for-the-badge&logo=linux&logoColor=black)

</div>

Built automatically via GitHub Actions (`.github/workflows/build.yml`) — one job compiles the Windows `.exe` on a Windows runner, another compiles the native Linux binary on a Linux runner. See `BUILD_INSTRUCTIONS.md`.

---

## ✉️ Email Design

`email_template.html` is your branded shell (logo, header/footer). Placeholders: `{{BODY}}`, `{{PREHEADER}}`, `{{FOOTER}}`, `{{CONTACT_EMAIL}}`, `{{WEBSITE}}`. Wording for each service/hook lives in `config.json` under `services`, `hooks`, and `templates`.

---

## ⚠️ Honest Limits

- **Employee counts** are best-effort — from the company's own site or a search snippet; blank when unknown.
- **Decision-makers** are only people the company itself publishes. No guessed emails — ever. If none found, a LinkedIn search link is provided instead.
- **Hint dates** only exist when the original source publishes one.
- News matches are name-based and can occasionally hit a same-named company, so **incidents are never cited in emails** unless `signals.cite_incident_news` is explicitly set to `true`.
- LinkedIn / Instagram / Discord are never scraped or auto-messaged — against ToS, risks account bans.

---

## 📬 Sending Safely

<div align="center">

```
Warm up slowly (25/day default)
Use a dedicated sending domain — SPF · DKIM · DMARC configured
Keep the postal address + unsubscribe line on every email
Know the rules that apply to you — CAN-SPAM · GDPR/PECR · PECA
```

</div>

---

<div align="center">

**🔐 Built for outreach that's targeted, honest, and compliant — not spray-and-pray.**

</div>
