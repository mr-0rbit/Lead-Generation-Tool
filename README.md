# Rynex Security - Automated Lead Generation & Follow-up System (v2)

Finds companies -> works out **which service they need and why** -> saves a proper CSV -> sends an intro email **written for that context** in your HTML design -> sends timed follow-ups -> stops on reply / bounce / unsubscribe.

## Setup
```
pip install -r requirements.txt
python rynex_leadgen.py init          (skip if config.json already exists; v1 configs upgrade automatically)
python rynex_leadgen.py doctor        (checks internet search, file saving, email settings)
```
Edit `config.json`: `company.sender_name`, `company.postal_address`, `smtp.password` (Gmail/Workspace **App Password**).

## Commands
| Command | What it does |
|---|---|
| `find` | Discover leads + hints, save to `data\leads.csv/.xlsx/.db` (saved after every search query) |
| `enrich` | Add hints/employees/decision-makers to leads found by the older version (`--force` redoes all) |
| `preview` | Show the intro + 3 follow-ups a lead would get; writes `data\preview_email.html` to view the design |
| `send [--dry-run]` | Send due intros and follow-ups |
| `run` / `daemon` | One full cycle / repeat every 6 hours |
| `replies`, `stats`, `export`, `import file.csv`, `unsubscribe x@y.com`, `test-email you@gmail.com --service soc` | Utilities |

## What each lead gets (CSV columns)
Company, Website, Industry, Location, **Service Needed**, **Hint / Evidence**, **Hint Posted Date**, Hint Source URL, Other Hints, **Employees (current)** + source, **Decision-Maker** name / title / email / LinkedIn, Other Leaders Found, a LinkedIn search link for finding the decision-maker manually, Company Email/Phone/LinkedIn, Send-To Email, score, status, emails sent, follow-up dates.

**Where hints come from**

| Hint | Service | Source | Dated? |
|---|---|---|---|
| Hiring security roles (SOC analyst, pentester, compliance officer...) | SOC / VAPT / GRC | company careers page, job boards | yes when the page publishes `datePosted`, else "Undated" |
| Security incident in news | SOC | news search | yes |
| Compliance / certification news | GRC | news search | yes |
| Launch / funding / expansion news | VAPT | news search | yes |
| Website mentions ISO 27001, PCI DSS, SOC 2, GDPR, HIPAA | GRC | company website | no |
| Payments / API / customer-data web apps | VAPT | company website | no |
| Nothing else found | by industry | inferred | no |

The strongest and most recent hint decides the email. Follow-ups keep the same service context.

## Email design
`email_template.html` is your Rynex design (logo, black header/footer). Placeholders: `{{BODY}}`, `{{PREHEADER}}`, `{{FOOTER}}`, `{{CONTACT_EMAIL}}`, `{{WEBSITE}}`. I replaced the "automated email, do not reply" line with your unsubscribe/address footer, because these are outreach emails that need replies (and the script reads replies to stop follow-ups). Email wording lives in `config.json` under `services`, `hooks` and `templates`.

## Honest limits
- **Employee count** is best-effort: taken from the company's own site (or structured data), else a search snippet. Blank/"Not found" when unknown.
- **Decision-makers** are only people the company itself lists on its site. Emails are only used if actually published; the script never guesses addresses. If none found, use the LinkedIn search link in the CSV.
- **Hint dates** exist only when the source publishes one.
- News matches are checked by company name but can occasionally be about a different company with a similar name, so **incidents are never mentioned in the email** unless you set `signals.cite_incident_news` to `true`.
- LinkedIn/Instagram/Discord are not scraped or auto-messaged (against their terms, gets accounts banned).

## Sending safely
Warm up slowly (default 25/day), use a separate sending domain with SPF/DKIM/DMARC, keep the postal address and unsubscribe line, and check the laws of the countries you target (CAN-SPAM, GDPR/PECR, PECA).
