# LinkedIn Company Scraper & Dashboard

Scrape LinkedIn company search results and manage them through a web dashboard.

---

## Requirements

- Python 3.10+
- Google Chrome installed

---

## Installation

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

---

## Scraper

Scrapes company search results from LinkedIn and exports them to an Excel file in the `good_result/` folder.

**1. Configure the search URL**

Open `linkedin_scraper.py` and update `SEARCH_URL` with your LinkedIn search link:

```python
SEARCH_URL = "https://www.linkedin.com/search/results/companies/?..."
```

**2. Run**

```bash
python linkedin_scraper.py
```

**3. Log in**

A Chrome window will open. Log in to LinkedIn manually, then press **Enter** in the terminal.

The scraper will go through all pages automatically and save results to:

```
good_result/linkedin_companies_YYYYMMDD_HHMMSS.xlsx
```

---

## Dashboard

Two-tab web dashboard:
- **LinkedIn** — manage scraped companies from `tracker.xlsx`
- **avaandmed.ariregister** — browse the Estonian Business Registry with financials

**Run**

```bash
python app.py
```

Then open **http://localhost:5000** in your browser.

---

## Äriregister Dashboard — Data Setup

The registry dashboard requires data files downloaded from the Estonian open data portal.  
Large files are not stored in this repo — download them manually and place in the project root.

### Download from [avaandmed.ariregister.rik.ee](https://avaandmed.ariregister.rik.ee/et/avaandmed)

| File to download | What it is |
|---|---|
| `ettevotja_rekvisiidid__yldandmed.json` | Company general data (name, address, contacts, EMTAK) |
| `ettevotja_rekvisiidid__kaardile_kantud_isikud.json` | Board members (Juhatuse liikmed) |
| `1.aruannete_yldandmed_kuni_<date>.csv` | Annual report index — links report IDs to company registry codes |
| `4.2025_aruannete_elemendid_kuni_<date>.csv` | Financial report elements (revenue, profit, assets, employees) |

> Rename the CSV files to match exactly what `app.py` expects, or update the paths in `app.py` (`CSV_FILE`, `CSV_META_FILE`).

### First run

On the first launch `app.py` automatically streams and indexes all data into a local SQLite cache (`ariregister_cache.db`). This takes **5–15 minutes** depending on your hardware. Subsequent starts are instant.

```
Loading report metadata…
Loading financial elements…
Loading board members…
Building ariregister cache from JSON (one-time, may take a few minutes)…
Cache built.
```

### LinkedIn tab features

| Feature | Details |
|---|---|
| Search | Filter by name, location, industry, or description |
| Status filter | Filter by outreach status |
| Score range | Filter by Automation Interest Score (1–10) |
| Inline editing | Edit any field inline — saves to `tracker.xlsx` instantly |
| Sort | Click any column header to sort |

### Äriregister tab features

| Feature | Details |
|---|---|
| Search | Name, address, activity, email |
| Filters | Status, legal form, city/region, EMTAK code, revenue, employees, profit |
| Sort | All columns — server-side, across full dataset |
| Financials | Revenue, operating profit, assets from latest annual report |
| Board members | Active Juhatuse liikmed |

### Status options

`New` · `Contacted` · `Interested` · `Not Interested` · `Follow Up` · `Closed Deal`

---

## Project Structure

```
linkedin-scraper/
├── linkedin_scraper.py                          # LinkedIn scraper
├── app.py                                       # Flask backend
├── tracker.xlsx                                 # LinkedIn company data
├── requirements.txt                             # Python dependencies
├── templates/
│   ├── index.html                               # LinkedIn dashboard
│   └── ariregister.html                         # Äriregister dashboard
├── good_result/                                 # Scraped Excel exports
│
│   — not in repo, download manually —
├── ettevotja_rekvisiidid__yldandmed.json        # ~4 GB
├── ettevotja_rekvisiidid__kaardile_kantud_isikud.json  # ~1 GB
├── 1.aruannete_yldandmed_kuni_<date>.csv        # ~200 MB
├── 4.2025_aruannete_elemendid_kuni_<date>.csv   # ~60 MB
└── ariregister_cache.db                         # auto-generated on first run
```
