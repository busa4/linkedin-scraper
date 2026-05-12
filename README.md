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

Web dashboard to view, filter, and manage all analysed companies from `tracker.xlsx`.

**Run**

```bash
python app.py
```

Then open **http://localhost:5000** in your browser.

### Features

| Feature | Details |
|---|---|
| Search | Filter by name, location, industry, or description |
| Status filter | Filter by outreach status |
| Field filter | Logistics / Retail |
| Potential Needs filter | Filter by automation need type |
| Score range | Filter by Automation Interest Score (1–10) |
| Inline status edit | Change status via dropdown — saves to `tracker.xlsx` instantly |
| Contact Person | Edit contact name inline — saves on Enter or focus loss |
| Sort | Click any column header to sort |
| Score badges | Color coded: 🔴 1–4 · 🟠 5–6 · 🟢 7–8 · 🟩 9–10 |

### Status options

`New` · `Contacted` · `Interested` · `Not Interested` · `Follow Up` · `Closed Deal`

---

## Project Structure

```
linkedin-scraper/
├── linkedin_scraper.py   # Scraper script
├── app.py                # Flask dashboard backend
├── tracker.xlsx          # Company data (source for dashboard)
├── requirements.txt      # Python dependencies
├── templates/
│   └── index.html        # Dashboard frontend
└── good_result/          # Scraped Excel exports
```
