# eBay Wireless Mouse Scraper

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python&logoColor=white)
![BeautifulSoup](https://img.shields.io/badge/BeautifulSoup4-lxml-green)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)

A professional Python web scraper that collects complete eBay listing data — including price, condition, availability and units sold — using a two-stage approach, since eBay's search results grid doesn't expose stock/sold data, only the individual item page does.

---

## 🎥 Demo Video
[![eBay Scraper Demo](https://img.youtube.com/vi/iL53M0S7kxk/maxresdefault.jpg)](https://youtu.be/iL53M0S7kxk)

**▶ [Watch on YouTube](https://youtu.be/iL53M0S7kxk)**

---

## What It Does

- Searches eBay for any keyword and collects unique listing links across multiple search-result pages
- Visits every individual item page directly to pull the details the search grid hides
- Extracts 10 complete data points for every listing
- Saves everything instantly to a clean, formatted Excel file
- Retries automatically with backoff if eBay throttles or blocks a request

---

## Key Features

| Feature | Description |
|---|---|
| Two-Stage Scraping | Search pages for links, item pages for full details |
| Session Warm-Up | Visits the homepage first to reduce first-request blocks |
| Automatic Retries | Exponential backoff on failed or blocked requests |
| Regex Link Extraction | Resilient to eBay's frequently-changing CSS class names |
| Condition Tracking | Captures New / Used / Refurbished status per listing |
| Sold & Stock Tracking | Records units sold and current availability |
| Seller Data | Captures seller name and feedback percentage |
| Random Delays | Human-like pacing between requests |
| Error Isolation | One failed item doesn't stop the whole run |
| Parser Fallback | Uses lxml, falls back to Python's built-in parser automatically |

---

## Data Collected

| Column | Description |
|---|---|
| Title | Full listing title |
| Price | Current listing price (USD) |
| Condition | New / Used / Refurbished etc. |
| Availability | Stock remaining, e.g. "More than 10 available" |
| Sold | Units sold on this listing |
| Image URL | Direct link to the listing's main image |
| Seller | Seller username |
| Seller Feedback | Seller's positive feedback percentage |
| Item URL | Direct link to the listing page |
| Notes | Records an error message if that item failed to scrape |

---

## Built With

- Python 3.14
- Requests — HTTP session handling with retry/backoff
- BeautifulSoup4 (lxml, with automatic fallback)
- Openpyxl — Excel file creation and formatting
- Re (Regex) — resilient item-link extraction
- Random — human-like delay patterns
- Logging — clean progress reporting

---

## How to Run

**Step 1 — Install libraries:**
```
pip install requests beautifulsoup4 openpyxl lxml
```

**Step 2 — Set your search keyword:**
Edit `SEARCH_KEYWORD` and `PAGES_TO_SCRAPE` at the top of the script.

**Step 3 — Run the scraper:**
```
python ebay_mouse_scraper.py
```

---

## Terminal Output

```
11:35:38  INFO  Stage 1: collecting item links from 2 search page(s) for 'wireless mouse'
11:35:41  INFO  Page 1: found 63 new unique item links (running total: 63)
11:35:47  INFO  Stage 2: visiting 126 item pages for full details ...
11:36:10  INFO  Progress: 10/126 items scraped
11:37:52  INFO  Progress: 60/126 items scraped
11:39:40  INFO  Progress: 126/126 items scraped
11:39:41  INFO  Saved 126 rows to ebay_wireless_mouse_results.xlsx
11:39:41  INFO  Done. 122/126 items scraped cleanly.
```

---

## Project Files

| File | Description |
|---|---|
| `ebay_mouse_scraper.py` | Main scraper script |
| `ebay_wireless_mouse_results.xlsx` | Scraped output — 126 listings |
| `README.md` | Project documentation |

---

## Sample Output

| Title | Price | Condition | Availability | Sold |
|---|---|---|---|---|
| Apple Magic Mouse 2 (MLA02LL/A) | US $59.99 | Excellent - Refurbished | More than 10 available | 22 sold |
| Dell Wireless Mouse (Black) - WM126 | US $14.99 | New | 6 available | 506 sold |
| New LOGITECH G305 Lightspeed Wireless Gaming Mouse - White | US $37.99 | New | 5 available | 2 sold |

---

## What I Learned

- How a two-stage architecture solves data that's hidden behind detail pages
- How to build resilient link extraction with regex instead of brittle CSS selectors
- How retry-with-backoff logic handles transient blocks gracefully
- How a session warm-up request reduces first-request bot flags
- How to design a parser fallback chain so the script runs even if a dependency is missing

---

## Real World Use Cases

- Resellers comparing competitor listing prices across platforms
- Dropshippers monitoring which products are selling fastest via the Sold count
- Buyers tracking price drops before purchasing
- Market researchers analyzing condition and pricing distribution

---

## About

Built by Aryan — BS Computer Engineering Student, COMSATS University Islamabad.
Specializing in Python web scraping, automation, and data collection.