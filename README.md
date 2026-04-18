# 🇮🇳 India Post Bulk Rate Calculator Pro

A Python desktop application to bulk-check India Post shipping rates across multiple pincodes and weights — with **built-in SQLite caching** so already-fetched combinations load instantly without hitting the website again.

---

## 📋 Features

- **Manual Entry** — Enter multiple From Pincodes, To Pincodes, and Weights (comma-separated); the app creates all combinations automatically
- **Excel Import** — Load a `.xlsx` file with rows of From / To / Weight for large batch jobs
- **Export to Excel** — Save the results table to a formatted `.xlsx` file
- **SQLite Cache** — Every fetched result is stored locally in `rates_cache.db`; same combinations load instantly on future runs
- **Manual Cache Control** — Cache never expires automatically; you clear it yourself using the 🗑 Clear Cache button when needed
- **Color-coded Results** — 🟢 Green rows = served from cache | 🔵 Blue rows = freshly fetched

---

## 🖥️ Requirements

- Python 3.8+
- Windows (tested), should work on macOS/Linux too

### Python packages

Install all dependencies with:

```bash
pip install pandas openpyxl playwright
```

Then install the Playwright browser (one-time setup):

```bash
playwright install chromium
```

---

## 📁 Project Structure

```
INDIAPOST_RATE/
│
├── indiapost_scraper.py     ← Main application file
├── rates_cache.db           ← Auto-created on first run (SQLite cache)
└── README.md
```

> `rates_cache.db` is created automatically the first time you run the app. Do not delete it unless you want to clear all cached data.

---

## 🚀 How to Run

```bash
python indiapost_scraper.py
```

---

## 📖 How to Use

### Option 1 — Manual Entry

1. Enter **From Pincode** (one or more, comma-separated)
2. Enter **To Pincodes** (one or more, comma-separated)
3. Enter **Weights in grams** (one or more, comma-separated)
4. Click **🔍 Fetch Manual Input**

The app creates all combinations. For example:

| From   | To                  | Weight    | Combinations |
|--------|---------------------|-----------|--------------|
| 400710 | 452001, 110001      | 50, 100   | 4 total      |

---

### Option 2 — Excel File Import

Prepare an `.xlsx` file with exactly these three column headers:

| From   | To     | Weight |
|--------|--------|--------|
| 400710 | 452001 | 50     |
| 400710 | 110001 | 100    |

Click **📂 Load Excel File**, select your file, and the app processes every row.

---

### Exporting Results

Click **💾 Export Results to Excel** after any fetch to save the results table as `.xlsx`.

---

## 💾 Cache System

| Behaviour | Detail |
|-----------|--------|
| **Where stored** | `rates_cache.db` in the same folder as the script |
| **What is cached** | From PIN + To PIN + Weight + all service prices |
| **Cache hit** | Result loads instantly, browser is never opened |
| **Cache expiry** | **Never** — results stay until you manually clear |
| **How to clear** | Click **🗑 Clear Cache** button → confirm prompt |
| **When to clear** | When India Post updates its rates and you need fresh data |

The status bar always shows how many combinations are currently stored in the cache.

---

## 🔍 Services Detected

The app automatically identifies and extracts rates for:

- Speed Post
- Letter
- Letter Card
- Postcard
- Parcel

---

## ⚠️ Notes

- The app uses a **headless Chromium browser** (Playwright) to load the India Post website. First launch may take a few seconds to start the browser.
- If the India Post website is slow or down, some rows may show `ERROR` in the Service column. Re-run to retry those combinations.
- The app fetches rates for the **Letter** product type by default (as per the India Post calculator URL). To change the product type, edit the `url` variable in `fetch_rates_backend()`.

---

## 👨‍💻 Author

Built by **Vitthal** for bulk India Post rate lookups.  
GitHub: [vitthal8](https://github.com/vitthal8)
