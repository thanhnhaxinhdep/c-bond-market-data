# HNX Corporate Bond Dashboard

Scrapes private-placement bond issuance data from [HNX cbonds](https://cbonds.hnx.vn/to-chuc-phat-hanh/thong-tin-phat-hanh), classifies each bond's sector, and publishes a filterable dashboard.

**Live dashboard:** https://thanhnhaxinhdep.github.io/c-bond-market-data/

## How it runs

A GitHub Actions workflow (`.github/workflows/weekly-scrape.yml`) runs every Monday, scraping the 2 most recent pages (~200 latest rows) from HNX, merging them into `hnx_scraped_raw.csv` (existing rows are never overwritten, only deduped by bond code + issue date + volume), rebuilding `hnx_bonds_processed.xlsx` and `index.html`, and committing the result back to this repo.

Trigger it manually from the Actions tab (`workflow_dispatch`), optionally overriding how many pages to scrape.

## Running locally

```bash
pip install -r requirements.txt
python hnx_pipeline.py            # interactive: asks how many pages
python hnx_pipeline.py --pages 2  # non-interactive: scrape 2 latest pages
```

## Files

- `hnx_pipeline.py` — scrape → merge/dedup → sector classification → Excel + dashboard build
- `hnx_scraped_raw.csv` — accumulated raw scrape history
- `hnx_bonds_processed.xlsx` — processed data (tenor, sector, remaining days, total value, etc.)
- `index.html` — the dashboard (served via GitHub Pages)
- `toan_bo_doanh_nghiep_cbonds.2.csv` — reference table of ~940 Vietnamese companies used to classify each bond's sector by matching its ticker
- `sector_fallback.csv` — manual sector overrides for tickers not found in the reference table
