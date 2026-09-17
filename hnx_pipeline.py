# ============================================================
# HNX BOND PIPELINE — Local Windows
# Scrape → Dedup → Process → Build Dashboard HTML
# ============================================================

import sys, os, io, time, json, argparse
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# Output files go to same folder as this script
BASE_DIR = Path(__file__).parent
CSV_PATH     = BASE_DIR / "hnx_scraped_raw.csv"
EXCEL_PATH   = BASE_DIR / "hnx_bonds_processed.xlsx"
HTML_PATH    = BASE_DIR / "index.html"
SECTOR_PATH  = BASE_DIR / "toan_bo_doanh_nghiep_cbonds.2.csv"
SECTOR_FALLBACK_PATH = BASE_DIR / "sector_fallback.csv"

# ─────────────────────────────────────────────────────────────
# Sector (Ngành) lookup — mirrors the Excel formula:
#   IFNA(INDEX(toan_bo_doanh_nghiep_cbonds!E, MATCH("*"&ticker&"*", ...!B, 0)),
#        INDEX('Tên công ty search ko ra'!E, MATCH("*"&ticker&"*", ...!B, 0)))
# ─────────────────────────────────────────────────────────────
def _table_to_code_sector_map(df):
    code = df["Tên doanh nghiệp"].astype(str).str.extract(r"^([^-\s]+)")[0].str.strip().str.upper()
    sector = df["Lĩnh vực hoạt động"]
    m = {}
    for c, s in zip(code, sector):
        if pd.notna(c) and c and c not in m and pd.notna(s) and str(s).strip():
            m[c] = s
    return m

def load_sector_map():
    m = {}
    if SECTOR_PATH.exists():
        m.update(_table_to_code_sector_map(pd.read_csv(SECTOR_PATH)))
    else:
        print(f"  ⚠ Không thấy {SECTOR_PATH.name}, bỏ qua bảng tra chính.")
    if SECTOR_FALLBACK_PATH.exists():
        fb = _table_to_code_sector_map(pd.read_csv(SECTOR_FALLBACK_PATH))
        for c, s in fb.items():
            m.setdefault(c, s)
    return m

# ─────────────────────────────────────────────────────────────
# STEP 0 — How many pages? (--pages flag for automation, prompt otherwise)
# ─────────────────────────────────────────────────────────────
_parser = argparse.ArgumentParser()
_parser.add_argument("--pages", type=int, default=None,
                      help="Số trang cần scrape (0 = tất cả). Bỏ qua để hỏi tương tác.")
_args, _ = _parser.parse_known_args()
INTERACTIVE = _args.pages is None

print("=" * 55)
print("  HNX Bond Market — Data Pipeline")
print("=" * 55)

if INTERACTIVE:
    print("  Mỗi trang = 100 dòng dữ liệu")
    print("  Nhập 0 = scrape TẤT CẢ trang đến trang cuối\n")
    while True:
        try:
            MAX_PAGES = int(input("  Scrape bao nhiêu trang? >>> ").strip())
            if MAX_PAGES >= 0:
                break
            print("  Nhập số >= 0 nhé.")
        except ValueError:
            print("  Nhập số nguyên thôi nhé.")
else:
    MAX_PAGES = _args.pages
    print(f"  Chế độ tự động: scrape {MAX_PAGES} trang gần nhất (--pages {MAX_PAGES})")

if MAX_PAGES == 0:
    print(f"\n  → Scrape TẤT CẢ trang\n")
else:
    print(f"\n  → Scrape tối đa {MAX_PAGES} trang (~{MAX_PAGES * 100} dòng)\n")

# ─────────────────────────────────────────────────────────────
# STEP 1 — Scrape
# ─────────────────────────────────────────────────────────────
print("─" * 55)
print("STEP 1 — Scraping HNX...")
print("─" * 55)

def make_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--log-level=3")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=opts)

def scrape(max_pages):
    print("  Khởi tạo Chrome (headless)...")
    driver = make_driver()
    url = "https://cbonds.hnx.vn/to-chuc-phat-hanh/thong-tin-phat-hanh"
    driver.get(url)

    current_page = 1
    all_frames = []

    try:
        print("  Chờ trang tải...")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//table//td"))
        )
        time.sleep(3)

        print("  Đổi hiển thị 100 dòng/trang...")
        try:
            sel = Select(driver.find_element(By.ID, "slChangeNumberRecord_1"))
            sel.select_by_value("100")
            time.sleep(5)
        except Exception as e:
            print(f"  ⚠ Không đổi được số dòng: {e}")

        while True:
            if max_pages > 0 and current_page > max_pages:
                print(f"  ✓ Đủ {max_pages} trang.")
                break

            print(f"  Trang {current_page}...", end=" ", flush=True)
            tables = pd.read_html(io.StringIO(driver.page_source))
            df = next(
                (t for t in tables if len(t) > 0 and
                 any(k in "".join(str(c) for c in t.columns)
                     for k in ["Mã TP", "Tên DN"])),
                None
            )
            if df is not None:
                df.dropna(how="all", inplace=True)
                if "STT" in df.columns:
                    df = df[df["STT"] != "STT"]
                all_frames.append(df)
                print(f"{len(df)} dòng")
            else:
                print("không có dữ liệu")

            try:
                next_btn = driver.find_element(By.XPATH, "//a[contains(@class,'next')]")
                href = next_btn.get_attribute("href") or ""
                if "javascript:void(0)" in href or not href:
                    print(f"  ✓ Đến trang cuối ({current_page} trang).")
                    break
                driver.execute_script("arguments[0].click();", next_btn)
                current_page += 1
                time.sleep(3)
            except NoSuchElementException:
                print(f"  ✓ Không còn nút Next. ({current_page} trang)")
                break
            except Exception as e:
                print(f"  ⚠ Lỗi chuyển trang: {e}")
                break

    except Exception as e:
        print(f"  ✗ Lỗi: {e}")
    finally:
        driver.quit()

    if not all_frames:
        raise RuntimeError("Không scrape được dữ liệu nào!")

    raw = pd.concat(all_frames, ignore_index=True)
    print(f"\n  Scrape xong: {len(raw)} dòng thô")
    return raw

new_df = scrape(MAX_PAGES)

# ─────────────────────────────────────────────────────────────
# STEP 1.5 — Merge with existing data (so old rows are never lost)
# ─────────────────────────────────────────────────────────────
print("\n" + "─" * 55)
print("STEP 1.5 — Merging with dữ liệu cũ...")
print("─" * 55)

if CSV_PATH.exists():
    old_df = pd.read_csv(CSV_PATH, encoding="utf-8-sig")
    print(f"  Đã có {len(old_df)} dòng cũ trong {CSV_PATH.name}")
    raw_df = pd.concat([old_df, new_df], ignore_index=True)
else:
    print(f"  Chưa có {CSV_PATH.name}, tạo mới.")
    raw_df = new_df

# ─────────────────────────────────────────────────────────────
# STEP 2 — Dedup
# ─────────────────────────────────────────────────────────────
print("\n" + "─" * 55)
print("STEP 2 — Deduplication...")
print("─" * 55)

before = len(raw_df)
dedup_cols = next(
    (cols for cols in [
        ["Mã TP", "Ngày phát hành", "Khối lượng"],
        ["Mã TP", "Ngày phát hành"],
        ["Mã TP"],
    ] if all(c in raw_df.columns for c in cols)),
    None
)

# keep="last": ưu tiên bản ghi mới scrape (cập nhật Tình trạng, v.v.) khi trùng key với dữ liệu cũ
if dedup_cols:
    raw_df = raw_df.drop_duplicates(subset=dedup_cols, keep="last")
    print(f"  Key: {dedup_cols}")
else:
    raw_df = raw_df.drop_duplicates(keep="last")
    print(f"  Key: tất cả cột")

dupes = before - len(raw_df)
raw_df = raw_df.reset_index(drop=True)
raw_df["STT"] = raw_df.index + 1
print(f"  Xóa {dupes} dòng trùng → còn {len(raw_df)} dòng")

raw_df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
print(f"  Đã lưu dữ liệu gộp → {CSV_PATH.name}")

# ─────────────────────────────────────────────────────────────
# STEP 3 — Process
# ─────────────────────────────────────────────────────────────
print("\n" + "─" * 55)
print("STEP 3 — Processing columns...")
print("─" * 55)

df = raw_df.copy()

# Extract ticker from company name "BID - Ngân hàng..."
if "Tên DN" in df.columns:
    extracted = df["Tên DN"].str.extract(r"^([^-\s]+)")[0].str.strip()
    if "Code công ty" not in df.columns:
        loc = df.columns.get_loc("Tên DN") + 1
        df.insert(loc, "Code công ty", extracted)
        print("  + Cột 'Code công ty' (trích từ Tên DN)")
    else:
        missing = df["Code công ty"].isna()
        if missing.any():
            df.loc[missing, "Code công ty"] = extracted[missing]
            print(f"  + Điền {missing.sum()} dòng thiếu 'Code công ty' (trích từ Tên DN)")

# Parse dates
for col in ["Ngày phát hành", "Ngày đáo hạn"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

# Drop HNX's built-in Kỳ hạn (recomputed below)
if "Kỳ hạn" in df.columns:
    df = df.drop(columns=["Kỳ hạn"])

today = pd.Timestamp(date.today())

if "Ngày phát hành" in df.columns and "Ngày đáo hạn" in df.columns:
    df["Kỳ hạn"]         = ((df["Ngày đáo hạn"] - df["Ngày phát hành"]).dt.days / 365).round(1)
    df["Month"]          = df["Ngày phát hành"].dt.month
    df["year"]           = df["Ngày phát hành"].dt.year
    df["Remaining days"] = (df["Ngày đáo hạn"] - today).dt.days.clip(lower=0)
    print("  + Kỳ hạn, Month, year, Remaining days")

if "Khối lượng" in df.columns and "Mệnh giá" in df.columns:
    df["Total value"] = (pd.to_numeric(df["Khối lượng"], errors="coerce") *
                         pd.to_numeric(df["Mệnh giá"],   errors="coerce"))
    print("  + Total value")

# Classify Ngành (sector) via lookup table, same formula as the Excel workbook
if "Code công ty" in df.columns:
    sector_map = load_sector_map()
    derived = df["Code công ty"].astype(str).str.upper().map(sector_map)
    df["Ngành"] = derived.combine_first(df["Ngành"]) if "Ngành" in df.columns else derived
    print(f"  + Ngành: khớp {derived.notna().sum()}/{len(df)} qua {SECTOR_PATH.name} (+ fallback)")

for col in ["Loại hình", "Ngành"]:
    if col not in df.columns:
        df[col] = ""
df["Ngành"] = df["Ngành"].fillna("")

print(f"  Columns cuối: {list(df.columns)}")

# ─────────────────────────────────────────────────────────────
# STEP 4 — Save Excel backup
# ─────────────────────────────────────────────────────────────
print("\n" + "─" * 55)
print("STEP 4 — Saving Excel backup...")
print("─" * 55)

df_excel = df.copy()
for col in ["Ngày phát hành", "Ngày đáo hạn"]:
    if col in df_excel.columns:
        df_excel[col] = df_excel[col].dt.strftime("%Y-%m-%d")
df_excel.to_excel(EXCEL_PATH, index=False)
print(f"  Saved: {EXCEL_PATH.name}")

# ─────────────────────────────────────────────────────────────
# STEP 5 — Build JSON
# ─────────────────────────────────────────────────────────────
print("\n" + "─" * 55)
print("STEP 5 — Building JSON for dashboard...")
print("─" * 55)

pp_cols = [
    "STT", "Ngày đăng tin", "Tên DN", "Code công ty", "Mã TP", "Tiền tệ",
    "Ngày phát hành", "Ngày đáo hạn", "Khối lượng", "Mệnh giá",
    "Loại lãi suất", "Lãi suất phát hành (%/năm)", "Tình trạng",
    "Ngành", "Loại hình", "Kỳ hạn", "Month", "year", "Remaining days", "Total value"
]
pp_cols = [c for c in pp_cols if c in df.columns]
df_pp = df[pp_cols].copy()

if "year" in df_pp.columns:
    df_pp = df_pp[df_pp["year"] != 1900]

for col in ["Ngày phát hành", "Ngày đáo hạn"]:
    if col in df_pp.columns:
        df_pp[col] = df_pp[col].dt.strftime("%Y-%m-%d")

def safe_val(v):
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return None if np.isnan(v) else float(v)
    return v

pp_records = [{k: safe_val(v) for k, v in row.items()}
              for row in df_pp.to_dict(orient="records")]
pp_json = json.dumps(pp_records, ensure_ascii=False)
print(f"  {len(pp_records)} records ready")

# ─────────────────────────────────────────────────────────────
# STEP 6 — Build Dashboard HTML
# ─────────────────────────────────────────────────────────────
print("\n" + "─" * 55)
print("STEP 6 — Building dashboard HTML...")
print("─" * 55)

today_str = datetime.today().strftime("%d %b %Y")

html = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HNX Bond Market Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh}
header{background:linear-gradient(135deg,#1e3a5f,#0ea5e9);padding:18px 32px;display:flex;align-items:center;justify-content:space-between;border-bottom:2px solid #0ea5e9}
header h1{font-size:22px;font-weight:700;color:#fff;letter-spacing:.5px}
header .subtitle{font-size:13px;color:#bae6fd;margin-top:2px}
header .updated{font-size:11px;color:#7dd3fc;background:rgba(0,0,0,.2);padding:4px 10px;border-radius:20px;white-space:nowrap}
.container{padding:20px 24px}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:20px}
.kpi{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:16px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--accent,#0ea5e9)}
.kpi-label{font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.kpi-value{font-size:22px;font-weight:700;color:#f1f5f9}
.kpi-sub{font-size:11px;color:#94a3b8;margin-top:4px}
.filters{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:16px;margin-bottom:18px;display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}
.filter-group{display:flex;flex-direction:column;gap:4px}
.filter-group label{font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px}
select,input[type=text]{background:#0f172a;border:1px solid #475569;color:#e2e8f0;border-radius:6px;padding:7px 10px;font-size:13px;outline:none;min-width:140px}
select:focus,input:focus{border-color:#0ea5e9}
.btn.secondary{background:#334155;color:#94a3b8;border:none;border-radius:6px;padding:8px 16px;cursor:pointer;font-size:13px;font-weight:600;transition:.15s}
.btn.secondary:hover{background:#475569;color:#e2e8f0}
.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.chart-box{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:16px}
.chart-box h3{font-size:13px;color:#94a3b8;margin-bottom:12px;font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.chart-box canvas{max-height:220px}
.chart-full{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:16px;margin-bottom:18px}
.chart-full h3{font-size:13px;color:#94a3b8;margin-bottom:12px;font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.chart-full canvas{max-height:200px}
#monthly-tooltip{position:fixed;pointer-events:none;background:#1e293b;border:1px solid #334155;border-radius:10px;padding:14px 16px;font-size:12px;min-width:210px;box-shadow:0 8px 24px rgba(0,0,0,.5);z-index:999;display:none}
#monthly-tooltip .tt-title{font-weight:700;color:#f1f5f9;font-size:13px;margin-bottom:10px;border-bottom:1px solid #334155;padding-bottom:6px}
#monthly-tooltip .tt-value{color:#0ea5e9;font-size:15px;font-weight:700;margin-bottom:8px}
#monthly-tooltip .tt-row{display:flex;justify-content:space-between;gap:20px;margin-bottom:4px;align-items:center}
#monthly-tooltip .tt-label{color:#64748b;font-size:11px;text-transform:uppercase;letter-spacing:.4px}
#monthly-tooltip .tt-num{font-weight:600;font-size:12px}
.up{color:#34d399}.down{color:#f87171}.na{color:#475569}
.table-wrap{background:#1e293b;border:1px solid #334155;border-radius:10px;overflow:hidden}
.table-header{padding:14px 16px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #334155}
.table-header h3{font-size:13px;color:#94a3b8;font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.record-count{font-size:12px;color:#64748b}
table{width:100%;border-collapse:collapse;font-size:12.5px}
thead th{background:#0f172a;color:#64748b;font-size:11px;text-transform:uppercase;letter-spacing:.5px;padding:10px 12px;text-align:left;cursor:pointer;white-space:nowrap;user-select:none;position:sticky;top:0;z-index:1}
thead th:hover{color:#0ea5e9}
thead th.sort-asc::after{content:' ↑'}thead th.sort-desc::after{content:' ↓'}
tbody tr{border-bottom:1px solid #1e293b;transition:.1s}
tbody tr:hover{background:#253347}
tbody td{padding:9px 12px;white-space:nowrap;color:#cbd5e1}
.badge{display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
.badge.active{background:#064e3b;color:#34d399}.badge.inactive{background:#4c1d1d;color:#f87171}
.badge.fixed{background:#1e3a5f;color:#7dd3fc}.badge.float{background:#3b1f64;color:#c084fc}.badge.combo{background:#3d2a00;color:#fbbf24}
.pagination{display:flex;justify-content:center;align-items:center;gap:8px;padding:14px;border-top:1px solid #334155}
.pg-btn{background:#334155;color:#94a3b8;border:none;border-radius:5px;padding:5px 11px;cursor:pointer;font-size:12px;transition:.15s}
.pg-btn:hover:not(:disabled){background:#475569;color:#e2e8f0}.pg-btn:disabled{opacity:.35;cursor:not-allowed}
.pg-info{font-size:12px;color:#64748b}
.search-box{position:relative;flex:1;min-width:200px}
.search-box input{width:100%;padding-left:30px}
.search-icon{position:absolute;left:9px;top:50%;transform:translateY(-50%);color:#475569;font-size:13px;pointer-events:none}
.rate-cell{color:#fbbf24;font-weight:600;text-align:right}.value-cell{text-align:right}
.filter-note{font-size:11px;color:#475569;font-style:italic;align-self:flex-end;padding-bottom:8px}
@media(max-width:768px){.charts-row{grid-template-columns:1fr}.kpi-row{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<header>
  <div>
    <h1>&#127970; HNX Bond Market Dashboard</h1>
    <div class="subtitle">Vietnam Corporate Bond Data — Hanoi Stock Exchange (HNX)</div>
  </div>
  <div class="updated">Data: ''' + today_str + '''</div>
</header>
<div class="container">
  <div class="kpi-row" id="pp-kpis"></div>
  <div class="filters">
    <div class="filter-group">
      <label>Search</label>
      <div class="search-box">
        <span class="search-icon">&#128269;</span>
        <input type="text" id="pp-search" placeholder="Company name, Bond code..." oninput="applyFilters()">
      </div>
    </div>
    <div class="filter-group">
      <label>Issuance Year</label>
      <select id="pp-year" onchange="applyFilters()"><option value="">All Years</option></select>
    </div>
    <div class="filter-group">
      <label>Sector</label>
      <select id="pp-nganh" onchange="applyFilters()"><option value="">All Sectors</option></select>
    </div>
    <div class="filter-group">
      <label>Status</label>
      <select id="pp-tinhtrang" onchange="applyFilters()">
        <option value="">All</option>
        <option value="Active">Active</option>
        <option value="Expired">Expired</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Rate Type</label>
      <select id="pp-laisuatloai" onchange="applyFilters()"><option value="">All Types</option></select>
    </div>
    <div class="filter-group" style="align-self:flex-end">
      <button class="btn secondary" onclick="resetFilters()">&#8635; Reset</button>
    </div>
    <span class="filter-note">* Avg Rate &amp; Peak Rate update with filters</span>
  </div>
  <div class="charts-row">
    <div class="chart-box"><h3>Issuance Value by Sector (VND bn)</h3><canvas id="pp-chart-industry"></canvas></div>
    <div class="chart-box"><h3>Number of Issuances by Year</h3><canvas id="pp-chart-year"></canvas></div>
  </div>
  <div class="chart-full">
    <h3>Monthly Issuance Value (VND bn)</h3>
    <canvas id="pp-chart-monthly"></canvas>
  </div>
  <div id="monthly-tooltip">
    <div class="tt-title" id="tt-title"></div>
    <div class="tt-value" id="tt-value"></div>
    <div class="tt-row"><span class="tt-label">M/M</span><span class="tt-num" id="tt-mm"></span></div>
    <div class="tt-row"><span class="tt-label">Y/Y</span><span class="tt-num" id="tt-yy"></span></div>
    <div class="tt-row"><span class="tt-label">YTD vs Prior Year</span><span class="tt-num" id="tt-ytd"></span></div>
  </div>
  <div class="table-wrap">
    <div class="table-header">
      <h3>Bond List</h3>
      <span class="record-count" id="pp-count"></span>
    </div>
    <div style="overflow-x:auto">
    <table id="pp-table">
      <thead><tr>
        <th data-col="Ma TP" onclick="sortTable('Ma TP')">Bond Code</th>
        <th data-col="Code cong ty" onclick="sortTable('Code cong ty')">Ticker</th>
        <th data-col="Ten DN" onclick="sortTable('Ten DN')" style="min-width:200px">Company Name</th>
        <th data-col="Ngay phat hanh" onclick="sortTable('Ngay phat hanh')">Issue Date</th>
        <th data-col="Ngay dao han" onclick="sortTable('Ngay dao han')">Maturity Date</th>
        <th data-col="Ky han" onclick="sortTable('Ky han')">Tenor (yrs)</th>
        <th data-col="Lai suat" onclick="sortTable('Lai suat')">Coupon (%)</th>
        <th data-col="Loai lai suat" onclick="sortTable('Loai lai suat')">Rate Type</th>
        <th data-col="Total value" onclick="sortTable('Total value')">Value (VND bn)</th>
        <th data-col="Nganh" onclick="sortTable('Nganh')">Sector</th>
        <th data-col="Tinh trang" onclick="sortTable('Tinh trang')">Status</th>
        <th data-col="year" onclick="sortTable('year')">Year</th>
      </tr></thead>
      <tbody id="pp-tbody"></tbody>
    </table>
    </div>
    <div class="pagination">
      <button class="pg-btn" id="pp-prev" onclick="changePage(-1)">&#8592; Prev</button>
      <span class="pg-info" id="pp-page-info"></span>
      <button class="pg-btn" id="pp-next" onclick="changePage(1)">Next &#8594;</button>
    </div>
  </div>
</div>
<script>
const PP_RAW=''' + pp_json + ''';
const SECTOR_EN={"Tổ chức tín dụng":"Financial Institution","Bất động sản":"Real Estate","Kinh doanh chứng khoán":"Securities","Thương mại, dịch vụ":"Trade & Services","Lĩnh vực khác":"Other Sectors","Năng lượng":"Energy","Sản xuất":"Manufacturing","Xây dựng":"Construction"};
const RATE_TYPE_EN={"Cố định":"Fixed","Thả nổi":"Floating","Kết hợp":"Hybrid"};
const STATUS_EN={"Hiệu lực":"Active","Hết hiệu lực":"Expired"};
const STATUS_VI={"Active":"Hiệu lực","Expired":"Hết hiệu lực"};
const PP_KEY={"Ma TP":"Mã TP","Code cong ty":"Code công ty","Ten DN":"Tên DN","Ngay phat hanh":"Ngày phát hành","Ngay dao han":"Ngày đáo hạn","Ky han":"Kỳ hạn","Lai suat":"Lãi suất phát hành (%/năm)","Loai lai suat":"Loại lãi suất","Total value":"Total value","Nganh":"Ngành","Tinh trang":"Tình trạng","year":"year"};
let ppFiltered=[...PP_RAW],ppSort={col:null,dir:1},ppPage=1;
const PAGE_SIZE=50;
let cI,cY,cM;
function fmt(n){if(n==null||n==='')return'-';return Number(n).toLocaleString('en-US');}
function fmtBil(n){if(n==null||n==='')return'-';return(n/1e9).toLocaleString('en-US',{maximumFractionDigits:1});}
function fmtRate(n){if(n==null||n==='')return'-';return Number(n).toFixed(2)+'%';}
function initFilters(){
  const years=[...new Set(PP_RAW.map(d=>d.year).filter(Boolean))].sort();
  const s=document.getElementById('pp-year');
  years.forEach(y=>{const o=document.createElement('option');o.value=y;o.textContent=y;s.appendChild(o);});
  const nganhs=[...new Set(PP_RAW.map(d=>d[PP_KEY['Nganh']]).filter(Boolean))].sort();
  const s2=document.getElementById('pp-nganh');
  nganhs.forEach(n=>{const o=document.createElement('option');o.value=n;o.textContent=SECTOR_EN[n]||n;s2.appendChild(o);});
  const ls=[...new Set(PP_RAW.map(d=>d[PP_KEY['Loai lai suat']]).filter(Boolean))].sort();
  const s3=document.getElementById('pp-laisuatloai');
  ls.forEach(n=>{const o=document.createElement('option');o.value=n;o.textContent=RATE_TYPE_EN[n]||n;s3.appendChild(o);});
}
function applyFilters(){
  const search=document.getElementById('pp-search').value.toLowerCase();
  const year=document.getElementById('pp-year').value;
  const nganh=document.getElementById('pp-nganh').value;
  const ttVi=STATUS_VI[document.getElementById('pp-tinhtrang').value]||document.getElementById('pp-tinhtrang').value;
  const ls=document.getElementById('pp-laisuatloai').value;
  ppFiltered=PP_RAW.filter(d=>{
    if(year&&String(d.year)!==String(year))return false;
    if(nganh&&d[PP_KEY['Nganh']]!==nganh)return false;
    if(ttVi&&d[PP_KEY['Tinh trang']]!==ttVi)return false;
    if(ls&&d[PP_KEY['Loai lai suat']]!==ls)return false;
    if(search&&!(d[PP_KEY['Ten DN']]||'').toLowerCase().includes(search)&&!(d[PP_KEY['Ma TP']]||'').toLowerCase().includes(search)&&!(d[PP_KEY['Code cong ty']]||'').toLowerCase().includes(search))return false;
    return true;
  });
  if(ppSort.col)doSort(ppSort.col);
  ppPage=1;renderKPIs();renderTable();updateCharts();
}
function resetFilters(){
  ['pp-search','pp-year','pp-nganh','pp-tinhtrang','pp-laisuatloai'].forEach(id=>{document.getElementById(id).value='';});
  ppFiltered=[...PP_RAW];ppPage=1;ppSort={col:null,dir:1};
  document.querySelectorAll('#pp-table th').forEach(t=>t.classList.remove('sort-asc','sort-desc'));
  renderKPIs();renderTable();updateCharts();
}
function sortTable(col){
  document.querySelectorAll('#pp-table th').forEach(t=>t.classList.remove('sort-asc','sort-desc'));
  if(ppSort.col===col)ppSort.dir*=-1;else{ppSort.col=col;ppSort.dir=1;}
  const th=document.querySelector('#pp-table th[data-col="'+col+'"]');
  if(th)th.classList.add(ppSort.dir===1?'sort-asc':'sort-desc');
  doSort(col);ppPage=1;renderTable();
}
function doSort(col){
  const k=PP_KEY[col]||col;
  ppFiltered.sort((a,b)=>{
    let av=a[k]??'',bv=b[k]??'';
    if(typeof av==='number'&&typeof bv==='number')return(av-bv)*ppSort.dir;
    return String(av).localeCompare(String(bv),'en')*ppSort.dir;
  });
}
function renderKPIs(){
  const active=PP_RAW.filter(d=>d[PP_KEY['Tinh trang']]==='Hiệu lực');
  const activeVal=active.reduce((s,d)=>s+(d['Total value']||0),0);
  const rates=ppFiltered.map(d=>d[PP_KEY['Lai suat']]).filter(v=>v!=null&&v>0);
  const avg=rates.length?rates.reduce((a,b)=>a+b,0)/rates.length:0;
  const max=rates.length?Math.max(...rates):0;
  const filtered=ppFiltered.length<PP_RAW.length;
  const kpis=[
    {label:'Total Bonds',value:PP_RAW.length.toLocaleString(),sub:'private placement records',accent:'#0ea5e9'},
    {label:'Active Bonds',value:active.length.toLocaleString(),sub:'of '+PP_RAW.length.toLocaleString()+' total',accent:'#10b981'},
    {label:'Active Value',value:(activeVal/1e12).toFixed(1)+' tril.',sub:'VND',accent:'#f59e0b'},
    {label:'Avg Coupon Rate',value:avg.toFixed(2)+'%',sub:filtered?'filtered ('+ppFiltered.length.toLocaleString()+' bonds)':'all bonds',accent:'#8b5cf6'},
    {label:'Peak Coupon Rate',value:max.toFixed(2)+'%',sub:filtered?'filtered selection':'all-time high',accent:'#ef4444'},
  ];
  document.getElementById('pp-kpis').innerHTML=kpis.map(k=>'<div class="kpi" style="--accent:'+k.accent+'"><div class="kpi-label">'+k.label+'</div><div class="kpi-value">'+k.value+'</div><div class="kpi-sub">'+k.sub+'</div></div>').join('');
}
function badgeLS(v){if(!v)return'-';const cls=v==='Cố định'?'fixed':v==='Thả nổi'?'float':'combo';return'<span class="badge '+cls+'">'+(RATE_TYPE_EN[v]||v)+'</span>';}
function badgeTT(v){if(!v)return'-';return'<span class="badge '+(v==='Hiệu lực'?'active':'inactive')+'">'+(STATUS_EN[v]||v)+'</span>';}
function renderTable(){
  const start=(ppPage-1)*PAGE_SIZE,end=start+PAGE_SIZE;
  document.getElementById('pp-count').textContent=ppFiltered.length.toLocaleString()+' records';
  document.getElementById('pp-tbody').innerHTML=ppFiltered.slice(start,end).map(d=>
    '<tr><td><b>'+(d[PP_KEY['Ma TP']]||'-')+'</b></td><td>'+((d[PP_KEY['Code cong ty']]||'').trim()||'-')+'</td>'+
    '<td style="white-space:normal;max-width:280px;min-width:160px">'+(d[PP_KEY['Ten DN']]||'-')+'</td>'+
    '<td>'+(d[PP_KEY['Ngay phat hanh']]||'-')+'</td><td>'+(d[PP_KEY['Ngay dao han']]||'-')+'</td>'+
    '<td style="text-align:center">'+(d[PP_KEY['Ky han']]!=null?d[PP_KEY['Ky han']]:'-')+'</td>'+
    '<td class="rate-cell">'+fmtRate(d[PP_KEY['Lai suat']])+'</td><td>'+badgeLS(d[PP_KEY['Loai lai suat']])+'</td>'+
    '<td class="value-cell">'+fmtBil(d['Total value'])+'</td>'+
    '<td><span style="font-size:11px;color:#94a3b8">'+(SECTOR_EN[d[PP_KEY['Nganh']]]||d[PP_KEY['Nganh']]||'-')+'</span></td>'+
    '<td>'+badgeTT(d[PP_KEY['Tinh trang']])+'</td><td style="text-align:center">'+(d['year']||'-')+'</td></tr>'
  ).join('');
  const total=Math.ceil(ppFiltered.length/PAGE_SIZE)||1;
  document.getElementById('pp-page-info').textContent='Page '+ppPage+' of '+total;
  document.getElementById('pp-prev').disabled=ppPage<=1;
  document.getElementById('pp-next').disabled=ppPage>=total;
}
function changePage(d){ppPage+=d;renderTable();}
function updateCharts(){
  const colors=['#0ea5e9','#10b981','#f59e0b','#8b5cf6','#ef4444','#06b6d4','#f97316','#84cc16'];
  const indMap={};
  ppFiltered.forEach(d=>{const k=d[PP_KEY['Nganh']]||'Other';indMap[k]=(indMap[k]||0)+(d['Total value']||0);});
  const sorted=Object.entries(indMap).sort((a,b)=>b[1]-a[1]);
  if(cI)cI.destroy();
  cI=new Chart(document.getElementById('pp-chart-industry'),{type:'bar',data:{labels:sorted.map(e=>SECTOR_EN[e[0]]||e[0]),datasets:[{data:sorted.map(e=>e[1]/1e9),backgroundColor:colors,borderRadius:6,borderWidth:0}]},options:{responsive:true,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+c.parsed.x.toLocaleString('en-US',{maximumFractionDigits:0})+' bn VND'}}},scales:{x:{ticks:{color:'#94a3b8',font:{size:10},callback:v=>v.toLocaleString()},grid:{color:'#1e293b66'}},y:{ticks:{color:'#94a3b8',font:{size:10}},grid:{display:false}}}}});
  const yrMap={};
  ppFiltered.forEach(d=>{if(d.year)yrMap[d.year]=(yrMap[d.year]||0)+1;});
  const yKeys=Object.keys(yrMap).sort();
  if(cY)cY.destroy();
  cY=new Chart(document.getElementById('pp-chart-year'),{type:'bar',data:{labels:yKeys,datasets:[{data:yKeys.map(k=>yrMap[k]),backgroundColor:'#0ea5e9',borderRadius:6,borderWidth:0}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#94a3b8'},grid:{display:false}},y:{ticks:{color:'#94a3b8'},grid:{color:'#1e293b66'}}}}});
  const mMap={};
  ppFiltered.forEach(d=>{const dt=d[PP_KEY['Ngay phat hanh']];if(!dt||dt==='None'||dt==='NaT')return;const ym=dt.substring(0,7);mMap[ym]=(mMap[ym]||0)+(d['Total value']||0);});
  const mKeys=Object.keys(mMap).sort(),mVals=mKeys.map(k=>mMap[k]/1e9);
  function pct(c,b){if(b==null||b===0)return null;return((c-b)/b)*100;}
  function fmtPct(p){if(p==null)return'<span class="na">N/A</span>';const s=p>=0?'+':'';const cls=p>=0?'up':'down';const a=p>=0?'▲':'▼';return'<span class="'+cls+'">'+a+' '+s+p.toFixed(1)+'%</span>';}
  function prevYM(ym){const y=parseInt(ym.substring(0,4)),m=parseInt(ym.substring(5,7));return m===1?(y-1)+'-12':y+'-'+String(m-1).padStart(2,'0');}
  function ytdV(ym){const y=ym.substring(0,4),m=parseInt(ym.substring(5,7));let s=0;for(let i=1;i<=m;i++)s+=mMap[y+'-'+String(i).padStart(2,'0')]||0;return s;}
  const ttEl=document.getElementById('monthly-tooltip');
  function showTT(evt,ym,val){
    const lyY=String(parseInt(ym.substring(0,4))-1),mo=parseInt(ym.substring(5,7));
    let ytdLY=0;for(let i=1;i<=mo;i++)ytdLY+=mMap[lyY+'-'+String(i).padStart(2,'0')]||0;
    document.getElementById('tt-title').textContent=ym;
    document.getElementById('tt-value').textContent=val.toLocaleString('en-US',{maximumFractionDigits:1})+' bn VND';
    document.getElementById('tt-mm').innerHTML=fmtPct(pct(val*1e9,mMap[prevYM(ym)]));
    document.getElementById('tt-yy').innerHTML=fmtPct(pct(val*1e9,mMap[lyY+'-'+ym.substring(5,7)]));
    document.getElementById('tt-ytd').innerHTML=fmtPct(pct(ytdV(ym),ytdLY>0?ytdLY:null));
    ttEl.style.display='block';
    let x=evt.clientX+16,y=evt.clientY+16;
    if(x+230>window.innerWidth)x=evt.clientX-230-16;
    if(y+160>window.innerHeight)y=evt.clientY-160-16;
    ttEl.style.left=x+'px';ttEl.style.top=y+'px';
  }
  if(cM)cM.destroy();
  cM=new Chart(document.getElementById('pp-chart-monthly'),{type:'bar',data:{labels:mKeys,datasets:[{data:mVals,backgroundColor:'rgba(14,165,233,0.7)',borderColor:'#0ea5e9',borderWidth:1,borderRadius:3,hoverBackgroundColor:'#0ea5e9'}]},options:{responsive:true,plugins:{legend:{display:false},tooltip:{enabled:false}},scales:{x:{ticks:{color:'#94a3b8',font:{size:10},maxRotation:45,autoSkip:true,maxTicksLimit:36},grid:{display:false}},y:{ticks:{color:'#94a3b8',font:{size:10},callback:v=>v.toLocaleString()},grid:{color:'#1e293b44'}}},onHover:(evt,els)=>{if(els.length)showTT(evt.native,mKeys[els[0].index],mVals[els[0].index]);else ttEl.style.display='none';}}});
  document.getElementById('pp-chart-monthly').addEventListener('mouseleave',()=>ttEl.style.display='none');
}
initFilters();renderKPIs();renderTable();updateCharts();
</script>
</body>
</html>'''

with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(html)

size_kb = HTML_PATH.stat().st_size // 1024
print(f"  Saved: {HTML_PATH.name}  ({size_kb} KB)")

# ─────────────────────────────────────────────────────────────
# DONE
# ─────────────────────────────────────────────────────────────
print("\n" + "=" * 55)
print("  DONE!")
print(f"  Dashboard : {HTML_PATH}")
print(f"  Excel     : {EXCEL_PATH}")
print(f"  Raw CSV   : {CSV_PATH}")
print("=" * 55)
if INTERACTIVE:
    input("\n  Nhấn Enter để đóng cửa sổ...")
