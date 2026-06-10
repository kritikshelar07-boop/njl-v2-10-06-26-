"""
NJL Catalogue — Azure Blob-Backed FastAPI Server
=================================================
• Fetches master_serial_enriched.parquet directly from Azure Blob Storage at startup.
• Caches the file in /tmp so cold-restarts skip re-download if the blob hasn't changed
  (checked via HTTP ETag / Last-Modified headers).
• Background thread re-checks the blob every REFRESH_INTERVAL_MINUTES and hot-reloads
  the dataset if a new version is detected — no restart required after you replace the file.
• Render-optimised: reads PORT env var, binds 0.0.0.0, single-process safe.
"""

import os
import math
import time
import socket
import hashlib
import logging
import threading
import numpy as np
import pandas as pd
import uvicorn
import httpx                          # lighter than requests; in requirements.txt
from fastapi import FastAPI, Query
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware

# ─── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("njl")

# ─── Configuration ────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
HTML_FILE  = os.path.join(BASE_DIR, "index.html")

# Azure public blob URL — update this if the container/blob path ever changes
BLOB_URL = os.environ.get(
    "PARQUET_BLOB_URL",
    "https://njlprodimages.blob.core.windows.net/products/master_serial_enriched.parquet",
)

# Local cache path in /tmp (survives within a Render dyno session)
LOCAL_CACHE_PATH = "/tmp/master_serial_enriched.parquet"
ETAG_CACHE_PATH  = "/tmp/master_serial_enriched.etag"   # stores last ETag/Last-Modified

# How often (minutes) the background thread checks for a new blob version
REFRESH_INTERVAL_MINUTES = int(os.environ.get("REFRESH_INTERVAL_MINUTES", "30"))

PORT = int(os.environ.get("PORT", 8000))
HOST = "0.0.0.0"
MAX_FILTER_RESULTS = 1000

DOWNLOAD_TIMEOUT_SECONDS = 120   # large parquet can take a while on cold start

# ─── In-memory state (protected by a read-write lock pattern via threading.Lock) ─
_data_lock = threading.Lock()

df_global: pd.DataFrame = None
TOTAL_COUNT: int        = 0
SERIAL_CACHE:      list = []
HUID_CACHE:        list = []
WAREHOUSE_CACHE:   list = []
LOCATION_CACHE:    list = []
DUP_HUID_CACHE:    list = []
ROW_TEXT_CACHE           = None

CATEGORY_CACHE:        list = []
SUBCATEGORY_CACHE:     list = []
PRODUCT_GROUP_CACHE:   list = []
SUB_PROD_GRP_CACHE:    list = []
SKU_STATUS_CACHE:      list = []
VENDOR_CACHE:          list = []

SEARCH_COLS = [
    "SERIALNUMBER", "ITEMID", "VENDACCOUNT", "HUID", "WAREHOUSE", "WAREHOUSE_NAME",
    "Category", "Subcategory", "Product_Group", "Sub_Product_Group", "PWC_SKUSTATUS",
    "LOCATION",
]

SORT_MAP = {
    "serial_asc":    ("SERIALNUMBER",       True),
    "serial_desc":   ("SERIALNUMBER",       False),
    "weight_asc":    ("NETWEIGHT",          True),
    "weight_desc":   ("NETWEIGHT",          False),
    "category":      ("Category",           True),
    "product_group": ("Product_Group",      True),
    "avail_desc":    ("AVAILABLE_PHYSICAL", False),
    "avail_asc":     ("AVAILABLE_PHYSICAL", True),
}

# ─── Azure Blob Helpers ───────────────────────────────────────────────────────

def _read_cached_etag() -> str:
    """Return the saved ETag/Last-Modified string, or '' if absent."""
    try:
        with open(ETAG_CACHE_PATH, "r") as f:
            return f.read().strip()
    except FileNotFoundError:
        return ""


def _write_cached_etag(value: str):
    with open(ETAG_CACHE_PATH, "w") as f:
        f.write(value)


def _blob_has_changed() -> bool:
    """
    HEAD the blob. Returns True if the remote ETag/Last-Modified differs from
    what we last downloaded, or if we have no local copy at all.
    """
    if not os.path.exists(LOCAL_CACHE_PATH):
        return True
    try:
        with httpx.Client(timeout=15) as client:
            r = client.head(BLOB_URL)
            r.raise_for_status()
        remote_sig = r.headers.get("etag") or r.headers.get("last-modified") or ""
        return remote_sig != _read_cached_etag()
    except Exception as exc:
        log.warning(f"[blob] HEAD check failed ({exc}); will attempt fresh download.")
        return True


def _download_blob() -> bool:
    """
    Download the parquet from Azure Blob to LOCAL_CACHE_PATH.
    Saves the ETag for future change-detection.
    Returns True on success, False on failure.
    """
    log.info(f"[blob] Downloading {BLOB_URL} …")
    tmp_path = LOCAL_CACHE_PATH + ".tmp"
    try:
        with httpx.stream("GET", BLOB_URL, timeout=DOWNLOAD_TIMEOUT_SECONDS, follow_redirects=True) as r:
            r.raise_for_status()
            etag = r.headers.get("etag") or r.headers.get("last-modified") or ""
            with open(tmp_path, "wb") as f:
                for chunk in r.iter_bytes(chunk_size=1 << 20):  # 1 MB chunks
                    f.write(chunk)
        os.replace(tmp_path, LOCAL_CACHE_PATH)   # atomic swap
        _write_cached_etag(etag)
        size_mb = os.path.getsize(LOCAL_CACHE_PATH) / (1024 * 1024)
        log.info(f"[blob] Download complete — {size_mb:.1f} MB saved to {LOCAL_CACHE_PATH}")
        return True
    except Exception as exc:
        log.error(f"[blob] Download failed: {exc}")
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False


def ensure_local_parquet() -> bool:
    """
    Guarantees LOCAL_CACHE_PATH exists and is up-to-date.
    Returns True if the file is ready to load.
    """
    if _blob_has_changed():
        return _download_blob()
    log.info("[blob] Local cache is current — skipping download.")
    return True

# ─── Data Loader ──────────────────────────────────────────────────────────────

def _build_caches(df: pd.DataFrame):
    """Build all search + filter caches from a freshly loaded DataFrame."""
    global TOTAL_COUNT, SERIAL_CACHE, HUID_CACHE, WAREHOUSE_CACHE, LOCATION_CACHE
    global CATEGORY_CACHE, SUBCATEGORY_CACHE, PRODUCT_GROUP_CACHE, SUB_PROD_GRP_CACHE
    global SKU_STATUS_CACHE, VENDOR_CACHE, DUP_HUID_CACHE, ROW_TEXT_CACHE

    total = len(df)
    log.info(f"[cache] Building caches for {total:,} rows …")

    # ── Search index ──
    present = [c for c in SEARCH_COLS if c in df.columns]
    if present:
        txt = df[present[0]].astype(str).str.lower()
        for col in present[1:]:
            txt = txt + " " + df[col].astype(str).str.lower()
    else:
        txt = pd.Series([""] * total, index=df.index)

    # ── Simple value caches ──
    def build(col):
        if col not in df.columns:
            return []
        vc = df[col].value_counts().sort_index()
        return [(k, int(v)) for k, v in vc.items() if k and k != "nan"]

    s_counts = df["SERIALNUMBER"].value_counts().sort_index()
    serial   = [(k, int(v)) for k, v in s_counts.items() if k]

    huid_map: dict = {}
    for raw in df["HUID"]:
        for part in str(raw).split(","):
            part = part.strip()
            if part:
                huid_map[part] = huid_map.get(part, 0) + 1
    huid = sorted(huid_map.items())

    w_counts  = df["WAREHOUSE"].value_counts().sort_index()
    warehouse = [(k, int(v)) for k, v in w_counts.items() if k]

    location = build("LOCATION")

    # ── Duplicate HUID cache ──
    huid_exploded = (
        df[["HUID", "ITEMID"]]
        .assign(HUID=df["HUID"].str.split(","))
        .explode("HUID")
    )
    huid_exploded["HUID"] = huid_exploded["HUID"].str.strip()
    huid_exploded = huid_exploded[huid_exploded["HUID"] != ""]

    grp        = huid_exploded.groupby("HUID")
    sku_counts = grp["ITEMID"].nunique()
    row_counts = grp["ITEMID"].count()
    dup_huids  = sku_counts[sku_counts > 1].index
    dup_skus   = (
        huid_exploded[huid_exploded["HUID"].isin(dup_huids)]
        .groupby("HUID")["ITEMID"]
        .apply(lambda s: sorted(s.unique().tolist()))
    )
    dup = sorted(
        [{"huid": h, "skus": dup_skus[h], "sku_count": int(sku_counts[h]),
          "row_count": int(row_counts[h])} for h in dup_huids],
        key=lambda x: -x["sku_count"],
    )

    # ── Atomic swap into globals ──
    with _data_lock:
        TOTAL_COUNT          = total
        ROW_TEXT_CACHE       = txt
        SERIAL_CACHE[:]      = serial
        HUID_CACHE[:]        = huid
        WAREHOUSE_CACHE[:]   = warehouse
        LOCATION_CACHE[:]    = location
        DUP_HUID_CACHE[:]    = dup
        CATEGORY_CACHE[:]    = build("Category")
        SUBCATEGORY_CACHE[:] = build("Subcategory")
        PRODUCT_GROUP_CACHE[:] = build("Product_Group")
        SUB_PROD_GRP_CACHE[:] = build("Sub_Product_Group")
        SKU_STATUS_CACHE[:]  = build("PWC_SKUSTATUS")
        VENDOR_CACHE[:]      = build("VENDACCOUNT")

    log.info(f"[cache] Done — {total:,} rows, {len(dup)} dup-HUID groups.")


def load_data(force_download: bool = False):
    """
    Main entry point: download if needed, parse parquet, rebuild caches.
    Thread-safe — called at startup and by the background refresh loop.
    """
    global df_global

    if force_download or _blob_has_changed():
        ok = _download_blob()
        if not ok:
            if not os.path.exists(LOCAL_CACHE_PATH):
                raise RuntimeError(
                    "Could not download parquet from Azure and no local cache exists. "
                    "Check PARQUET_BLOB_URL and network connectivity."
                )
            log.warning("[data] Using stale local cache (download failed).")
    else:
        log.info("[data] Local parquet cache is current.")

    log.info(f"[data] Reading {LOCAL_CACHE_PATH} …")
    df = pd.read_parquet(LOCAL_CACHE_PATH)

    for col in ("GROSSQTY", "NETWEIGHT"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    str_cols = [c for c in df.columns if c not in ("GROSSQTY", "NETWEIGHT")]
    df[str_cols] = df[str_cols].fillna("").astype(str)
    df.replace([np.inf, -np.inf], 0.0, inplace=True)

    _build_caches(df)

    with _data_lock:
        df_global = df

    log.info("[data] Dataset live.")


# ─── Background Refresh Thread ────────────────────────────────────────────────

def _refresh_loop():
    """
    Runs in a daemon thread. Every REFRESH_INTERVAL_MINUTES it checks whether
    the blob has been replaced and, if so, transparently reloads the dataset.
    The API keeps serving stale data during the reload — zero downtime.
    """
    interval = REFRESH_INTERVAL_MINUTES * 60
    log.info(f"[refresh] Background refresh thread started — interval {REFRESH_INTERVAL_MINUTES} min.")
    while True:
        time.sleep(interval)
        try:
            log.info("[refresh] Checking blob for updates …")
            if _blob_has_changed():
                log.info("[refresh] New version detected — reloading dataset.")
                load_data(force_download=True)
            else:
                log.info("[refresh] No change detected.")
        except Exception as exc:
            log.error(f"[refresh] Error during background reload: {exc}")


# ─── Initial Load ─────────────────────────────────────────────────────────────
load_data()

refresh_thread = threading.Thread(target=_refresh_loop, daemon=True)
refresh_thread.start()

# ─── FastAPI App ──────────────────────────────────────────────────────────────
app = FastAPI(title="NJL Catalogue API", version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── Static Asset Routes ──────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def serve_html():
    if not os.path.exists(HTML_FILE):
        return HTMLResponse("<h2>Error: index.html not found.</h2>", status_code=404)
    with open(HTML_FILE, "r", encoding="utf-8") as fh:
        return fh.read()

@app.get("/style.css")
def get_css():
    return FileResponse(os.path.join(BASE_DIR, "style.css"))

@app.get("/script.js")
def get_js():
    return FileResponse(os.path.join(BASE_DIR, "script.js"))

@app.get("/logo.jpg")
def get_logo():
    return FileResponse(os.path.join(BASE_DIR, "logo.jpg"))

# ─── API: Health / Stats ──────────────────────────────────────────────────────
@app.get("/api/stats")
def api_stats():
    with _data_lock:
        return {"total": TOTAL_COUNT}

@app.get("/api/health")
def api_health():
    """Render health-check endpoint."""
    with _data_lock:
        ready = df_global is not None
    return {"status": "ok" if ready else "loading", "total": TOTAL_COUNT}

# ─── API: Filter Values ───────────────────────────────────────────────────────
@app.get("/api/filter-values")
def api_filter_values(
    type:  str = Query("serial"),
    q:     str = Query(""),
    limit: int = Query(MAX_FILTER_RESULTS, ge=1, le=10_000),
):
    cache_map = {
        "serial":    SERIAL_CACHE,
        "warehouse": WAREHOUSE_CACHE,
        "location":  LOCATION_CACHE,
        "huid":      HUID_CACHE,
    }
    with _data_lock:
        cache = list(cache_map.get(type, HUID_CACHE))

    if q:
        ql = q.lower()
        cache = [(v, c) for v, c in cache if ql in v.lower()]

    return {"values": cache[:limit]}

@app.get("/api/dropdown-values")
def api_dropdown_values(
    type: str = Query(...),
    q:    str = Query(""),
):
    cache_map = {
        "category":          CATEGORY_CACHE,
        "subcategory":       SUBCATEGORY_CACHE,
        "product_group":     PRODUCT_GROUP_CACHE,
        "sub_product_group": SUB_PROD_GRP_CACHE,
        "sku_status":        SKU_STATUS_CACHE,
        "vendor":            VENDOR_CACHE,
    }
    with _data_lock:
        cache = list(cache_map.get(type, []))

    if q.strip():
        ql = q.strip().lower()
        cache = [(v, cnt) for v, cnt in cache if ql in v.lower()]
    return {"values": cache}

# ─── API: Inventory ───────────────────────────────────────────────────────────
@app.get("/api/inventory")
def api_inventory(
    page:           int = Query(1,   ge=1),
    page_size:      int = Query(20,  ge=1, le=200),
    sort:           str = Query("default"),
    q:              str = Query(""),
    serials:        str = Query(""),
    huids:          str = Query(""),
    warehouses:     str = Query(""),
    locations:      str = Query(""),
    categories:     str = Query(""),
    subcategories:  str = Query(""),
    product_groups: str = Query(""),
    sub_prod_grps:  str = Query(""),
    sku_statuses:   str = Query(""),
    vendors:        str = Query(""),
):
    with _data_lock:
        df          = df_global
        row_text    = ROW_TEXT_CACHE

    if df is None:
        return {"error": "Data not yet loaded — please retry in a moment."}, 503

    # ── Filters ──
    if serials.strip():
        s_set = {s.strip() for s in serials.split(",") if s.strip()}
        df    = df[df["SERIALNUMBER"].isin(s_set)]

    if huids.strip():
        h_set = {h.strip() for h in huids.split(",") if h.strip()}
        mask  = df["HUID"].apply(
            lambda cell: any(tok.strip() in h_set for tok in str(cell).split(",") if tok.strip())
        )
        df = df[mask]

    if warehouses.strip():
        w_set = {w.strip() for w in warehouses.split(",") if w.strip()}
        df    = df[df["WAREHOUSE"].isin(w_set)]

    if locations.strip() and "LOCATION" in df.columns:
        l_set = {l.strip() for l in locations.split(",") if l.strip()}
        df    = df[df["LOCATION"].isin(l_set)]

    def apply_exact(df_in, col, param):
        if not param.strip():
            return df_in
        vals = {v.strip() for v in param.split("|") if v.strip()}
        return df_in[df_in[col].isin(vals)] if vals else df_in

    df = apply_exact(df, "Category",          categories)
    df = apply_exact(df, "Subcategory",       subcategories)
    df = apply_exact(df, "Product_Group",     product_groups)
    df = apply_exact(df, "Sub_Product_Group", sub_prod_grps)
    df = apply_exact(df, "PWC_SKUSTATUS",     sku_statuses)
    df = apply_exact(df, "VENDACCOUNT",       vendors)

    if q.strip():
        import re as _re
        terms    = [t.lower() for t in _re.split(r"[,\s]+", q.strip()) if t.strip()]
        row_text = row_text.loc[df.index]
        for term in terms:
            mask     = row_text.str.contains(term, na=False, regex=False)
            df       = df[mask]
            row_text = row_text.loc[df.index]

    total_filtered = len(df)

    if sort in SORT_MAP:
        sort_col, ascending = SORT_MAP[sort]
        df = df.sort_values(sort_col, ascending=ascending, kind="stable")

    start   = (page - 1) * page_size
    page_df = df.iloc[start : start + page_size]

    records = page_df.replace({np.nan: None, np.inf: None, -np.inf: None}).to_dict(orient="records")
    for rec in records:
        for key in ("GROSSQTY", "NETWEIGHT"):
            v = rec.get(key)
            rec[key] = float(v) if v is not None else 0.0

    return {
        "total_filtered": total_filtered,
        "total_all":      TOTAL_COUNT,
        "page":           page,
        "page_size":      page_size,
        "total_pages":    math.ceil(total_filtered / page_size) if total_filtered else 0,
        "data":           records,
    }

# ─── API: Duplicate HUIDs ─────────────────────────────────────────────────────
@app.get("/api/duplicate-huids")
def api_duplicate_huids(
    q:         str = Query(""),
    page:      int = Query(1,   ge=1),
    page_size: int = Query(100, ge=1, le=500),
):
    with _data_lock:
        data = list(DUP_HUID_CACHE)

    if q.strip():
        ql   = q.strip().lower()
        data = [d for d in data if ql in d["huid"].lower()]

    total = len(data)
    start = (page - 1) * page_size
    return {"total": total, "page": page, "page_size": page_size,
            "data": data[start : start + page_size]}

# ─── API: Manual Refresh Trigger (optional, can be secured) ──────────────────
@app.post("/api/refresh")
def api_refresh(secret: str = Query("")):
    """
    Manually trigger a blob re-check and reload.
    Optionally protect with ?secret=<REFRESH_SECRET env var>.
    """
    expected = os.environ.get("REFRESH_SECRET", "")
    if expected and secret != expected:
        from fastapi import HTTPException
        raise HTTPException(status_code=403, detail="Invalid secret.")

    def _bg():
        try:
            load_data(force_download=True)
        except Exception as exc:
            log.error(f"[refresh-api] {exc}")

    threading.Thread(target=_bg, daemon=True).start()
    return {"status": "refresh triggered"}

# ─── Local Execution ──────────────────────────────────────────────────────────
def _find_free_port(start: int) -> int:
    for p in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind((HOST, p))
                return p
            except OSError:
                continue
    raise RuntimeError(f"No free port in [{start}, {start+20})")


if __name__ == "__main__":
    port = _find_free_port(PORT)
    if port != PORT:
        log.warning(f"Port {PORT} busy — using {port}.")
    uvicorn.run("server:app", host=HOST, port=port, reload=False, log_level="info")
