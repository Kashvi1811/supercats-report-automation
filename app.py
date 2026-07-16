"""
Report Automation Tool — FastAPI Backend
Transforms raw supercat Excel data and generates a two-sheet output
(raw + summary) matching the reference format.

Stack:
  - FastAPI + uvicorn       (web server)
  - Polars                  (DataFrame processing — columnar, multi-threaded, Rust/Arrow hash join)
  - PyArrow                 (Arrow columnar interchange format)
  - NumPy                   (numeric aggregations in summary)
  - OpenPyXL                (Excel read engine via Polars)
  - XlsxWriter              (Excel write — identical output format)
  - Pandera (Polars schema) (emp_info validation)
  - Logging                 (structured pipeline logs)
"""

import io
import os
import asyncio
import logging
from contextlib import asynccontextmanager
from datetime import datetime


import numpy as np
import polars as pl
import pyarrow  # noqa: F401  (Arrow columnar format used by Polars internally)
import pandera.polars as pa
from pandera.polars import DataFrameSchema, Column
import xlsxwriter

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger(__name__)

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EMP_INFO_PATH = os.path.join(BASE_DIR, "emp_info_1783919003.xlsx")

# On Render, /data is the persistent disk mount; locally, use BASE_DIR
CACHE_DIR = "/data" if os.path.isdir("/data") else BASE_DIR

# Fixed city list from reference output (order preserved)
SUMMARY_CITIES = [
    "Ahmedabad", "Bangalore", "Chandigarh", "Chennai",
    "Coimbatore", "Delhi", "Hyderabad", "Jaipur",
    "Kolkata", "Mumbai", "Pune",
]

# Columns S–W: tme, tme_name, me, me_name, expired_on
NULL_CLEAN_COLS = ["tme", "tme_name", "me", "me_name", "expired_on"]
ZERO_CLEAN_COLS = ["tme", "me"]   # replace '00' and '99999'

# ── Pandera schema for emp_info validation ─────────────────────────────────────
# Uses pandera.polars.DataFrameSchema — the Polars-native API
EMP_INFO_SCHEMA = DataFrameSchema(
    {
        "Employee code": Column(pl.Utf8, nullable=False),
        "Team name":     Column(pl.Utf8, nullable=True),
        "Team_type":     Column(pl.Utf8, nullable=True),
    },
    coerce=True,
)

# ── Startup / Shutdown ─────────────────────────────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("Ready! Starting Report Automation server on http://localhost:8000")
    yield  # server is running here
    log.info("Shutting down …")


# ── FastAPI app ────────────────────────────────────────────────────────────────
app = FastAPI(
    title="SuperCat Report Automation",
    lifespan=lifespan,
)


# ── Helper: load emp_info ──────────────────────────────────────────────────────
def load_emp_mapping(emp_bytes: bytes) -> tuple[dict[str, str], dict[str, str]]:
    """
    Load emp_info mapping from uploaded bytes.
    Validates schema with Pandera.
    """
    log.info("Loading emp_info mapping from uploaded file …")

    # Read with Polars (calamine engine) — infer_schema_length=0 forces all cols as Utf8
    # Calamine is used for maximum performance under 1 minute
    df_emp = pl.read_excel(
        io.BytesIO(emp_bytes),
        engine="xlsx2csv",
        read_options={"columns": ["Employee code", "Team name", "Team_type"]},
        infer_schema_length=0,   # all columns → Utf8 (string)
    )

    # Cast all to string (Polars may infer numeric for Employee code)
    df_emp = df_emp.with_columns([
        pl.col("Employee code").cast(pl.Utf8),
        pl.col("Team name").cast(pl.Utf8),
        pl.col("Team_type").cast(pl.Utf8),
    ])

    # Pandera validation
    try:
        EMP_INFO_SCHEMA.validate(df_emp)
        log.info("  emp_info schema validation: PASSED")
    except Exception as val_err:
        log.warning("  emp_info schema validation warning: %s", val_err)

    # Clean Employee code: strip whitespace, remove trailing ".0"
    df_emp = df_emp.with_columns([
        pl.col("Employee code")
          .str.strip_chars()
          .str.replace(r"\.0$", "", literal=False)
          .alias("Employee code"),
        pl.col("Team name")
          .str.strip_chars()
          .str.to_lowercase()
          .alias("Team name"),
    ])

    # Build Python dicts for O(1) lookup
    emp_codes  = df_emp["Employee code"].to_list()
    team_names = df_emp["Team name"].to_list()
    team_types = df_emp["Team_type"].to_list()

    _EMP_TEAM_NAME_MAP = {
        str(k): (v if v is not None else "") for k, v in zip(emp_codes, team_names)
    }
    _EMP_TEAM_TYPE_MAP = {
        str(k): (v if v is not None else "") for k, v in zip(emp_codes, team_types)
    }

    log.info("  emp_info loaded: %d records", len(df_emp))
    return _EMP_TEAM_NAME_MAP, _EMP_TEAM_TYPE_MAP


# ── Main pipeline ──────────────────────────────────────────────────────────────
def run_pipeline(raw_bytes: bytes, emp_bytes: bytes, date_label: str | None = None) -> dict:
    """
    Full transformation pipeline.

    Parameters
    ----------
    raw_bytes : bytes
        Uploaded raw Excel file content.
    date_label : str, optional
        Date string for the summary sheet title.

    Returns
    -------
    dict with keys:
        'excel_bytes': bytes — the final output Excel file
        'summary'    : list[dict] — city breakdown for UI preview
        'stats'      : dict — row counts at key pipeline stages
    """
    stats: dict[str, int] = {}

    # ── Step 1: Read raw input ────────────────────────────────────────────────
    # xlsx2csv engine: converts xlsx → CSV internally (streaming, ~5–10x faster
    # than openpyxl which parses XML cell-by-cell). infer_schema_length=0 forces
    # all columns to Utf8 (string) — identical data guarantee as before.
    log.info("Step 1: Reading raw input …")
    df = pl.read_excel(
        io.BytesIO(raw_bytes),
        engine="xlsx2csv",
        infer_schema_length=0,  # all columns → Utf8 strings
    )

    # Safety: cast any column not already Utf8 to Utf8 (defensive, ensures
    # identical string-only schema regardless of xlsx2csv inference quirks)
    non_str_cols = [c for c in df.columns if df[c].dtype != pl.Utf8]
    if non_str_cols:
        df = df.with_columns([
            pl.col(c).cast(pl.Utf8) for c in non_str_cols
        ])

    # Strip whitespace from all string columns (equivalent to pandas .str.strip())
    str_cols = df.columns   # all columns are now Utf8
    df = df.with_columns([
        pl.col(c).str.strip_chars().alias(c) for c in str_cols
    ])

    stats["input_rows"] = len(df)
    log.info("  Raw rows: %d", stats["input_rows"])

    # ── Step 2: Load emp_info mapping ────────────────────────────────────────
    team_name_map, team_type_map = load_emp_mapping(emp_bytes)

    # ── Step 3: Map team_type + team_name onto each row by empcode ───────────
    log.info("Step 3: Mapping team_type …")

    # Clean empcode first (strip + remove trailing .0)
    df = df.with_columns([
        pl.col("empcode")
          .cast(pl.Utf8)
          .str.strip_chars()
          .str.replace(r"\.0$", "", literal=False)
          .alias("empcode"),
    ])

    # Build emp lookup DataFrame and join with Polars' native hash join
    # Polars uses Rust/Arrow internally — vectorized, same performance as DuckDB
    emp_df = pl.DataFrame({
        "empcode":    list(team_name_map.keys()),
        "_team_name": list(team_name_map.values()),
        "team_type":  list(team_type_map.values()),
    })

    # LEFT JOIN on empcode — Polars hash join (vectorized, multi-threaded)
    df = df.join(emp_df, on="empcode", how="left")

    # ── Step 4: Filter — keep only records with team_type='SJ' OR team_name='super cats' ──
    log.info("Step 4: Filtering to team_type='SJ' OR team_name='super cats' …")
    df = df.filter(
        (pl.col("team_type").str.strip_chars().str.to_uppercase() == "SJ") |
        (pl.col("_team_name").fill_null("") == "super cats")
    ).drop("_team_name")

    # ── Step 4b: Map 'SJ' to 'super cats' in team_type ────────────────────────
    # (After filter, so we know these are the correct records)
    df = df.with_columns(
        pl.when(pl.col("team_type").str.strip_chars().str.to_uppercase() == "SJ")
          .then(pl.lit("super cats"))
          .otherwise(pl.col("team_type"))
          .alias("team_type")
    )

    stats["after_supercat_filter"] = len(df)
    log.info("  After super cats filter: %d", stats["after_supercat_filter"])

    # ── Step 5: Delete rows where final_data_city == '\N' ────────────────────
    log.info("Step 5: Removing \\N rows in final_data_city …")
    before = len(df)
    df = df.filter(pl.col("final_data_city") != r"\N")
    stats["after_city_filter"] = len(df)
    log.info(
        "  Removed %d \\N rows; remaining: %d",
        before - len(df),
        stats["after_city_filter"],
    )

    # ── Step 6: Replace '\N', '/N', null in ALL string columns with blank space ─
    # The raw data contains the literal 2-character string backslash-N (\N).
    # We must match exactly that string using literal=True with the actual chars.
    log.info("Step 6: Cleaning \\N and /N in all string columns …")
    null_exprs = []
    for col in df.columns:
        if df[col].dtype == pl.Utf8:
            null_exprs.append(
                pl.col(col)
                  .fill_null("")                            # true Polars null → ""
                  .str.replace_all("\\N", " ", literal=True)  # backslash-N → " "
                  .str.replace_all("/N",  " ", literal=True)  # forward-slash-N → " "
                  .alias(col)
            )
    if null_exprs:
        df = df.with_columns(null_exprs)

    # ── Step 7: Replace '00' and '99999' in tme, me ───────────────────────────
    log.info("Step 7: Cleaning 00/99999 in tme/me …")
    zero_exprs = []
    for col in ZERO_CLEAN_COLS:
        if col in df.columns:
            zero_exprs.append(
                pl.col(col)
                  .str.replace("^00$", "", literal=False)
                  .str.replace("^99999$", "", literal=False)
                  .alias(col)
            )
    if zero_exprs:
        df = df.with_columns(zero_exprs)

    # ── Step 8: Map main_city_flag ────────────────────────────────────────────
    log.info("Step 8: Mapping main_city_flag …")
    df = df.with_columns([
        pl.col("main_city_flag")
          .map_elements(
              lambda v: {"1": "Main", "0": "Remote"}.get(v, v) if v is not None else v,
              return_dtype=pl.Utf8,
          )
          .alias("main_city_flag")
    ])

    # ── Step 9: Map paid_flag ─────────────────────────────────────────────────
    log.info("Step 9: Mapping paid_flag …")
    paid_map = {"1": "Paid", "0": "Paid Sibling", "2": "Paid Expired"}
    df = df.with_columns([
        pl.col("paid_flag")
          .map_elements(
              lambda v: paid_map.get(v, v) if v is not None else v,
              return_dtype=pl.Utf8,
          )
          .alias("paid_flag")
    ])

    # ── Step 10: Owner/Orphan classification ──────────────────────────────────
    log.info("Step 10: Classifying Owner/Orphan …")
    # Strip tme and empcode, then classify:
    # Own  → tme is non-empty AND tme == empcode
    # Orphan → everything else
    df = df.with_columns([
        pl.col("tme").str.strip_chars().alias("tme"),
        pl.col("empcode").str.strip_chars().alias("empcode"),
    ])
    df = df.with_columns([
        pl.when(
            (pl.col("tme") != "") & (pl.col("tme") == pl.col("empcode"))
        )
        .then(pl.lit("Own"))
        .otherwise(pl.lit("Orphan"))
        .alias("Owner/Orphan")
    ])



    # ── Step 11: Reorder columns to match output reference ───────────────────
    log.info("Step 11: Reordering columns …")
    desired_cols = [
        "parentid", "companyname", "inserted_on", "empcode", "empname",
        "team_type",
        "red_flag", "irocode", "ironame", "updatetime", "process_type",
        "hotlead_source", "docid", "data_city", "source", "final_data_city",
        "main_city_flag", "paid_flag",
        "tme", "tme_name", "me", "me_name", "expired_on",
        "reporting_head_code", "reporting_head_name",
        "reporting_head_code_2", "reporting_head_name_2",
        "Owner/Orphan",
    ]
    output_cols = [c for c in desired_cols if c in df.columns]
    df = df.select(output_cols)

    stats["final_rows"] = len(df)
    log.info("  Final row count: %d", stats["final_rows"])

    # ── Step 12: Build summary data ───────────────────────────────────────────
    log.info("Step 12: Building summary …")
    summary_data = _build_summary(df)

    # ── Step 13: Write output Excel ───────────────────────────────────────────
    log.info("Step 13: Writing output Excel …")
    excel_bytes = _write_excel(df, summary_data, date_label)

    return {
        "excel_bytes": excel_bytes,
        "summary":     summary_data,
        "stats":       stats,
    }


# ── Summary builder ────────────────────────────────────────────────────────────
def _build_summary(df: pl.DataFrame) -> list[dict]:
    """
    Build city-level Main/Remote summary.
    Returns a list of dicts: [{city, main, remote, total}, …]
    Uses NumPy for the aggregation count (integer array ops).
    """
    # Normalise city column for lookup
    df_s = df.with_columns([
        pl.col("final_data_city").str.strip_chars().str.to_titlecase().alias("_city_norm")
    ])

    flag_series = df_s["main_city_flag"].to_numpy(allow_copy=True)
    city_series = df_s["_city_norm"].to_numpy(allow_copy=True)

    def counts_for(city_filter: str | None) -> tuple[int, int, int]:
        if city_filter is None:
            mask = np.ones(len(flag_series), dtype=bool)
        else:
            mask = city_series == city_filter
        subset_flags = flag_series[mask]
        main   = int(np.sum(subset_flags == "Main"))
        remote = int(np.sum(subset_flags == "Remote"))
        return main, remote, main + remote

    rows = []
    # Pan India total first
    m, r, t = counts_for(None)
    rows.append({"city": "Pan India", "main": m, "remote": r, "total": t})

    for city in SUMMARY_CITIES:
        m, r, t = counts_for(city)
        rows.append({"city": city, "main": m, "remote": r, "total": t})

    return rows


# ── Excel writer ───────────────────────────────────────────────────────────────
def _write_excel(
    df: pl.DataFrame,
    summary_data: list[dict],
    date_label: str | None = None,
) -> bytes:
    """
    Write the two-sheet output Excel and return its bytes.
    Uses XlsxWriter directly (same format as original).
    """
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"constant_memory": False, "in_memory": True})

    # ── Summary sheet ──────────────────────────────────────────────────────────
    ws_sum = wb.add_worksheet("summary")

    # Formats
    title_fmt = wb.add_format({
        "bold": True, "font_size": 13, "font_color": "#FFFFFF",
        "bg_color": "#1F3864", "align": "center", "valign": "vcenter",
        "border": 1,
    })
    header_fmt = wb.add_format({
        "bold": True, "bg_color": "#2E75B6", "font_color": "#FFFFFF",
        "border": 1, "align": "center",
    })
    city_fmt = wb.add_format({"border": 1, "align": "left"})
    num_fmt  = wb.add_format({"border": 1, "align": "center", "num_format": "#,##0"})
    pan_fmt  = wb.add_format({
        "bold": True, "border": 1, "align": "left",
        "bg_color": "#D6E4F7",
    })
    pan_num_fmt = wb.add_format({
        "bold": True, "border": 1, "align": "center",
        "num_format": "#,##0", "bg_color": "#D6E4F7",
    })

    # Column widths
    ws_sum.set_column("A:B", 4)
    ws_sum.set_column("C:C", 38)
    ws_sum.set_column("D:F", 14)

    # Title (row index 1 = row 2 in Excel)
    label = date_label or datetime.now().strftime("%B'%y")
    title_text = f"Super cat hot data| {label} | PAN India"
    ws_sum.merge_range(1, 2, 1, 5, title_text, title_fmt)
    ws_sum.set_row(1, 22)

    # Header row (row index 2)
    ws_sum.write(2, 2, "Branch",       header_fmt)
    ws_sum.write(2, 3, "Main",         header_fmt)
    ws_sum.write(2, 4, "Remote",       header_fmt)
    ws_sum.write(2, 5, "Main+ Remote", header_fmt)
    ws_sum.set_row(2, 18)

    # Data rows start at index 3
    for i, row in enumerate(summary_data):
        r = i + 3
        is_pan = row["city"] == "Pan India"
        c_fmt  = pan_fmt     if is_pan else city_fmt
        n_fmt  = pan_num_fmt if is_pan else num_fmt
        ws_sum.write(r, 2, row["city"],   c_fmt)
        ws_sum.write(r, 3, row["main"],   n_fmt)
        ws_sum.write(r, 4, row["remote"], n_fmt)
        ws_sum.write(r, 5, row["total"],  n_fmt)
        ws_sum.set_row(r, 16)

    # ── Raw sheet ──────────────────────────────────────────────────────────────
    ws_raw = wb.add_worksheet("raw")

    raw_header_fmt = wb.add_format({
        "bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF",
        "border": 1, "align": "center", "text_wrap": True,
    })

    col_names = df.columns

    # Write header row
    for col_idx, col_name in enumerate(col_names):
        ws_raw.write(0, col_idx, col_name, raw_header_fmt)

    # Auto-width columns (approximate)
    for col_idx, col_name in enumerate(col_names):
        max_len = max(len(str(col_name)), 10)
        ws_raw.set_column(col_idx, col_idx, min(max_len + 2, 30))

    # Freeze top row
    ws_raw.freeze_panes(1, 0)

    # Write data column-by-column — much faster than row-by-row.
    # Polars extracts each column as a Python list in one vectorised call;
    # XlsxWriter writes the whole list with a single write_column() call.
    # None → "" so xlsxwriter writes a blank string cell (not an error).
    # Final safety net: also blank out any remaining \N or /N strings.
    _BLANK_VALS = {"\\N", "/N", r"\N"}
    for col_idx, col_name in enumerate(col_names):
        col_data = []
        for v in df[col_name].to_list():
            if v is None:
                col_data.append("")
            elif isinstance(v, str) and v.strip() in _BLANK_VALS:
                col_data.append(" ")
            else:
                col_data.append(v)
        ws_raw.write_column(1, col_idx, col_data)

    wb.close()
    buf.seek(0)
    return buf.read()


# ── Static files ───────────────────────────────────────────────────────────────
# Serve index.html, style.css, app.js from the same directory
app.mount(
    "/static",
    StaticFiles(directory=BASE_DIR),
    name="static",
)


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index():
    """Serve the main UI."""
    html_path = os.path.join(BASE_DIR, "index.html")
    with open(html_path, encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/style.css")
async def serve_css():
    """Serve style.css."""
    return FileResponse(os.path.join(BASE_DIR, "style.css"), media_type="text/css")


@app.get("/app.js")
async def serve_js():
    """Serve app.js."""
    return FileResponse(
        os.path.join(BASE_DIR, "app.js"),
        media_type="application/javascript",
    )


def extract_month_from_filename(filename: str) -> str:
    months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    filename_lower = filename.lower()
    for m in months:
        if m.lower() in filename_lower:
            return m.capitalize()
    return datetime.now().strftime("%B")

@app.post("/process")
async def process(
    file: UploadFile = File(...),
    emp_file: UploadFile = File(...),
    date_label: str = Form(default=""),
):
    """
    POST /process
    Form fields:
        file       — uploaded raw Excel (.xlsx)
        emp_file   — uploaded emp_info Excel (.xlsx)
        date_label — optional string for summary title
    Returns JSON with { summary, stats, status, month }
    """
    if not file.filename.lower().endswith(".xlsx") or not emp_file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported.")

    raw_bytes = await file.read()
    emp_bytes = await emp_file.read()
    
    extracted_month = extract_month_from_filename(file.filename)
    dl = date_label.strip() if date_label.strip() else extracted_month

    try:
        result = await asyncio.to_thread(run_pipeline, raw_bytes, emp_bytes, date_label=dl)
    except Exception as exc:
        log.exception("Pipeline error")
        raise HTTPException(status_code=500, detail=str(exc))

    # Cache output Excel for download
    out_path = os.path.join(CACHE_DIR, "_output_cache.xlsx")
    with open(out_path, "wb") as f:
        f.write(result["excel_bytes"])

    return JSONResponse({
        "status":  "ok",
        "summary": result["summary"],
        "stats":   result["stats"],
        "month":   dl,
    })


@app.get("/download")
async def download(month: str = "July"):
    """GET /download — send the cached output Excel file."""
    out_path = os.path.join(CACHE_DIR, "_output_cache.xlsx")
    if not os.path.exists(out_path):
        raise HTTPException(
            status_code=404,
            detail="No output ready. Run /process first.",
        )
    filename = f"supercats_output_report_{month}.xlsx"
    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
        },
    )


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8000,
        reload=False,
        log_level="info",
    )
