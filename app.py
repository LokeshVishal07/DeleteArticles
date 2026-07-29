"""
PUMA EAN Master Mapper
========================
Builds one master reference table across all your listed EANs:

    EAN  |  Marketplace  |  Marketplace PID  |  Color No  |  In ZeCom?  |  Stock (Jan/Jun/Dec/Today)  |  No Stock All Year?

Inputs:
  1. Marketplace listing files — Lazada / Shopee (ZIP) / TikTok / Zalora — as many
     as you have for this region. Each just needs an EAN and the marketplace's own
     Product ID column; everything else in the file is ignored.
  2. ONE Content file (EAN -> Color No) — the bridge, since marketplace files don't
     carry Color No directly.
  3. ONE ZeCom tracker for the selected region (same file works whether it's the
     combined SG+MY workbook or the standalone PH one — the right sheet is picked
     automatically) — used to flag whether each Color No is actually active/valid
     in ZeCom right now, not to pull pricing.
  4. Up to 4 inventory snapshots (Jan / Jun / Dec / Today). Any EAN with zero (or
     missing) stock in every snapshot you provide gets flagged "No stock all year"
     — flagged, not deleted, so you can review before actually delisting anything.

Assumption made explicit: this run is scoped to ONE region/country per session
(matching how the pricing mapper works) — e.g. all 4 marketplace files here would
be Lazada MY + Shopee MY + TikTok MY + Zalora MY together, not multiple countries
in one run. If you actually need multiple countries combined in a single output,
say so and this can be restructured.
"""

import io
import re
import zipfile

import pandas as pd
import streamlit as st

st.set_page_config(page_title="EAN Master Mapper", layout="wide")

# ---------------------------------------------------------------------------
# Shared low-level file reading (same approach as the ZeCom pricing mapper)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def list_sheets(file_bytes, filename):
    attempts = []
    if filename.lower().endswith(".csv"):
        return "csv", ["(csv)"], ["Detected .csv extension"]
    for engine in ["openpyxl", "calamine", "xlrd"]:
        try:
            xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
            return engine, xls.sheet_names, attempts + [f"{engine}: OK"]
        except ImportError:
            attempts.append(f"{engine}: not installed (pip install {('python-calamine' if engine=='calamine' else engine)})")
        except Exception as e:
            attempts.append(f"{engine}: {type(e).__name__}: {e}")
    try:
        tables = pd.read_html(io.BytesIO(file_bytes))
        if tables:
            return "html", [f"Sheet1 ({len(tables)} table(s) found, using 1st)"], attempts + ["html: OK"]
    except Exception as e:
        attempts.append(f"html: {type(e).__name__}: {e}")
    return None, [], attempts


def _read_excel_any_engine(file_bytes, sheet_name, header, nrows=None):
    """
    Try engines in order for the ACTUAL data read, independent of whichever
    engine list_sheets() used to enumerate sheet names. This matters because
    some files (confirmed: real Shopee 'mass update' exports) can have their
    sheet names listed fine by openpyxl, but fail with a ValueError when
    openpyxl actually reads the data — it strictly validates worksheet view
    properties (e.g. the frozen-pane 'activePane' attribute) and rejects
    files where an export tool wrote a non-standard value for it. calamine
    doesn't do this strict validation and reads such files fine.
    Returns (df, engine_used, attempts_log).
    """
    attempts = []
    for engine in ["openpyxl", "calamine", "xlrd"]:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header, dtype=str, engine=engine, nrows=nrows)
            return df, engine, attempts + [f"{engine}: OK"]
        except ImportError:
            attempts.append(f"{engine}: not installed (pip install {('python-calamine' if engine=='calamine' else engine)})")
        except Exception as e:
            attempts.append(f"{engine}: {type(e).__name__}: {e}")
    return None, None, attempts


@st.cache_data(show_spinner=False)
def read_preview(file_bytes, filename, engine, sheet_name, nrows=40):
    if filename.lower().endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str, nrows=nrows)
    if engine == "html":
        tables = pd.read_html(io.BytesIO(file_bytes), header=None)
        raw = tables[0].head(nrows)
        raw.columns = range(raw.shape[1])
        return raw
    raw, used_engine, attempts = _read_excel_any_engine(file_bytes, sheet_name, None, nrows)
    if raw is None:
        raise RuntimeError("Could not read data with any engine:\n" + "\n".join(attempts))
    raw.columns = range(raw.shape[1])
    return raw


@st.cache_data(show_spinner=False)
def read_full(file_bytes, filename, engine, sheet_name, header_row):
    if filename.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes), header=header_row, dtype=str)
    elif engine == "html":
        tables = pd.read_html(io.BytesIO(file_bytes), header=header_row)
        df = tables[0]
    else:
        df, used_engine, attempts = _read_excel_any_engine(file_bytes, sheet_name, header_row)
        if df is None:
            raise RuntimeError("Could not read data with any engine:\n" + "\n".join(attempts))
    df.columns = [str(c) for c in df.columns]
    return df


def excel_col_letter(idx: int) -> str:
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def find_header_row(raw_df: pd.DataFrame, keywords, max_scan=20):
    best_row, best_hits = 0, -1
    for i in range(min(max_scan, len(raw_df))):
        row_vals = [str(v).strip().lower() for v in raw_df.iloc[i].tolist()]
        hits = sum(1 for kw in keywords if any(kw in v for v in row_vals))
        if hits > best_hits:
            best_hits, best_row = hits, i
    return best_row


def build_label_map(columns, banner_vals):
    labels = {}
    for i, col in enumerate(columns):
        letter = excel_col_letter(i)
        banner_val = banner_vals[i] if i < len(banner_vals) else None
        banner_str = (
            str(banner_val).strip()
            if pd.notna(banner_val) and str(banner_val).strip().lower() not in ("", "nan", "none")
            else ""
        )
        display_name = f"Col_{letter}" if str(col).startswith("Unnamed:") else str(col)
        label = f"{letter}: " + (f"{banner_str} — {display_name}" if banner_str else display_name)
        labels[col] = label
    return labels


def strip_template_subheader_rows(df: pd.DataFrame, max_check=5) -> pd.DataFrame:
    """Some marketplace templates (confirmed on a real Lazada export) insert
    'Optional'/'Mandatory' + long description rows right under the header,
    before real data starts. Detect and drop these."""
    drop_idx = []
    for i in range(min(max_check, len(df))):
        row = df.iloc[i]
        non_null = [str(v).strip() for v in row if pd.notna(v) and str(v).strip() != ""]
        if not non_null:
            continue
        junk_hits = sum(
            1 for v in non_null
            if v.lower() in ("optional", "mandatory", "m", "o", "required", "n/a")
            or len(v) > 60
        )
        if junk_hits / len(non_null) >= 0.5:
            drop_idx.append(df.index[i])
        else:
            break
    if drop_idx:
        df = df.drop(index=drop_idx)
    return df.reset_index(drop=True)


def clean_id_str(val, normalize=False):
    if pd.isna(val):
        return None
    if isinstance(val, float):
        s = str(int(val)) if val.is_integer() else str(val)
    else:
        s = str(val).strip()
        if s == "" or s.lower() == "nan":
            return None
        if re.fullmatch(r"\d+\.0+", s):
            s = s.split(".")[0]
    if normalize:
        s = s.strip().upper()
    return s if s != "" else None


def guess_column(columns, hints):
    """Hint-priority order matters: check exact + substring per-hint before moving to the next hint."""
    cols_lower = {c: str(c).strip().lower() for c in columns}
    for h in hints:
        for c, cl in cols_lower.items():
            if cl == h:
                return c
        for c, cl in cols_lower.items():
            if h in cl:
                return c
    return None


def _read_error_message(label, attempt_log):
    all_missing = all(("not installed" in a) or ("importerror" in a.lower()) for a in attempt_log)
    if all_missing:
        return (
            f"**Could not read {label} — but this isn't a problem with your file.**\n\n"
            "None of the Excel-reading packages are installed in the Python environment "
            "currently running this app:\n\n" + "\n".join(f"- {a}" for a in attempt_log)
            + "\n\n**Fix:** make sure `requirements.txt` is committed at the same repo path "
            "as `app.py`, then reboot the app."
        )
    return (
        f"**Could not read {label}.** Tried multiple formats and none worked:\n\n"
        + "\n".join(f"- {a}" for a in attempt_log)
        + "\n\nIf this is a genuine Excel file, try re-saving it as .xlsx from Excel first."
    )


@st.cache_data(show_spinner=False)
def read_zip_combined(file_bytes, header_hints):
    """Unzip every .xlsx/.xls inside, auto-detect each one's header, combine into
    one table, drop duplicate header rows / fully-empty rows / template junk rows."""
    per_file_log = []
    frames = []
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        for name in zf.namelist():
            lower = name.lower()
            if not (lower.endswith(".xlsx") or lower.endswith(".xls")):
                continue
            if "__macosx" in lower or name.startswith("."):
                continue
            try:
                inner_bytes = zf.read(name)
                # Sheet name enumeration (cheap; separate from the actual data
                # read, which gets its own engine-fallback chain below).
                sheet = 0
                for list_engine in ["openpyxl", "calamine", "xlrd"]:
                    try:
                        xls = pd.ExcelFile(io.BytesIO(inner_bytes), engine=list_engine)
                        sheet = xls.sheet_names[0]
                        break
                    except Exception:
                        continue

                preview, engine, attempts = _read_excel_any_engine(inner_bytes, sheet, None, nrows=40)
                if preview is None:
                    per_file_log.append((name, None, None, "all engines failed: " + "; ".join(attempts)))
                    continue
                preview.columns = range(preview.shape[1])
                hdr = find_header_row(preview, header_hints)

                df, engine, attempts = _read_excel_any_engine(inner_bytes, sheet, hdr)
                if df is None:
                    per_file_log.append((name, None, None, "all engines failed: " + "; ".join(attempts)))
                    continue
                df.columns = [str(c).strip() for c in df.columns]
                df = df.dropna(axis=0, how="all")
                df = strip_template_subheader_rows(df)
                frames.append(df)
                per_file_log.append((name, hdr, df.shape, None))
            except Exception as e:
                per_file_log.append((name, None, None, str(e)))

    if not frames:
        return None, per_file_log

    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined = combined.dropna(axis=0, how="all")
    col_set = {str(c).strip().lower() for c in combined.columns}

    def _looks_like_header(row):
        vals = [str(v).strip().lower() for v in row if pd.notna(v)]
        if not vals:
            return False
        hits = sum(1 for v in vals if v in col_set)
        return hits >= max(2, len(vals) // 2)

    mask = combined.apply(_looks_like_header, axis=1)
    combined = combined[~mask].reset_index(drop=True)
    combined = combined.dropna(axis=1, how="all")
    return combined, per_file_log


def read_single_file_auto(label, uploaded_file, header_hints, key_prefix, default_sheet_hint=None, is_zip=False):
    """Full pipeline for one uploaded file: auto sheet/header detection, with a
    collapsed 'fix it' expander as the only manual override. Returns (df, labels)."""
    if uploaded_file is None:
        return None, None
    file_bytes = uploaded_file.getvalue()

    if is_zip:
        with st.spinner(f"Unzipping and combining {label}…"):
            df, per_file_log = read_zip_combined(file_bytes, header_hints)
        if df is None:
            st.error(f"Could not find any readable Excel files inside the ZIP for {label}.")
            for name, hdr, shape, err in per_file_log:
                if err:
                    st.caption(f"- {name}: failed — {err}")
            return None, None
        with st.expander(f"📦 {label} — {len(per_file_log)} file(s) found in ZIP"):
            for name, hdr, shape, err in per_file_log:
                st.caption(f"- {name}: ⚠ skipped ({err})" if err else f"- {name}: header row {hdr}, {shape[0]} rows × {shape[1]} cols")
            st.caption(f"Combined: {df.shape[0]} rows × {df.shape[1]} cols after removing duplicate headers/empty rows.")
        return df, {c: c for c in df.columns}

    engine, sheet_names, attempt_log = list_sheets(file_bytes, uploaded_file.name)
    if engine is None:
        st.error(_read_error_message(label, attempt_log))
        return None, None

    sheet_name = sheet_names[0]
    if len(sheet_names) > 1 and default_sheet_hint:
        for s in sheet_names:
            if default_sheet_hint.lower() == str(s).lower():
                sheet_name = s
                break

    with st.spinner(f"Reading {label}…"):
        raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name)
        auto_header_row = find_header_row(raw, header_hints)

    with st.expander(f"⚙️ Fix detection for {label} (only open if something looks wrong)"):
        if len(sheet_names) > 1:
            sheet_name = st.selectbox("Sheet", options=sheet_names, index=sheet_names.index(sheet_name), key=f"{key_prefix}_sheet")
            raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name)
            auto_header_row = find_header_row(raw, header_hints)
        st.dataframe(raw.head(12), use_container_width=True, height=200)
        header_row = st.number_input("Header row (0 = first row)", min_value=0, max_value=500, value=int(auto_header_row), key=f"{key_prefix}_header_row")
    header_row = int(header_row)

    if header_row >= len(raw):
        raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name, nrows=header_row + 10)

    with st.spinner(f"Loading {label}…"):
        df = read_full(file_bytes, uploaded_file.name, engine, sheet_name, header_row)
    banner_vals = raw.iloc[header_row - 1].tolist() if header_row > 0 else [None] * len(df.columns)
    labels = build_label_map(df.columns, banner_vals)
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    df = strip_template_subheader_rows(df)
    labels = {k: v for k, v in labels.items() if k in df.columns}
    return df, labels


# ---------------------------------------------------------------------------
# Hints
# ---------------------------------------------------------------------------

MARKETPLACE_EAN_HINTS = {
    "Lazada": ["sellersku", "seller sku"],
    "Shopee": ["seller sku", "sku reference no", "sku"],
    "Zalora": ["sellersku", "seller sku"],
    "TikTok Shop": ["seller sku"],
}
MARKETPLACE_PID_HINTS = {
    "Lazada": ["product id", "productid", "sku.skuid", "skuid"],
    "Shopee": ["product id", "productid", "item id", "itemid"],
    "Zalora": ["product id", "productid", "productsetid", "product set id"],
    "TikTok Shop": ["product id", "productid"],
}
MARKETPLACE_HEADER_HINTS = ["sellersku", "seller sku", "sku", "product id", "product name", "seller"]

CONTENT_EAN_HINTS = ["ean"]
CONTENT_PARENT_HINTS = ["color no", "article no", "colorno", "articleno", "style#", "style #"]
CONTENT_HEADER_HINTS = CONTENT_EAN_HINTS + CONTENT_PARENT_HINTS

ZECOM_PARENT_HINTS = [
    "pim article", "pim_article", "pim style",
    "article no", "articleno", "color no", "colorno",
    "style#", "style #",
]
ZECOM_HEADER_HINTS = ZECOM_PARENT_HINTS + ["price", "srp", "rrp", "md price"]

INVENTORY_EAN_HINTS = ["ean", "sellersku", "seller sku", "sku"]
INVENTORY_STOCK_HINTS = ["stock", "qty", "quantity", "available", "inventory", "on hand"]
INVENTORY_HEADER_HINTS = INVENTORY_EAN_HINTS + INVENTORY_STOCK_HINTS


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

st.title("🧭 EAN Master Mapper")
st.caption(
    "Combines all your marketplace listings into one master table: EAN → Color No → Marketplace PID, "
    "cross-checked against ZeCom, with dead-stock flagging across 4 inventory snapshots."
)

with st.sidebar:
    st.header("1. Region")
    region = st.selectbox("Region (for this run's marketplace + ZeCom files)", ["MY", "PH", "SG"])

    st.header("2. Marketplace listing files")
    st.caption("Upload whichever you have for this region — all optional, combined into one output.")
    lazada_file = st.file_uploader("Lazada", type=["xlsx", "xls", "csv"], key="lazada_up")
    shopee_file = st.file_uploader("Shopee (export ZIP)", type=["zip"], key="shopee_up")
    tiktok_file = st.file_uploader("TikTok Shop", type=["xlsx", "xls", "csv"], key="tiktok_up")
    zalora_file = st.file_uploader("Zalora", type=["xlsx", "xls", "csv"], key="zalora_up")

    st.header("3. Content file (bridge: EAN → Color No)")
    content_file = st.file_uploader("Content file", type=["xlsx", "xls", "csv"], key="content_up")

    st.header("4. ZeCom tracker")
    zecom_file = st.file_uploader(f"ZeCom tracker ({region})", type=["xlsx", "xls", "csv"], key="zecom_up")

    st.header("5. Inventory snapshots")
    st.caption("Upload whichever you have — the dead-stock flag only checks the periods you provide.")
    inv_jan = st.file_uploader("January inventory", type=["xlsx", "xls", "csv"], key="inv_jan")
    inv_jun = st.file_uploader("June inventory", type=["xlsx", "xls", "csv"], key="inv_jun")
    inv_dec = st.file_uploader("December inventory", type=["xlsx", "xls", "csv"], key="inv_dec")
    inv_now = st.file_uploader("Current / today inventory", type=["xlsx", "xls", "csv"], key="inv_now")

    st.header("6. Options")
    normalize_keys = st.checkbox("Normalize join keys (trim spaces, uppercase, strip stray .0)", value=True)

marketplace_uploads = {
    "Lazada": (lazada_file, False),
    "Shopee": (shopee_file, True),
    "TikTok Shop": (tiktok_file, False),
    "Zalora": (zalora_file, False),
}
active_marketplaces = {name: f for name, (f, _) in marketplace_uploads.items() if f is not None}

if not active_marketplaces:
    st.info("Upload at least one marketplace listing file in the sidebar to get started.")
    st.stop()
if content_file is None:
    st.info("Upload the Content file (EAN → Color No bridge) to get started.")
    st.stop()

# ---------------------------------------------------------------------------
# Read marketplace files -> one long [EAN, Marketplace, Marketplace PID] table
# ---------------------------------------------------------------------------

st.subheader("Marketplace listings")
listing_frames = []
for name, (f, is_zip) in marketplace_uploads.items():
    if f is None:
        continue
    df, labels = read_single_file_auto(f"{name} listing file", f, MARKETPLACE_HEADER_HINTS, f"mp_{name}", is_zip=is_zip)
    if df is None:
        continue
    ean_guess = guess_column(df.columns, MARKETPLACE_EAN_HINTS[name])
    pid_guess = guess_column(df.columns, MARKETPLACE_PID_HINTS[name])
    with st.expander(f"⚙️ Fix {name} column detection"):
        c1, c2 = st.columns(2)
        with c1:
            ean_col = st.selectbox(
                f"{name} — EAN / Seller SKU column", options=list(df.columns),
                index=list(df.columns).index(ean_guess) if ean_guess in df.columns else 0,
                format_func=lambda c: labels.get(c, c), key=f"{name}_ean_col",
            )
        with c2:
            pid_col = st.selectbox(
                f"{name} — Product ID column", options=list(df.columns),
                index=list(df.columns).index(pid_guess) if pid_guess in df.columns else 0,
                format_func=lambda c: labels.get(c, c), key=f"{name}_pid_col",
            )
    sub = pd.DataFrame({
        "EAN": df[ean_col].apply(lambda v: clean_id_str(v, normalize_keys)),
        "Marketplace": name,
        "Marketplace PID": df[pid_col].apply(lambda v: clean_id_str(v, False)),
    })
    sub = sub.dropna(subset=["EAN"])
    listing_frames.append(sub)
    st.caption(f"✓ {name}: {len(sub)} listed EAN rows detected (EAN column: {labels.get(ean_col, ean_col)}, PID column: {labels.get(pid_col, pid_col)})")

if not listing_frames:
    st.error("None of the uploaded marketplace files could be read.")
    st.stop()

master = pd.concat(listing_frames, ignore_index=True)

# ---------------------------------------------------------------------------
# Content file -> EAN to Color No bridge
# ---------------------------------------------------------------------------

st.subheader("Content file")
content_df, content_labels = read_single_file_auto("Content file", content_file, CONTENT_HEADER_HINTS, "content")
if content_df is None:
    st.stop()

content_ean_guess = guess_column(content_df.columns, CONTENT_EAN_HINTS)
content_parent_guess = guess_column(content_df.columns, CONTENT_PARENT_HINTS)
with st.expander("⚙️ Fix Content file column detection"):
    cc1, cc2 = st.columns(2)
    with cc1:
        content_ean_col = st.selectbox(
            "EAN column", options=list(content_df.columns),
            index=list(content_df.columns).index(content_ean_guess) if content_ean_guess in content_df.columns else 0,
            format_func=lambda c: content_labels.get(c, c), key="content_ean_col",
        )
    with cc2:
        content_parent_col = st.selectbox(
            "Color No / Article No column", options=list(content_df.columns),
            index=list(content_df.columns).index(content_parent_guess) if content_parent_guess in content_df.columns else 0,
            format_func=lambda c: content_labels.get(c, c), key="content_parent_col",
        )

content_df["_EAN_KEY"] = content_df[content_ean_col].apply(lambda v: clean_id_str(v, normalize_keys))
content_df["_PARENT_KEY"] = content_df[content_parent_col].apply(lambda v: clean_id_str(v, normalize_keys))
ean_to_color = (
    content_df.dropna(subset=["_EAN_KEY"])
    .drop_duplicates(subset=["_EAN_KEY"], keep="first")
    .set_index("_EAN_KEY")["_PARENT_KEY"]
)

master["Color No"] = master["EAN"].apply(lambda v: clean_id_str(v, normalize_keys)).map(ean_to_color)

# ---------------------------------------------------------------------------
# ZeCom tracker -> which Color Nos are currently valid/active
# ---------------------------------------------------------------------------

zecom_valid_set = None
if zecom_file is not None:
    st.subheader("ZeCom tracker")
    zecom_df, zecom_labels = read_single_file_auto("ZeCom tracker", zecom_file, ZECOM_HEADER_HINTS, "zecom", default_sheet_hint=region)
    if zecom_df is not None:
        zecom_parent_guess = guess_column(zecom_df.columns, ZECOM_PARENT_HINTS)
        with st.expander("⚙️ Fix ZeCom join-key column detection"):
            zecom_parent_col = st.selectbox(
                "PIM_Article# / Color No / Style# column", options=list(zecom_df.columns),
                index=list(zecom_df.columns).index(zecom_parent_guess) if zecom_parent_guess in zecom_df.columns else 0,
                format_func=lambda c: zecom_labels.get(c, c), key="zecom_parent_col",
            )
        zecom_keys = zecom_df[zecom_parent_col].apply(lambda v: clean_id_str(v, normalize_keys))
        zecom_valid_set = set(zecom_keys.dropna().unique())

if zecom_valid_set is not None:
    master[f"In ZeCom ({region})"] = master["Color No"].apply(
        lambda v: "Yes" if (pd.notna(v) and v in zecom_valid_set) else ("N/A (no Color No)" if pd.isna(v) else "No")
    )
else:
    master[f"In ZeCom ({region})"] = "ZeCom file not uploaded"

# ---------------------------------------------------------------------------
# Inventory snapshots -> per-period stock + dead-stock flag
# ---------------------------------------------------------------------------

inventory_periods = [("Jan", inv_jan), ("Jun", inv_jun), ("Dec", inv_dec), ("Today", inv_now)]
active_periods = [(label, f) for label, f in inventory_periods if f is not None]

stock_cols = []
if active_periods:
    st.subheader("Inventory snapshots")
    for label, f in active_periods:
        df, labels = read_single_file_auto(f"{label} inventory", f, INVENTORY_HEADER_HINTS, f"inv_{label}")
        if df is None:
            continue
        ean_guess = guess_column(df.columns, INVENTORY_EAN_HINTS)
        stock_guess = guess_column(df.columns, INVENTORY_STOCK_HINTS)
        with st.expander(f"⚙️ Fix {label} inventory column detection"):
            c1, c2 = st.columns(2)
            with c1:
                inv_ean_col = st.selectbox(
                    f"{label} — EAN column", options=list(df.columns),
                    index=list(df.columns).index(ean_guess) if ean_guess in df.columns else 0,
                    format_func=lambda c: labels.get(c, c), key=f"inv_{label}_ean_col",
                )
            with c2:
                inv_stock_col = st.selectbox(
                    f"{label} — Stock column", options=list(df.columns),
                    index=list(df.columns).index(stock_guess) if stock_guess in df.columns else 0,
                    format_func=lambda c: labels.get(c, c), key=f"inv_{label}_stock_col",
                )
        keyed = df[[inv_ean_col, inv_stock_col]].copy()
        keyed["_EAN_KEY"] = keyed[inv_ean_col].apply(lambda v: clean_id_str(v, normalize_keys))
        keyed["_STOCK_NUM"] = pd.to_numeric(keyed[inv_stock_col], errors="coerce").fillna(0)
        lookup = keyed.dropna(subset=["_EAN_KEY"]).groupby("_EAN_KEY")["_STOCK_NUM"].sum()

        col_name = f"Stock_{label}"
        master[col_name] = master["EAN"].apply(lambda v: clean_id_str(v, normalize_keys)).map(lookup).fillna(0)
        stock_cols.append(col_name)
        st.caption(f"✓ {label} inventory: {len(lookup)} unique EANs read (EAN: {labels.get(inv_ean_col, inv_ean_col)}, Stock: {labels.get(inv_stock_col, inv_stock_col)})")

if stock_cols:
    master["No Stock All Year"] = (master[stock_cols] == 0).all(axis=1)
    periods_checked = ", ".join(label for label, f in active_periods)
    master["No Stock All Year"] = master["No Stock All Year"].map({True: f"⚠ Flagged (0 stock in: {periods_checked})", False: ""})
else:
    master["No Stock All Year"] = "No inventory files uploaded"

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Master output")

output_cols = ["EAN", "Marketplace", "Marketplace PID", "Color No", f"In ZeCom ({region})"] + stock_cols + ["No Stock All Year"]
final_df = master[output_cols].copy()

total_rows = len(final_df)
no_color_match = int(final_df["Color No"].isna().sum())
flagged_dead = int(final_df["No Stock All Year"].astype(str).str.startswith("⚠").sum()) if stock_cols else 0
not_in_zecom = int((final_df[f"In ZeCom ({region})"] == "No").sum()) if zecom_valid_set is not None else 0

s1, s2, s3, s4 = st.columns(4)
s1.metric("Total EAN-marketplace rows", total_rows)
s2.metric("EAN not mapped to Color No", no_color_match)
s3.metric("Color No not in ZeCom", not_in_zecom)
s4.metric("Flagged: no stock all year", flagged_dead)

st.dataframe(final_df.head(30), use_container_width=True, height=400)

buf = io.BytesIO()

def _safe_sheet_name(name: str) -> str:
    # Excel sheet names: max 31 chars, no : \ / ? * [ ]
    cleaned = re.sub(r"[:\\/?*\[\]]", "-", str(name))
    return cleaned[:31]

with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    final_df.to_excel(writer, index=False, sheet_name="All Marketplaces")
    for mp_name in final_df["Marketplace"].dropna().unique():
        mp_df_out = final_df[final_df["Marketplace"] == mp_name]
        mp_df_out.to_excel(writer, index=False, sheet_name=_safe_sheet_name(mp_name))
buf.seek(0)

st.download_button(
    "⬇️ Download master mapping (.xlsx)",
    data=buf,
    file_name=f"EAN_Master_Mapping_{region}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
