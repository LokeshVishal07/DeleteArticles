"""
PUMA ZeCom Column Mapper (v3 — streamlined)
=============================================
Upload a MARKETPLACE file (Lazada / Shopee / Zalora / TikTok) that only has EAN
(Seller SKU) at the row level, and get back the SAME file with ZeCom pricing
(and any other ZeCom columns you choose) mapped in, plus an "Article No" column.

JOIN CHAIN (ZeCom has no EAN, so we bridge through Content):
  Marketplace EAN  --(Content file)-->  Color No / Article No  --(ZeCom file)-->  Selected columns

v3 changes:
  - Sheet / header-row / EAN-column / parent-key-column detection all run silently
    in the backend. Nothing needs to be picked by hand unless the auto-detection
    is actually wrong for a given file — in which case a small collapsed
    "Fix detection" expander is there as a safety net, but it's not part of the
    main flow anymore. The ONLY thing you choose each time is which ZeCom
    columns to map in.
  - Fewer live widgets in the main flow = fewer full Streamlit reruns = faster
    to use, on top of the memory fix from v2.
  - The output always includes a literal "Article No" column with the resolved
    join key, regardless of what that column was actually called in the source
    files (Color No / Style# / PIM Article# / STYLE# / etc).
  - Shopee marketplace files are uploaded as a ZIP (matching how Shopee exports
    multiple files per batch): the app unzips every .xlsx inside, auto-detects
    each one's header row, combines them into one table, and drops duplicated
    header rows and fully-empty rows that show up from combining multiple
    exports.
"""

import io
import re
import zipfile

import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="ZeCom Column Mapper", layout="wide")

# ---------------------------------------------------------------------------
# Low-level file reading — format sniffing + multi-engine fallback
# ---------------------------------------------------------------------------

def _looks_like_html(b: bytes) -> bool:
    head = b[:512].lstrip().lower()
    return head.startswith(b"<html") or head.startswith(b"<!doctype") or b"<table" in head[:2000].lower()


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes, filename):
    """Try each engine/format in turn just to enumerate sheet names."""
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
    engine list_sheets() used to enumerate sheet names. Some files (confirmed:
    real Shopee 'mass update' exports) list their sheet names fine with
    openpyxl but fail with a ValueError when openpyxl actually reads the data
    — it strictly validates worksheet view properties (e.g. the frozen-pane
    'activePane' attribute) and rejects files where an export tool wrote a
    non-standard value for it. calamine doesn't do this strict validation and
    reads such files fine.
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
    """Cheap, small read — just enough rows to detect the header row and show a preview."""
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
    """ONE full read, header applied natively by pandas (memory-efficient, single pass)."""
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
    """
    Some marketplace templates (confirmed on a real Lazada export) include 1-3
    metadata rows directly below the real header — 'Optional'/'Mandatory'
    markers and long field-description sentences — before actual data starts.
    Left in place, these get treated as real data rows (and can even get
    misread as the EAN for row 1). Detect and drop them.
    """
    drop_idx = []
    for i in range(min(max_check, len(df))):
        row = df.iloc[i]
        non_null = [str(v).strip() for v in row if pd.notna(v) and str(v).strip() != ""]
        if not non_null:
            continue  # blank rows are handled separately by dropna
        junk_hits = sum(
            1 for v in non_null
            if v.lower() in ("optional", "mandatory", "m", "o", "required", "n/a")
            or len(v) > 60  # long descriptive/instructional sentences
        )
        if junk_hits / len(non_null) >= 0.5:
            drop_idx.append(df.index[i])
        else:
            break  # stop at the first row that looks like real data
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


# ---------------------------------------------------------------------------
# Shopee: ZIP of multiple exports -> one combined table
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def read_shopee_zip(file_bytes, header_hints):
    """
    Unzip every .xlsx/.xls inside, auto-detect each file's header row, combine
    into one table, then drop fully-empty rows and rows that are actually a
    duplicated header (an artifact of combining several exports).
    Returns (combined_df, per_file_log) — combined_df is None if nothing readable was found.
    """
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

    # Drop rows that are actually a duplicated header row (artifact of combining
    # several files, e.g. a file's header block ending up parsed as a data row).
    col_set = {str(c).strip().lower() for c in combined.columns}

    def _looks_like_header(row):
        vals = [str(v).strip().lower() for v in row if pd.notna(v)]
        if not vals:
            return False
        hits = sum(1 for v in vals if v in col_set)
        return hits >= max(2, len(vals) // 2)

    header_like_mask = combined.apply(_looks_like_header, axis=1)
    combined = combined[~header_like_mask].reset_index(drop=True)
    combined = combined.dropna(axis=1, how="all")
    return combined, per_file_log


# ---------------------------------------------------------------------------
# Column-guessing hints
# ---------------------------------------------------------------------------

MARKETPLACE_EAN_COLUMN_HINTS = {
    "Lazada": ["sellersku", "seller sku"],
    "Shopee": ["seller sku", "sku reference no", "sku"],
    "Zalora": ["sellersku", "seller sku"],
    "TikTok Shop": ["seller sku"],
}
MARKETPLACE_HEADER_HINTS = ["sellersku", "seller sku", "sku", "product", "seller"]

CONTENT_EAN_HINTS = ["ean"]
CONTENT_PARENT_HINTS = ["color no", "article no", "colorno", "articleno", "style#", "style #"]
CONTENT_HEADER_HINTS = CONTENT_EAN_HINTS + CONTENT_PARENT_HINTS

ZECOM_PARENT_HINTS = [
    "pim article", "pim_article", "pim style",
    "article no", "articleno", "color no", "colorno",
    "style#", "style #",
]
ZECOM_HEADER_HINTS = ZECOM_PARENT_HINTS + ["price", "srp", "rrp", "md price"]


def guess_column(columns, hints):
    """Hint-priority order matters — see v2 notes. Checks exact + substring per hint before moving on."""
    cols_lower = {c: str(c).strip().lower() for c in columns}
    for h in hints:
        for c, cl in cols_lower.items():
            if cl == h:
                return c
        for c, cl in cols_lower.items():
            if h in cl:
                return c
    return None


# ---------------------------------------------------------------------------
# File reading pipeline — auto-detect everything, hide overrides in an expander
# ---------------------------------------------------------------------------

def _read_error_message(label, attempt_log):
    all_missing = all(("not installed" in a) or ("importerror" in a.lower()) for a in attempt_log)
    if all_missing:
        return (
            f"**Could not read {label} — but this isn't a problem with your file.**\n\n"
            "None of the Excel-reading packages are installed in the Python environment "
            "currently running this app:\n\n"
            + "\n".join(f"- {a}" for a in attempt_log)
            + "\n\n**Fix:** make sure `requirements.txt` is committed at the same repo path "
            "as `app.py`, then reboot the app from the Streamlit Cloud dashboard."
        )
    return (
        f"**Could not read {label}.** Tried multiple formats and none worked:\n\n"
        + "\n".join(f"- {a}" for a in attempt_log)
        + "\n\nIf this is a genuine Excel file, try re-saving it as .xlsx from Excel first."
    )


def configure_file_auto(label, uploaded_file, header_hints, key_prefix, default_sheet_hint=None, is_zip=False):
    """
    Full pipeline for one uploaded file, fully automatic by default:
    sheet + header row are auto-detected and never require interaction unless
    the (collapsed) "Fix detection" expander is opened and changed.
    Returns (df, labels) or (None, None).
    """
    if uploaded_file is None:
        return None, None

    file_bytes = uploaded_file.getvalue()

    if is_zip:
        with st.spinner(f"Unzipping and combining {label}…"):
            df, per_file_log = read_shopee_zip(file_bytes, header_hints)
        if df is None:
            st.error(f"Could not find any readable Excel files inside the ZIP for {label}.")
            for name, hdr, shape, err in per_file_log:
                if err:
                    st.caption(f"- {name}: failed — {err}")
            return None, None
        with st.expander(f"📦 {label} — {len(per_file_log)} file(s) found in ZIP"):
            for name, hdr, shape, err in per_file_log:
                if err:
                    st.caption(f"- {name}: ⚠ skipped ({err})")
                else:
                    st.caption(f"- {name}: header row {hdr}, {shape[0]} rows × {shape[1]} cols")
            st.caption(f"Combined: {df.shape[0]} rows × {df.shape[1]} cols after removing duplicate headers/empty rows.")
        labels = {c: c for c in df.columns}
        return df, labels

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

    with st.expander(f"⚙️ Fix detection for {label} (only open this if something looks wrong)"):
        if len(sheet_names) > 1:
            sheet_name = st.selectbox(
                "Sheet", options=sheet_names, index=sheet_names.index(sheet_name), key=f"{key_prefix}_sheet"
            )
            raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name)
            auto_header_row = find_header_row(raw, header_hints)
        st.dataframe(raw.head(12), width="stretch", height=200)
        header_row = st.number_input(
            "Header row (0 = first row)", min_value=0, max_value=500,
            value=int(auto_header_row), key=f"{key_prefix}_header_row",
        )
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
# Sidebar — setup
# ---------------------------------------------------------------------------

st.title("🔗 ZeCom Column Mapper")
st.caption(
    "Upload your marketplace export + Content file + ZeCom file. "
    "Pick which ZeCom columns to map in — everything else runs automatically."
)

with st.sidebar:
    st.header("1. Setup")
    marketplace = st.selectbox("Marketplace", ["Lazada", "Shopee", "Zalora", "TikTok Shop"])
    region = st.selectbox("Region", ["MY", "PH", "SG"])

    st.header("2. Upload files")
    if marketplace == "Shopee":
        mp_file = st.file_uploader("Marketplace file (Shopee) — upload the export ZIP", type=["zip"])
    else:
        mp_file = st.file_uploader(f"Marketplace file ({marketplace})", type=["xlsx", "xls", "csv"])
    content_file = st.file_uploader("Content file (EAN → Color No)", type=["xlsx", "xls", "csv"])
    zecom_file = st.file_uploader("ZeCom file (pricing etc.)", type=["xlsx", "xls", "csv"])

    st.header("3. Options")
    normalize_keys = st.checkbox(
        "Normalize join keys (trim spaces, uppercase, strip stray .0)",
        value=True,
        help="Turn on if rows aren't matching due to minor formatting differences between files.",
    )

if not (mp_file and content_file and zecom_file):
    st.info("Upload all three files in the sidebar to get started.")
    st.stop()

# ---------------------------------------------------------------------------
# Marketplace file — fully automatic
# ---------------------------------------------------------------------------

mp_df, mp_labels = configure_file_auto(
    f"Marketplace file ({marketplace})", mp_file, MARKETPLACE_HEADER_HINTS, "mp",
    is_zip=(marketplace == "Shopee"),
)
if mp_df is None:
    st.stop()

mp_ean_guess = guess_column(mp_df.columns, MARKETPLACE_EAN_COLUMN_HINTS[marketplace])
with st.expander("⚙️ Fix EAN / Seller SKU column detection"):
    mp_ean_col = st.selectbox(
        "EAN / Seller SKU column",
        options=list(mp_df.columns),
        index=list(mp_df.columns).index(mp_ean_guess) if mp_ean_guess in mp_df.columns else 0,
        format_func=lambda c: mp_labels.get(c, c),
        key="mp_ean_col",
    )

# ---------------------------------------------------------------------------
# Content file — fully automatic
# ---------------------------------------------------------------------------

content_df, content_labels = configure_file_auto("Content file", content_file, CONTENT_HEADER_HINTS, "content")
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
            "Color No / Article No / Style# column", options=list(content_df.columns),
            index=list(content_df.columns).index(content_parent_guess) if content_parent_guess in content_df.columns else 0,
            format_func=lambda c: content_labels.get(c, c), key="content_parent_col",
        )

# ---------------------------------------------------------------------------
# ZeCom file — sheet/header/parent-key automatic; column picker is the main control
# ---------------------------------------------------------------------------

st.subheader("Pick ZeCom columns to map in")
zecom_df, zecom_labels = configure_file_auto(
    "ZeCom file", zecom_file, ZECOM_HEADER_HINTS, "zecom", default_sheet_hint=region
)
if zecom_df is None:
    st.stop()

zecom_parent_guess = guess_column(zecom_df.columns, ZECOM_PARENT_HINTS)
with st.expander("⚙️ Fix ZeCom join-key column detection"):
    zecom_parent_col = st.selectbox(
        "PIM_Article# / Color No / Style# column in ZeCom file",
        options=list(zecom_df.columns),
        index=list(zecom_df.columns).index(zecom_parent_guess) if zecom_parent_guess in zecom_df.columns else 0,
        format_func=lambda c: zecom_labels.get(c, c),
        key="zecom_parent_col",
    )

other_zecom_cols = [c for c in zecom_df.columns if c != zecom_parent_col]
st.caption(
    "Tracker columns often repeat per campaign tier (BAU / Payday / Mega / Shopee-specific, etc). "
    "Each option shows its real Excel column letter and the campaign banner text above it."
)
selected_zecom_cols = st.multiselect(
    "ZeCom columns to map into the output",
    options=other_zecom_cols,
    format_func=lambda c: zecom_labels.get(c, c),
)

if not selected_zecom_cols:
    st.warning("Pick at least one ZeCom column above to see mapped results.")
    st.stop()

# ---------------------------------------------------------------------------
# Build the join
# ---------------------------------------------------------------------------

with st.spinner("Mapping…"):
    mp_df["_EAN_KEY"] = mp_df[mp_ean_col].apply(lambda v: clean_id_str(v, normalize_keys))
    content_df["_EAN_KEY"] = content_df[content_ean_col].apply(lambda v: clean_id_str(v, normalize_keys))
    content_df["_PARENT_KEY"] = content_df[content_parent_col].apply(lambda v: clean_id_str(v, normalize_keys))
    zecom_df["_PARENT_KEY"] = zecom_df[zecom_parent_col].apply(lambda v: clean_id_str(v, normalize_keys))

    dup_parents = zecom_df["_PARENT_KEY"].value_counts()
    dup_parents = set(dup_parents[dup_parents > 1].index) - {None}

    ean_to_parent = (
        content_df.dropna(subset=["_EAN_KEY"])
        .drop_duplicates(subset=["_EAN_KEY"], keep="first")
        .set_index("_EAN_KEY")["_PARENT_KEY"]
    )
    zecom_lookup = (
        zecom_df.dropna(subset=["_PARENT_KEY"])
        .drop_duplicates(subset=["_PARENT_KEY"], keep="first")
        .set_index("_PARENT_KEY")[selected_zecom_cols]
    )

    mp_df["_PARENT_KEY"] = mp_df["_EAN_KEY"].map(ean_to_parent)
    mapped = mp_df.join(zecom_lookup, on="_PARENT_KEY")
    mapped["Article No"] = mapped["_PARENT_KEY"]

    for c in selected_zecom_cols:
        mapped[c] = mapped[c].where(mapped["_PARENT_KEY"].notna(), "Not Available (no EAN→Article No match)")
        mapped[c] = mapped[c].where(
            ~(mapped["_PARENT_KEY"].notna() & mapped[c].isna()),
            "Not Available (Article No not in ZeCom)",
        )
    mapped["ZeCom_Duplicate_Flag"] = mapped["_PARENT_KEY"].apply(
        lambda k: "⚠ Multiple ZeCom entries for this Article No" if k in dup_parents else ""
    )

    output_label_map = {c: zecom_labels.get(c, c) for c in selected_zecom_cols}
    # Insert "Article No" right after the EAN/Seller SKU column for readability.
    base_cols = list(mp_df.columns.drop(["_EAN_KEY", "_PARENT_KEY"]))
    ean_pos = base_cols.index(mp_ean_col)
    output_cols = base_cols[: ean_pos + 1] + ["Article No"] + base_cols[ean_pos + 1 :] + selected_zecom_cols + ["ZeCom_Duplicate_Flag"]
    final_df = mapped[output_cols].copy()
    final_df = final_df.rename(columns=output_label_map)

total_rows = len(final_df)
no_content_match = int((mapped["_PARENT_KEY"].isna()).sum())
first_sel = selected_zecom_cols[0]
no_zecom_match = int(
    ((mapped["_PARENT_KEY"].notna()) & (mapped[first_sel].astype(str).str.startswith("Not Available"))).sum()
)
dup_hits = int((mapped["ZeCom_Duplicate_Flag"] != "").sum())

st.subheader("Mapped output")
s1, s2, s3, s4 = st.columns(4)
s1.metric("Total rows", total_rows)
s2.metric("EAN not found in Content", no_content_match)
s3.metric("Article No not found in ZeCom", no_zecom_match)
s4.metric("Duplicate ZeCom entries hit", dup_hits)

st.dataframe(final_df.head(20), width="stretch", height=350)

buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    final_df.to_excel(writer, index=False, sheet_name="Mapped Output")
buf.seek(0)

st.download_button(
    "⬇️ Download mapped file (.xlsx)",
    data=buf,
    file_name=f"{marketplace}_{region}_ZeCom_Mapped.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
