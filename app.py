"""
PUMA EAN Master Mapper
========================
Builds one master reference table across all your listed EANs:

    EAN  |  Marketplace  |  Marketplace PID  |  Color No  |  In ZeCom?  |  Stock (Jan/Jun/Dec/Today)  |  No Stock All Year?  |  Recommend Delete

Inputs:
  1. Marketplace listing files — Lazada / Shopee (ZIP) / TikTok / Zalora — as many
     as you have for this region. Each just needs an EAN and the marketplace's own
     Product ID column; everything else in the file is ignored.
  2. Content file (EAN -> Color No) — the bridge, since marketplace files don't
     carry Color No directly.
  3. ZeCom tracker for the selected region (same file works whether it's the
     combined SG+MY workbook or the standalone PH one — the right sheet is picked
     automatically) — used to flag whether each Color No is actually active/valid
     in ZeCom right now, not to pull pricing.
  4. Up to 4 inventory snapshots (Jan / Jun / Dec / Today). Any EAN with zero (or
     missing) stock in every snapshot you provide gets flagged "No stock all year"
     — flagged, not deleted, so you can review before actually delisting anything.

Shared files: the Content file and ZeCom tracker change twice a week and are used
by everyone, so instead of re-uploading them every run, an admin uploads the
current version once (sidebar -> "Admin: update shared files") and it's reused
automatically until the next update. Marketplace listing files and inventory
snapshots stay per-run uploads, since those are specific to whoever's checking.

Assumption made explicit: this run is scoped to ONE region/country per session
(matching how the pricing mapper works) — e.g. all 4 marketplace files here would
be Lazada MY + Shopee MY + TikTok MY + Zalora MY together, not multiple countries
in one run. If you actually need multiple countries combined in a single output,
say so and this can be restructured.
"""

import io
import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="EAN Master Mapper", layout="wide")

# ---------------------------------------------------------------------------
# Shared file storage (Content file + ZeCom tracker, updated twice a week)
# ---------------------------------------------------------------------------
# NOTE on Streamlit Community Cloud free tier: this folder lives on the app's
# local disk, which is wiped whenever the app restarts (redeploy, or waking up
# from sleep after inactivity). That's fine given these files are refreshed a
# couple of times a week anyway — just re-upload via the admin panel if you
# notice the "last updated" timestamp is missing/stale after a restart.

SHARED_DIR = Path(__file__).parent / "shared_data"
SHARED_DIR.mkdir(exist_ok=True)


def _shared_paths(kind: str):
    return SHARED_DIR / f"{kind}_file.bin", SHARED_DIR / f"{kind}_meta.json"


def load_shared_file(kind: str):
    data_path, meta_path = _shared_paths(kind)
    if not data_path.exists() or not meta_path.exists():
        return None
    try:
        meta = json.loads(meta_path.read_text())
        return {
            "bytes": data_path.read_bytes(),
            "filename": meta["filename"],
            "uploaded_at": meta["uploaded_at"],
        }
    except Exception:
        return None


def save_shared_file(kind: str, uploaded_file) -> dict:
    data_path, meta_path = _shared_paths(kind)
    data_path.write_bytes(uploaded_file.getvalue())
    meta = {
        "filename": uploaded_file.name,
        "uploaded_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
    }
    meta_path.write_text(json.dumps(meta))
    return meta


class _StoredFile:
    """Wraps bytes read back off disk so it can be passed through the same
    detect_file_structure()/read_*() pipeline as a fresh st.file_uploader
    result, which only ever calls .getvalue() and reads .name."""

    def __init__(self, data: bytes, name: str):
        self._data = data
        self.name = name

    def getvalue(self) -> bytes:
        return self._data


# ---------------------------------------------------------------------------
# Shared low-level file reading (same approach as the ZeCom pricing mapper)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def list_sheets(file_bytes, filename):
    attempts = []
    if filename.lower().endswith(".csv"):
        return "csv", ["(csv)"], ["Detected .csv extension"]
    for engine in ["calamine", "openpyxl", "xlrd"]:
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


def _read_excel_any_engine(file_bytes, sheet_name, header, nrows=None, usecols=None):
    """
    Try engines in order for the ACTUAL data read, independent of whichever
    engine list_sheets() used to enumerate sheet names. This matters because
    some files (confirmed: real Shopee 'mass update' exports) can have their
    sheet names listed fine by openpyxl, but fail with a ValueError when
    openpyxl actually reads the data — it strictly validates worksheet view
    properties (e.g. the frozen-pane 'activePane' attribute) and rejects
    files where an export tool wrote a non-standard value for it. calamine
    doesn't do this strict validation and reads such files fine.
    `usecols`, when given, limits which columns pandas actually loads — this
    is the main memory lever for wide/tall files where only 1-2 columns are
    actually needed (e.g. EAN + PID out of a 16-column, 120k-row export).
    Returns (df, engine_used, attempts_log).
    """
    attempts = []
    for engine in ["calamine", "openpyxl", "xlrd"]:
        try:
            df = pd.read_excel(
                io.BytesIO(file_bytes), sheet_name=sheet_name, header=header,
                dtype=str, engine=engine, nrows=nrows, usecols=usecols,
            )
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


@st.cache_data(show_spinner=False)
def read_full_narrow(file_bytes, filename, engine, sheet_name, header_row, usecols):
    """
    Same as read_full, but only loads the given columns (by name). This is the
    main memory fix for this app: it uploads up to 4 marketplace files + a
    Content file + a ZeCom tracker + up to 4 inventory files in one session,
    but each of those only ever needs 1-2 columns out of however many the
    source file has. Caching a 2-column result instead of the full 16-111
    column file is the difference between this staying well under Streamlit
    Cloud's memory ceiling and silently crashing under it.
    """
    if filename.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes), header=header_row, dtype=str, usecols=usecols)
    elif engine == "html":
        tables = pd.read_html(io.BytesIO(file_bytes), header=header_row)
        df = tables[0][usecols]
    else:
        df, used_engine, attempts = _read_excel_any_engine(file_bytes, sheet_name, header_row, usecols=usecols)
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
                for list_engine in ["calamine", "openpyxl", "xlrd"]:
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


@st.cache_data(show_spinner=False)
def read_zip_combined_narrow(file_bytes, header_hints, ean_hints, pid_hints):
    """
    Same idea as read_zip_combined, but for each inner file: detect its header,
    guess EAN/PID from that file's own column names, and load ONLY those 2
    columns (standardized to 'EAN'/'PID') before concatenating. A 24-file
    Shopee ZIP previously meant holding 24 full-width dataframes in memory at
    once before ever trimming down to what's needed — this keeps peak memory
    to roughly 2 columns x total rows, regardless of how wide each source file is.
    Returns (combined_df, per_file_log).
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
                for list_engine in ["calamine", "openpyxl", "xlrd"]:
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

                header_vals = preview.iloc[hdr].tolist() if hdr < len(preview) else []
                col_names = []
                for i, h in enumerate(header_vals):
                    h_str = str(h).strip() if pd.notna(h) and str(h).strip().lower() not in ("", "nan", "none") else None
                    col_names.append(h_str or f"Unnamed: {i}")
                col_names = _dedupe_like_pandas(col_names)

                ean_col = guess_column(col_names, ean_hints)
                pid_col = guess_column(col_names, pid_hints)
                if ean_col is None:
                    per_file_log.append((name, hdr, None, "could not find an EAN/SKU column"))
                    continue
                wanted = list(dict.fromkeys([c for c in [ean_col, pid_col] if c]))

                df, engine, attempts = _read_excel_any_engine(inner_bytes, sheet, hdr, usecols=wanted)
                if df is None:
                    per_file_log.append((name, hdr, None, "all engines failed: " + "; ".join(attempts)))
                    continue
                df.columns = [str(c).strip() for c in df.columns]
                df = df.dropna(axis=0, how="all")
                df = strip_template_subheader_rows(df)
                rename_map = {ean_col: "EAN"}
                if pid_col:
                    rename_map[pid_col] = "PID"
                df = df.rename(columns=rename_map)
                if "PID" not in df.columns:
                    df["PID"] = None
                frames.append(df[["EAN", "PID"]])
                per_file_log.append((name, hdr, df.shape, None))
            except Exception as e:
                per_file_log.append((name, None, None, str(e)))

    if not frames:
        return None, per_file_log

    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined = combined.dropna(axis=0, how="all")
    return combined, per_file_log


def _dedupe_like_pandas(names):
    """Mirror pandas' own column-dedup convention (Col, Col.1, Col.2, ...) so
    column names guessed from a cheap preview match what a real pandas read
    with usecols=[...] will actually produce."""
    seen = {}
    result = []
    for n in names:
        if n not in seen:
            seen[n] = 0
            result.append(n)
        else:
            seen[n] += 1
            result.append(f"{n}.{seen[n]}")
    return result


def detect_file_structure(label, uploaded_file, header_hints, key_prefix, default_sheet_hint=None):
    """
    Cheap structure detection ONLY — sheet, header row, and column names from a
    small preview. Does NOT load the full file, so this stays fast and light
    regardless of how large the actual file is. The caller decides which 1-2
    columns it actually needs and loads only those via read_full_narrow.
    Returns (file_bytes, engine, sheet_name, header_row, column_names, labels)
    or (None, None, None, None, None, None) if the file can't be read.
    """
    if uploaded_file is None:
        return None, None, None, None, None, None
    file_bytes = uploaded_file.getvalue()
    engine, sheet_names, attempt_log = list_sheets(file_bytes, uploaded_file.name)
    if engine is None:
        st.error(_read_error_message(label, attempt_log))
        return None, None, None, None, None, None

    sheet_name = sheet_names[0]
    if len(sheet_names) > 1 and default_sheet_hint:
        for s in sheet_names:
            if default_sheet_hint.lower() == str(s).lower():
                sheet_name = s
                break

    with st.spinner(f"Scanning {label}…"):
        raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name)
        auto_header_row = find_header_row(raw, header_hints)

    with st.expander(f"⚙️ Fix detection for {label} (only open if something looks wrong)"):
        if len(sheet_names) > 1:
            sheet_name = st.selectbox("Sheet", options=sheet_names, index=sheet_names.index(sheet_name), key=f"{key_prefix}_sheet")
            raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name)
            auto_header_row = find_header_row(raw, header_hints)
        st.dataframe(raw.head(12), width="stretch", height=200)
        header_row = st.number_input("Header row (0 = first row)", min_value=0, max_value=500, value=int(auto_header_row), key=f"{key_prefix}_header_row")
    header_row = int(header_row)

    if header_row >= len(raw):
        raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name, nrows=header_row + 10)

    header_vals = raw.iloc[header_row].tolist() if header_row < len(raw) else []
    col_names = []
    for i, h in enumerate(header_vals):
        h_str = str(h).strip() if pd.notna(h) and str(h).strip().lower() not in ("", "nan", "none") else None
        col_names.append(h_str or f"Unnamed: {i}")
    col_names = _dedupe_like_pandas(col_names)

    banner_vals = raw.iloc[header_row - 1].tolist() if header_row > 0 else [None] * len(col_names)
    labels = build_label_map(col_names, banner_vals)

    return file_bytes, engine, sheet_name, header_row, col_names, labels


def load_narrow(file_bytes, filename, engine, sheet_name, header_row, wanted_cols):
    """Load ONLY the given columns for the full file, then apply the same
    blank-row / template-subheader-row cleanup as the wide path."""
    df = read_full_narrow(file_bytes, filename, engine, sheet_name, header_row, list(dict.fromkeys(wanted_cols)))
    df = df.dropna(axis=0, how="all")
    df = strip_template_subheader_rows(df)
    return df


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
    region = st.selectbox("Region (for this run's marketplace + ZeCom sheet)", ["MY", "PH", "SG"])

    st.header("2. Marketplace listing files")
    st.caption("Upload whichever you have for this region — all optional, combined into one output.")
    lazada_file = st.file_uploader("Lazada", type=["xlsx", "xls", "csv"], key="lazada_up")
    shopee_file = st.file_uploader("Shopee (export ZIP)", type=["zip"], key="shopee_up")
    tiktok_file = st.file_uploader("TikTok Shop", type=["xlsx", "xls", "csv"], key="tiktok_up")
    zalora_file = st.file_uploader("Zalora", type=["xlsx", "xls", "csv"], key="zalora_up")

    st.header("3. Shared files: Content & ZeCom")
    st.caption("These update twice a week and are shared across everyone using this app — no need to re-upload unless you're the one refreshing them.")

    content_shared = load_shared_file("content")
    zecom_shared = load_shared_file("zecom")

    if content_shared:
        st.success(f"📄 Content file: **{content_shared['filename']}**\n\nUpdated {content_shared['uploaded_at']}")
    else:
        st.warning("📄 No Content file shared yet.")

    if zecom_shared:
        st.success(f"📄 ZeCom tracker: **{zecom_shared['filename']}**\n\nUpdated {zecom_shared['uploaded_at']}")
    else:
        st.warning("📄 No ZeCom tracker shared yet.")

    with st.expander("🔒 Admin: update shared Content / ZeCom files"):
        admin_password = None
        try:
            admin_password = st.secrets.get("admin_password")
        except Exception:
            admin_password = None

        pw_ok = True
        if admin_password:
            pw_input = st.text_input("Admin password", type="password", key="admin_pw")
            pw_ok = pw_input == admin_password
            if pw_input and not pw_ok:
                st.error("Wrong password.")
        else:
            st.caption("⚠ No admin password set (add `admin_password` to Streamlit secrets to lock this down) — anyone can currently update these files.")

        new_content = st.file_uploader("Replace Content file", type=["xlsx", "xls", "csv"], key="admin_content_up", disabled=not pw_ok)
        if new_content is not None and pw_ok:
            meta = save_shared_file("content", new_content)
            st.success(f"Content file updated: {meta['filename']}")
            st.rerun()

        new_zecom = st.file_uploader("Replace ZeCom tracker", type=["xlsx", "xls", "csv"], key="admin_zecom_up", disabled=not pw_ok)
        if new_zecom is not None and pw_ok:
            meta = save_shared_file("zecom", new_zecom)
            st.success(f"ZeCom tracker updated: {meta['filename']}")
            st.rerun()

    st.header("4. Inventory snapshots (optional)")
    st.caption("Upload whichever you have — the dead-stock flag only checks the periods you provide.")
    inv_jan = st.file_uploader("January inventory", type=["xlsx", "xls", "csv"], key="inv_jan")
    inv_jun = st.file_uploader("June inventory", type=["xlsx", "xls", "csv"], key="inv_jun")
    inv_dec = st.file_uploader("December inventory", type=["xlsx", "xls", "csv"], key="inv_dec")
    inv_now = st.file_uploader("Current / today inventory", type=["xlsx", "xls", "csv"], key="inv_now")

    st.header("5. Options")
    normalize_keys = st.checkbox("Normalize join keys (trim spaces, uppercase, strip stray .0)", value=True)

content_file = _StoredFile(content_shared["bytes"], content_shared["filename"]) if content_shared else None
zecom_file = _StoredFile(zecom_shared["bytes"], zecom_shared["filename"]) if zecom_shared else None

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
    st.info("No Content file has been shared yet. Ask an admin to upload one via **🔒 Admin: update shared files** in the sidebar.")
    st.stop()

# ---------------------------------------------------------------------------
# Read marketplace files -> one long [EAN, Marketplace, Marketplace PID] table
# ---------------------------------------------------------------------------

st.subheader("Marketplace listings")
listing_frames = []
for name, (f, is_zip) in marketplace_uploads.items():
    if f is None:
        continue

    if is_zip:
        file_bytes = f.getvalue()
        with st.spinner(f"Unzipping and combining {name}…"):
            df, per_file_log = read_zip_combined_narrow(
                file_bytes, MARKETPLACE_HEADER_HINTS, MARKETPLACE_EAN_HINTS[name], MARKETPLACE_PID_HINTS[name]
            )
        if df is None:
            st.error(f"Could not find any readable Excel files inside the ZIP for {name}.")
            for fname, hdr, shape, err in per_file_log:
                if err:
                    st.caption(f"- {fname}: failed — {err}")
            continue
        with st.expander(f"📦 {name} — {len(per_file_log)} file(s) found in ZIP"):
            for fname, hdr, shape, err in per_file_log:
                st.caption(f"- {fname}: ⚠ skipped ({err})" if err else f"- {fname}: header row {hdr}, {shape[0]} rows × {shape[1]} cols")
            st.caption(f"Combined: {df.shape[0]} rows total.")
        sub = pd.DataFrame({
            "EAN": df["EAN"].apply(lambda v: clean_id_str(v, normalize_keys)),
            "Marketplace": name,
            "Marketplace PID": df["PID"].apply(lambda v: clean_id_str(v, False)),
        })
        sub = sub.dropna(subset=["EAN"])
        listing_frames.append(sub)
        st.caption(f"✓ {name}: {len(sub)} listed EAN rows detected")
        continue

    file_bytes, engine, sheet_name, header_row, col_names, labels = detect_file_structure(
        f"{name} listing file", f, MARKETPLACE_HEADER_HINTS, f"mp_{name}"
    )
    if file_bytes is None:
        continue

    ean_guess = guess_column(col_names, MARKETPLACE_EAN_HINTS[name])
    pid_guess = guess_column(col_names, MARKETPLACE_PID_HINTS[name])
    with st.expander(f"⚙️ Fix {name} column detection"):
        c1, c2 = st.columns(2)
        with c1:
            ean_col = st.selectbox(
                f"{name} — EAN / Seller SKU column", options=col_names,
                index=col_names.index(ean_guess) if ean_guess in col_names else 0,
                format_func=lambda c: labels.get(c, c), key=f"{name}_ean_col",
            )
        with c2:
            pid_col = st.selectbox(
                f"{name} — Product ID column", options=col_names,
                index=col_names.index(pid_guess) if pid_guess in col_names else 0,
                format_func=lambda c: labels.get(c, c), key=f"{name}_pid_col",
            )

    with st.spinner(f"Loading {name}…"):
        df = load_narrow(file_bytes, f.name, engine, sheet_name, header_row, [ean_col, pid_col])

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
content_bytes, content_engine, content_sheet, content_header_row, content_col_names, content_labels = detect_file_structure(
    "Content file", content_file, CONTENT_HEADER_HINTS, "content"
)
if content_bytes is None:
    st.stop()

content_ean_guess = guess_column(content_col_names, CONTENT_EAN_HINTS)
content_parent_guess = guess_column(content_col_names, CONTENT_PARENT_HINTS)
with st.expander("⚙️ Fix Content file column detection"):
    cc1, cc2 = st.columns(2)
    with cc1:
        content_ean_col = st.selectbox(
            "EAN column", options=content_col_names,
            index=content_col_names.index(content_ean_guess) if content_ean_guess in content_col_names else 0,
            format_func=lambda c: content_labels.get(c, c), key="content_ean_col",
        )
    with cc2:
        content_parent_col = st.selectbox(
            "Color No / Article No column", options=content_col_names,
            index=content_col_names.index(content_parent_guess) if content_parent_guess in content_col_names else 0,
            format_func=lambda c: content_labels.get(c, c), key="content_parent_col",
        )

with st.spinner("Loading Content file…"):
    content_df = load_narrow(content_bytes, content_file.name, content_engine, content_sheet, content_header_row, [content_ean_col, content_parent_col])

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
    zecom_bytes, zecom_engine, zecom_sheet, zecom_header_row, zecom_col_names, zecom_labels = detect_file_structure(
        "ZeCom tracker", zecom_file, ZECOM_HEADER_HINTS, "zecom", default_sheet_hint=region
    )
    if zecom_bytes is not None:
        zecom_parent_guess = guess_column(zecom_col_names, ZECOM_PARENT_HINTS)
        with st.expander("⚙️ Fix ZeCom join-key column detection"):
            zecom_parent_col = st.selectbox(
                "PIM_Article# / Color No / Style# column", options=zecom_col_names,
                index=zecom_col_names.index(zecom_parent_guess) if zecom_parent_guess in zecom_col_names else 0,
                format_func=lambda c: zecom_labels.get(c, c), key="zecom_parent_col",
            )
        with st.spinner("Loading ZeCom tracker…"):
            zecom_df = load_narrow(zecom_bytes, zecom_file.name, zecom_engine, zecom_sheet, zecom_header_row, [zecom_parent_col])
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
        inv_bytes, inv_engine, inv_sheet, inv_header_row, inv_col_names, inv_labels = detect_file_structure(
            f"{label} inventory", f, INVENTORY_HEADER_HINTS, f"inv_{label}"
        )
        if inv_bytes is None:
            continue
        ean_guess = guess_column(inv_col_names, INVENTORY_EAN_HINTS)
        stock_guess = guess_column(inv_col_names, INVENTORY_STOCK_HINTS)
        with st.expander(f"⚙️ Fix {label} inventory column detection"):
            c1, c2 = st.columns(2)
            with c1:
                inv_ean_col = st.selectbox(
                    f"{label} — EAN column", options=inv_col_names,
                    index=inv_col_names.index(ean_guess) if ean_guess in inv_col_names else 0,
                    format_func=lambda c: inv_labels.get(c, c), key=f"inv_{label}_ean_col",
                )
            with c2:
                inv_stock_col = st.selectbox(
                    f"{label} — Stock column", options=inv_col_names,
                    index=inv_col_names.index(stock_guess) if stock_guess in inv_col_names else 0,
                    format_func=lambda c: inv_labels.get(c, c), key=f"inv_{label}_stock_col",
                )
        with st.spinner(f"Loading {label} inventory…"):
            df = load_narrow(inv_bytes, f.name, inv_engine, inv_sheet, inv_header_row, [inv_ean_col, inv_stock_col])
        keyed = df[[inv_ean_col, inv_stock_col]].copy()
        keyed["_EAN_KEY"] = keyed[inv_ean_col].apply(lambda v: clean_id_str(v, normalize_keys))
        keyed["_STOCK_NUM"] = pd.to_numeric(keyed[inv_stock_col], errors="coerce").fillna(0)
        lookup = keyed.dropna(subset=["_EAN_KEY"]).groupby("_EAN_KEY")["_STOCK_NUM"].sum()

        col_name = f"Stock_{label}"
        master[col_name] = master["EAN"].apply(lambda v: clean_id_str(v, normalize_keys)).map(lookup).fillna(0)
        stock_cols.append(col_name)
        st.caption(f"✓ {label} inventory: {len(lookup)} unique EANs read (EAN: {inv_labels.get(inv_ean_col, inv_ean_col)}, Stock: {inv_labels.get(inv_stock_col, inv_stock_col)})")

if stock_cols:
    master["No Stock All Year"] = (master[stock_cols] == 0).all(axis=1)
    periods_checked = ", ".join(label for label, f in active_periods)
    master["No Stock All Year"] = master["No Stock All Year"].map({True: f"⚠ Flagged (0 stock in: {periods_checked})", False: ""})
else:
    master["No Stock All Year"] = "No inventory files uploaded"

# ---------------------------------------------------------------------------
# Recommend Delete — combines all three checks into one actionable flag
# ---------------------------------------------------------------------------

zecom_col = f"In ZeCom ({region})"


def _delete_reasons(row):
    reasons = []
    if pd.isna(row["Color No"]):
        reasons.append("EAN not found in Content file")
    if row.get(zecom_col) == "No":
        reasons.append("Color No not found in ZeCom tracker")
    if stock_cols and str(row.get("No Stock All Year", "")).startswith("⚠"):
        reasons.append("No stock in any uploaded snapshot")
    return "; ".join(reasons)


master["Delete Reason"] = master.apply(_delete_reasons, axis=1)
master["Recommend Delete"] = master["Delete Reason"] != ""

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Master output")

output_cols = (
    ["EAN", "Marketplace", "Marketplace PID", "Color No", zecom_col]
    + stock_cols
    + ["No Stock All Year", "Recommend Delete", "Delete Reason"]
)
final_df = master[output_cols].copy()

total_rows = len(final_df)
no_color_match = int(final_df["Color No"].isna().sum())
flagged_dead = int(final_df["No Stock All Year"].astype(str).str.startswith("⚠").sum()) if stock_cols else 0
not_in_zecom = int((final_df[zecom_col] == "No").sum()) if zecom_valid_set is not None else 0
recommend_delete = int(final_df["Recommend Delete"].sum())

s1, s2, s3, s4, s5 = st.columns(5)
s1.metric("Total EAN-marketplace rows", total_rows)
s2.metric("EAN not mapped to Color No", no_color_match)
s3.metric("Color No not in ZeCom", not_in_zecom)
s4.metric("Flagged: no stock all year", flagged_dead)
s5.metric("Recommend delete", recommend_delete)

st.dataframe(final_df.head(30), width="stretch", height=400)

buf = io.BytesIO()


def _safe_sheet_name(name: str) -> str:
    # Excel sheet names: max 31 chars, no : \ / ? * [ ]
    cleaned = re.sub(r"[:\\/?*\[\]]", "-", str(name))
    return cleaned[:31]


with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    final_df.to_excel(writer, index=False, sheet_name="All Marketplaces")
    final_df[final_df["Recommend Delete"]].to_excel(writer, index=False, sheet_name="Recommend Delete")
    for mp_name in final_df["Marketplace"].dropna().unique():
        mp_df_out = final_df[final_df["Marketplace"] == mp_name]
        mp_df_out.to_excel(writer, index=False, sheet_name=_safe_sheet_name(mp_name))
buf.seek(0)

st.download_button(
    "⬇️ Download master mapping (.xlsx)",
    data=buf,
    file_name=f"EAN_Master_Mapping_{region}_{datetime.now().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
