# EAN Master Mapper

Builds one master reference table across everything you have listed:

```
EAN | Marketplace | Marketplace PID | Color No | In ZeCom? | Stock (Jan/Jun/Dec/Today) | No Stock All Year?
```

## Inputs

1. **Marketplace listing files** — Lazada / Shopee (as a ZIP) / TikTok / Zalora.
   Upload as many as you have for this region; each only needs an EAN and the
   marketplace's own Product ID column, everything else is ignored. All get
   combined into one long table (one row per EAN-marketplace pair).
2. **One Content file** (EAN → Color No) — the bridge, since marketplace files
   don't carry Color No directly.
3. **One ZeCom tracker** for the selected region — same file works whether it's
   the combined SG+MY workbook or the standalone PH one; the right sheet is
   picked automatically based on the region you select. This is used to flag
   whether each Color No is *currently active in ZeCom*, not to pull pricing.
4. **Up to 4 inventory snapshots** (Jan / Jun / Dec / Today). Any EAN with zero
   (or missing) stock in *every* snapshot you provide gets flagged
   `⚠ Flagged (0 stock in: ...)` — flagged, not deleted, so you can review
   before actually delisting anything.

## Scope assumption

This is built for **one region per run** (matching how the pricing mapper
works) — e.g. all 4 marketplace files in one run would be Lazada MY + Shopee MY
+ TikTok MY + Zalora MY together, not multiple countries combined into a
single output. If you need multiple countries combined in one file, that would
need a small restructure — just ask.

## How it runs

Sheet detection, header row, EAN column, PID column, and the ZeCom join-key
column are all auto-detected. Each has a small collapsed "⚙️ Fix detection"
expander next to it as a safety net — you shouldn't normally need to open
these. Nothing else needs manual configuration.

## How to run

**Locally:**
```bash
pip install -r requirements.txt
streamlit run app.py
```

**On Streamlit Community Cloud:** push `app.py` + `requirements.txt` to a
GitHub repo (same folder), then deploy from share.streamlit.io.

## Performance fix (if you deployed an earlier version)

If your app was silently crashing (health check "connection reset by peer",
no clear error) after uploading several large real files — this was a real
issue, now fixed. Two problems, both in the file reader:

1. **Memory**: every file was cached at full width/height, even when only 1-2
   columns were ever used (e.g. EAN + Product ID out of a 16-column, 120k-row
   Lazada export; 1 column out of 111 in a PH ZeCom tracker). With up to 4
   marketplace files + Content + ZeCom + 4 inventory snapshots all cached at
   once, this added up well past Streamlit Cloud's free-tier memory ceiling.
   Fixed: files are now scanned cheaply for their structure first, then only
   the specific columns actually needed are loaded and cached.
2. **Speed**: pandas' `usecols` parameter is *slower* with the `openpyxl`
   engine at scale (confirmed: 37s → hung well past 60s on a 123k-row file),
   while `calamine` reads the same file in ~6s and its `usecols` behaves
   correctly. Engine order was switched to try `calamine` first for actual
   data reads (falling back to `openpyxl`/`xlrd` if unavailable) — this also
   happens to be the engine that correctly reads real Shopee exports, which
   `openpyxl` rejects due to an unrelated strict-validation bug.

## Notes

- Rows are never silently dropped: unmapped EANs, EANs whose Color No isn't in
  ZeCom, and dead-stock candidates are all flagged in dedicated columns rather
  than removed.
- Marketplace SKU/PID columns auto-detected per platform — override in the
  "Fix [Marketplace] column detection" expander if a template changes.
- Template metadata rows (Optional/Mandatory markers, long field-description
  rows some marketplaces insert under the header) are automatically stripped —
  confirmed against a real Lazada export during testing.
- Shopee ZIP handling: unzips every file inside, auto-detects each one's
  header, combines them, and strips duplicate header rows / empty rows that
  show up from combining multiple exports.
