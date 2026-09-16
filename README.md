# EAN Master Mapper

Builds one master reference table across everything you have listed:

```
EAN | Marketplace | Marketplace PID | Color No | In ZeCom? | Stock (Jan/Jun/Dec/Today) | No Stock All Year? | Recommend Delete
```

## Inputs

1. **Marketplace listing files** — Lazada / Shopee (as a ZIP) / TikTok / Zalora.
   Upload as many as you have for this region; each only needs an EAN and the
   marketplace's own Product ID column, everything else is ignored. All get
   combined into one long table (one row per EAN-marketplace pair).
2. **Shared Content file** (EAN → Color No) — the bridge, since marketplace files
   don't carry Color No directly.
3. **Shared ZeCom tracker** for the selected region — same file works whether
   it's the combined SG+MY workbook or the standalone PH one; the right sheet
   is picked automatically based on the region you select. This is used to
   flag whether each Color No is *currently active in ZeCom*, not to pull
   pricing.
4. **Up to 4 inventory snapshots** (Jan / Jun / Dec / Today), optional. Any EAN
   with zero (or missing) stock in *every* snapshot you provide gets flagged
   `⚠ Flagged (0 stock in: ...)` — flagged, not deleted, so you can review
   before actually delisting anything.

## Shared files (Content & ZeCom)

Content and ZeCom change twice a week and are the same for everyone using the
app, so they're no longer uploaded per run. Instead:

- The sidebar's **"3. Shared files: Content & ZeCom"** section shows what's
  currently loaded and when it was last updated.
- Whoever has the new version uploads it via **🔒 Admin: update shared
  Content / ZeCom files**. From then on, everyone using the app automatically
  gets that version — no re-upload needed until the next update.
- Optionally password-protect that admin panel: add an `admin_password` value
  to Streamlit secrets (see `.streamlit/secrets.toml.example`). Without it,
  anyone with the app link can update the shared files.

**Caveat (Streamlit Community Cloud free tier only):** the shared files live
on the app's local disk, which is wiped when the app restarts — which happens
automatically after a period of inactivity, and on every redeploy. Since
Content/ZeCom are refreshed a couple of times a week anyway, this is usually a
non-issue — just check the "last updated" line in the sidebar, and if it's
missing after a while, re-upload. If this becomes a real problem, the fix is
to move shared-file storage off local disk (e.g. a small database, cloud
storage, or committing the files into the git repo) — ask and this can be
restructured.

## Recommend Delete

Each EAN-marketplace row gets a combined flag if **any** of these are true:

- the EAN isn't found in the Content file at all,
- its Color No isn't found in the ZeCom tracker for the selected region, or
- (if you uploaded inventory snapshots) it had zero stock in every snapshot.

The reasons are listed per row in **Delete Reason**, and the download includes
a dedicated **"Recommend Delete"** sheet with just the flagged rows.

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

**On Streamlit Community Cloud:** push `app.py` + `requirements.txt` (+
`.gitignore`) to your GitHub repo (same folder), then deploy from
share.streamlit.io. To set the admin password, open the app's settings on
Streamlit Cloud → **Secrets**, and paste:
```toml
admin_password = "your-password-here"
```

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
