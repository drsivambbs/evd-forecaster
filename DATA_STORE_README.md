# EVD Snapshot Store (Firestore)

Continuously-updated case-count store that feeds **Step 1** of the EVD
Forecaster. Built as a **separate** entry app; the forecasting `app.py` only
ever reads from it.

## What got provisioned

| Thing | Value |
|---|---|
| Firebase / GCP project | `evd-snapshot-store` |
| Firestore | Native mode, region `asia-south1` (Mumbai), **free tier** |
| Billing | **Not linked** — billing account hit Google's project-count quota. Not needed; the free Spark tier covers this workload. |
| Service account | `evd-snapshot-writer@evd-snapshot-store.iam.gserviceaccount.com` (`roles/datastore.user`) |
| Key file | `.secrets/firebase-service-account.json` (git-ignored — **do not commit**) |

## Files

- `firestore_store.py` — shared access layer. **Both** apps import it.
  - `add_snapshot(...)`, `add_snapshots_bulk(...)`, `delete_snapshot(id)`
  - `load_snapshots(location=None)` — every version, for the management UI
  - `load_for_step1(location, value_type, as_of=None)` — the shaped frame
    Step 1 expects, with as-of reduction applied
- `data_entry.py` — standalone Streamlit write/admin UI.

## Run the data-entry app

```powershell
cd "C:\Users\Windows\Documents\Ebola Modelling\forecast-tool"
streamlit run data_entry.py
```

Entry flow: **Location → Source → Value type → Confirmed / Suspected / Deaths**.
Three tabs: *Add one*, *Bulk add* (one bulletin, many dates), *Browse / manage*.

## Schema (`snapshots` collection)

One document per reading: `location`, `event_date`, `value_type`
(`cumulative`|`incidence`), `confirmed`, `suspected`, `deaths`, `source_type`
(`WHO DON`|`Africa CDC`|`Other`), `source_name`, `source_url`, `as_of_date`,
`note`, `ingested_at`.

**Overwrite key (document id)** = `sha1(location | value_type | event_date |
source_name | as_of_date)` — **Decision A**: re-uploading the *same bulletin's*
reading overwrites the previous one (corrections), while a *different* bulletin
revising an old date is kept as a new version. That preserves the as-of trail
for backtesting.

## Wiring into Step 1 (app.py) — next step, not yet done

Add a third input method alongside "Manual entry" / "CSV upload":

```python
import firestore_store as fss
# value_type is "cumulative" or "incidence" from the existing radio
df = fss.load_for_step1(location, value_type, as_of=chosen_date)
# df columns already match VALUE_COLS + 'source', feed straight into
# interpolate_from_cumulative(df)  or  expand_incidence(df)
```
