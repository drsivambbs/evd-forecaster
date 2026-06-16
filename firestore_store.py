"""
firestore_store.py — shared Firestore access layer for the EVD snapshot store.

Both data_entry.py (the write/admin UI) and app.py Step 1 (read only) import
from here. The forecasting app must ONLY ever call load_for_step1() /
load_snapshots(); it never writes.

Project : evd-snapshot-store   (Firestore Native, asia-south1, free tier)
Collection: `snapshots` — one document per reading:

    location      "DRC" | "Uganda"
    event_date    "YYYY-MM-DD"   date the counts refer to
    value_type    "cumulative" | "incidence"
    confirmed     int            cumulative_* OR new_* depending on value_type
    suspected     int
    deaths        int
    source_type   "WHO DON" | "Africa CDC" | "INSP SitRep" | "Other"  (category / filter)
    source_name   e.g. "WHO DON605"   the SPECIFIC bulletin
    source_url    link / DOI
    as_of_date    "YYYY-MM-DD"   publication date of that bulletin
    note          optional free text (e.g. "DON607 revised earlier count")
    ingested_at   server timestamp

OVERWRITE KEY (document id) = sha1(
    location | value_type | event_date | source_name | as_of_date )

Decision A (chosen by the user): the key includes the SPECIFIC bulletin
(source_name + as_of_date), not just the source_type channel. So:
  * re-uploading the same bulletin's reading  -> same id -> OVERWRITES
    (corrections replace the previous value); and
  * a later bulletin revising an old event_date -> different source_name /
    as_of_date -> NEW document, old one preserved -> as-of history for
    backtesting stays intact.
"""

from __future__ import annotations

import hashlib
import os
from datetime import date, datetime
from pathlib import Path

import pandas as pd

PROJECT_ID = "evd-snapshot-store"
COLLECTION = "snapshots"

LOCATIONS = ["DRC", "Uganda"]
SOURCE_TYPES = ["WHO DON", "Africa CDC", "INSP SitRep", "Other"]
VALUE_TYPES = ["cumulative", "incidence"]
MEASURES = ["confirmed", "suspected", "deaths"]

# Column names Step 1 consumes, by value_type (see app.py VALUE_COLS).
_STEP1_RENAME = {
    "cumulative": {
        "confirmed": "cumulative_confirmed",
        "suspected": "cumulative_suspected",
        "deaths": "cumulative_deaths",
    },
    "incidence": {
        "confirmed": "new_confirmed",
        "suspected": "new_suspected",
        "deaths": "new_deaths",
    },
}

_DEFAULT_KEY = Path(__file__).resolve().parent / ".secrets" / "firebase-service-account.json"

_client = None  # cached firestore.Client singleton


# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------
def get_client():
    """Lazily initialise firebase_admin and return a Firestore client.

    Credential resolution order:
      1. GOOGLE_APPLICATION_CREDENTIALS env var (a key-file path)
      2. st.secrets['firebase_service_account'] (dict, e.g. on Streamlit Cloud)
      3. the bundled .secrets/firebase-service-account.json key file
      4. Application Default Credentials (gcloud auth on this machine)
    """
    global _client
    if _client is not None:
        return _client

    import firebase_admin
    from firebase_admin import credentials, firestore

    if not firebase_admin._apps:
        cred = None
        env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

        # 2. Streamlit secrets (only if streamlit is importable and configured)
        if cred is None:
            try:
                import streamlit as st
                if "firebase_service_account" in st.secrets:
                    cred = credentials.Certificate(
                        dict(st.secrets["firebase_service_account"]))
            except Exception:
                pass

        # 1 / 3. key file on disk
        if cred is None:
            if env_path and Path(env_path).exists():
                cred = credentials.Certificate(env_path)
            elif _DEFAULT_KEY.exists():
                cred = credentials.Certificate(str(_DEFAULT_KEY))

        # 4. ADC fallback
        if cred is None:
            cred = credentials.ApplicationDefault()

        firebase_admin.initialize_app(cred, {"projectId": PROJECT_ID})

    _client = firestore.client()
    return _client


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _as_iso(d) -> str:
    """Coerce a date / datetime / string to a strict 'YYYY-MM-DD' string."""
    if isinstance(d, (datetime, date)):
        return d.strftime("%Y-%m-%d")
    return pd.to_datetime(str(d)).strftime("%Y-%m-%d")


def doc_id(location: str, value_type: str, event_date, source_name: str,
           as_of_date) -> str:
    """Deterministic document id = the overwrite key (Decision A)."""
    key = "|".join([
        str(location).strip(),
        str(value_type).strip(),
        _as_iso(event_date),
        str(source_name).strip().lower(),
        _as_iso(as_of_date),
    ])
    return hashlib.sha1(key.encode("utf-8")).hexdigest()


def _validate(location, value_type, source_type):
    if location not in LOCATIONS:
        raise ValueError(f"location must be one of {LOCATIONS}, got {location!r}")
    if value_type not in VALUE_TYPES:
        raise ValueError(f"value_type must be one of {VALUE_TYPES}, got {value_type!r}")
    if source_type not in SOURCE_TYPES:
        raise ValueError(f"source_type must be one of {SOURCE_TYPES}, got {source_type!r}")


# ---------------------------------------------------------------------------
# Writes
# ---------------------------------------------------------------------------
def add_snapshot(*, location: str, value_type: str, event_date,
                 confirmed: int, suspected: int, deaths: int,
                 source_type: str, source_name: str, source_url: str = "",
                 as_of_date, note: str = "") -> str:
    """Upsert a single reading. Returns the document id.

    Because the id is the overwrite key, calling this again with the same
    (location, value_type, event_date, source_name, as_of_date) REPLACES the
    previous document — exactly the correction behaviour requested.
    """
    from firebase_admin import firestore as _fs
    _validate(location, value_type, source_type)
    payload = {
        "location": location,
        "value_type": value_type,
        "event_date": _as_iso(event_date),
        "confirmed": int(confirmed),
        "suspected": int(suspected),
        "deaths": int(deaths),
        "source_type": source_type,
        "source_name": str(source_name).strip(),
        "source_url": str(source_url).strip(),
        "as_of_date": _as_iso(as_of_date),
        "note": str(note).strip(),
        "ingested_at": _fs.SERVER_TIMESTAMP,
    }
    did = doc_id(location, value_type, event_date, source_name, as_of_date)
    get_client().collection(COLLECTION).document(did).set(payload)  # overwrite
    return did


def add_snapshots_bulk(rows: pd.DataFrame, *, location: str, value_type: str,
                       source_type: str, source_name: str, source_url: str = "",
                       as_of_date, note: str = "") -> int:
    """Upsert many per-date rows that share one bulletin's metadata.

    `rows` must have columns: event_date, confirmed, suspected, deaths.
    Same overwrite semantics as add_snapshot, applied per event_date.
    Returns the number of documents written.
    """
    from firebase_admin import firestore as _fs
    _validate(location, value_type, source_type)
    need = {"event_date", "confirmed", "suspected", "deaths"}
    missing = need - set(rows.columns)
    if missing:
        raise ValueError(f"bulk rows missing columns: {sorted(missing)}")

    client = get_client()
    col = client.collection(COLLECTION)
    batch = client.batch()
    n = 0
    for _, r in rows.iterrows():
        if pd.isna(r["event_date"]):
            continue
        payload = {
            "location": location,
            "value_type": value_type,
            "event_date": _as_iso(r["event_date"]),
            "confirmed": int(r["confirmed"]),
            "suspected": int(r["suspected"]),
            "deaths": int(r["deaths"]),
            "source_type": source_type,
            "source_name": str(source_name).strip(),
            "source_url": str(source_url).strip(),
            "as_of_date": _as_iso(as_of_date),
            "note": str(note).strip(),
            "ingested_at": _fs.SERVER_TIMESTAMP,
        }
        did = doc_id(location, value_type, r["event_date"], source_name, as_of_date)
        batch.set(col.document(did), payload)
        n += 1
        if n % 450 == 0:          # Firestore batch limit is 500
            batch.commit()
            batch = client.batch()
    batch.commit()
    return n


def delete_snapshot(document_id: str) -> None:
    get_client().collection(COLLECTION).document(document_id).delete()


# ---------------------------------------------------------------------------
# Reads
# ---------------------------------------------------------------------------
def load_snapshots(location: str | None = None) -> pd.DataFrame:
    """Return ALL stored documents (optionally for one location) as a tidy
    DataFrame — every version, no reduction. Used by the management UI.

    Query uses an equality filter only, so it relies on Firestore's automatic
    single-field index; no composite index needed.
    """
    from google.cloud.firestore_v1.base_query import FieldFilter
    col = get_client().collection(COLLECTION)
    q = col.where(filter=FieldFilter("location", "==", location)) if location else col
    recs = []
    for snap in q.stream():
        d = snap.to_dict()
        d["document_id"] = snap.id
        recs.append(d)
    cols = ["document_id", "location", "value_type", "event_date",
            "confirmed", "suspected", "deaths", "source_type",
            "source_name", "source_url", "as_of_date", "note", "ingested_at"]
    df = pd.DataFrame(recs)
    if df.empty:
        return pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = None
    df["event_date"] = pd.to_datetime(df["event_date"], errors="coerce")
    df["as_of_date"] = pd.to_datetime(df["as_of_date"], errors="coerce")
    return df[cols].sort_values(["location", "value_type", "event_date",
                                 "as_of_date"], ignore_index=True)


def available_sources(location: str) -> pd.DataFrame:
    """Distinct (source_type, source_name) pairs stored for a location, with a
    count of documents each — used to populate the Step 1 source picker."""
    df = load_snapshots(location)
    if df.empty:
        return pd.DataFrame(columns=["source_type", "source_name", "n"])
    g = (df.groupby(["source_type", "source_name"])
           .size().reset_index(name="n"))
    return g.sort_values(["source_type", "source_name"], ignore_index=True)


def load_for_step1(location: str, value_type: str, as_of=None,
                   source_type=None, source_name=None) -> pd.DataFrame:
    """Return the exact frame Step 1 expects, with as-of reduction applied.

    For each event_date, keeps the reading from the LATEST bulletin published
    on or before `as_of` (newest as_of_date wins; ties broken by ingested_at).
    Optionally restrict to one `source_type` (e.g. "INSP SitRep") and/or one
    specific `source_name` bulletin before the as-of reduction, so a single
    source can be charted on its own.
    Columns returned match app.py:
        cumulative -> date, cumulative_confirmed/suspected/deaths, source
        incidence  -> date, new_confirmed/suspected/deaths, source
    """
    if value_type not in VALUE_TYPES:
        raise ValueError(f"value_type must be one of {VALUE_TYPES}")

    df = load_snapshots(location)
    if df.empty:
        return df
    df = df[df["value_type"] == value_type].copy()
    if df.empty:
        return df

    if source_type:
        df = df[df["source_type"] == source_type]
    if source_name:
        df = df[df["source_name"] == source_name]
    if df.empty:
        return df

    if as_of is not None:
        cutoff = pd.to_datetime(_as_iso(as_of))
        df = df[df["as_of_date"] <= cutoff]
        if df.empty:
            return df

    # as-of reduction: newest as_of_date (then ingested_at) per event_date
    df = df.sort_values(["event_date", "as_of_date", "ingested_at"])
    df = df.drop_duplicates(subset="event_date", keep="last")

    rename = _STEP1_RENAME[value_type]
    out = pd.DataFrame({
        "date": df["event_date"].values,
        rename["confirmed"]: df["confirmed"].astype(float).values,
        rename["suspected"]: df["suspected"].astype(float).values,
        rename["deaths"]: df["deaths"].astype(float).values,
        # Step 1 reads a single free-text `source`; prefer the URL, fall back
        # to the bulletin name.
        "source": [u or n for u, n in zip(df["source_url"].fillna(""),
                                          df["source_name"].fillna(""))],
    })
    return out.sort_values("date").reset_index(drop=True)


if __name__ == "__main__":
    # Smoke test: connect and report current document count.
    c = get_client()
    docs = list(c.collection(COLLECTION).limit(1000).stream())
    print(f"Connected to '{PROJECT_ID}'. snapshots in store: {len(docs)}")
