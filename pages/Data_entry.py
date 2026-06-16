"""
Data_entry.py — write/admin page for the EVD snapshot store.

A page within the forecaster app (pages/Data_entry.py); reach it from the
"Open data entry" button on the main page — same app, same URL. Run the whole
app with `streamlit run app.py` (the page's links need the multipage context).

It only writes to / manages Firestore via firestore_store.py. The forecasting
app (app.py) imports the SAME module but reads only. Entry flow, as agreed:
    Location -> Source -> Value type -> Confirmed / Suspected / Deaths
with both single-row and bulk entry, and same-bulletin overwrite (Decision A).
"""

from datetime import date

import pandas as pd
import streamlit as st

import firestore_store as fss

st.set_page_config(page_title="EVD Snapshot Store — Data entry", layout="wide",
                   initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
      [data-testid="stSidebarNav"] {display: none;}
      .block-container {max-width: 1200px; padding-top: 1.4rem;}
      h1.de-title {font-size: 1.5rem; font-weight: 600; color: #1f4e79;
                   margin-bottom: .1rem;}
      .de-sub {color: #5b6573; font-size: .92rem; margin-bottom: 1rem;}
      .de-step {font-size: .8rem; font-weight: 600; color: #5b6573;
                text-transform: uppercase; letter-spacing: .05em;
                margin: .6rem 0 .3rem 0;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<h1 class="de-title">EVD Snapshot Store — Data entry</h1>',
            unsafe_allow_html=True)
st.markdown(
    f'<div class="de-sub">Writes to Firestore project '
    f'<code>{fss.PROJECT_ID}</code>. Same source + value type + event date '
    f'from the <b>same bulletin</b> overwrites the previous reading; a '
    f'different bulletin revising an old date is kept as a new version.</div>',
    unsafe_allow_html=True,
)

# Back to the forecaster (same app, same URL)
st.page_link("app.py", label="Back to forecaster", icon="🏠")


def _require_access():
    """Gate this page behind a user id + password set in Streamlit secrets
    (`data_entry_user`, `data_entry_password`). If no password secret exists,
    access is allowed but a warning is shown (so a fresh deploy isn't
    accidentally locked out)."""
    try:
        pw = st.secrets["data_entry_password"]
    except Exception:
        pw = None
    try:
        user = st.secrets["data_entry_user"]
    except Exception:
        user = None
    if not pw:
        st.warning(
            "⚠️ Data entry is unprotected — anyone with this app's URL can add "
            "or delete records. Set `data_entry_user` / `data_entry_password` "
            "in Streamlit secrets to lock it.")
        return
    if st.session_state.get("_de_authed"):
        return
    st.markdown("#### 🔒 Data entry is restricted")
    u_in = st.text_input("User ID", key="_de_user") if user else None
    pw_in = st.text_input("Password", type="password", key="_de_pw")
    if st.button("Unlock", key="_de_unlock"):
        if pw_in == pw and (user is None or u_in == user):
            st.session_state["_de_authed"] = True
            st.rerun()
        else:
            st.error("Incorrect user ID or password.")
    st.stop()


_require_access()

# Connection check ----------------------------------------------------------
try:
    fss.get_client()
    st.caption("🟢 Connected to Firestore.")
except Exception as e:
    st.error(f"Could not connect to Firestore: {e}")
    st.stop()


def _source_block(prefix: str):
    """Render Source-type / name / URL / as-of inputs. Returns a dict."""
    st.markdown('<div class="de-step">2 · Source</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1, 2])
    with c1:
        s_type = st.selectbox("Source type", fss.SOURCE_TYPES,
                              key=f"{prefix}_stype")
    with c2:
        s_name = st.text_input(
            "Source name (the specific bulletin)",
            placeholder="e.g. WHO DON605  /  Africa CDC SitRep #12",
            key=f"{prefix}_sname",
            help="The exact bulletin. This is part of the overwrite key, so "
                 "DON605 and DON607 are kept as separate versions.")
    c3, c4 = st.columns([2, 1])
    with c3:
        s_url = st.text_input("Source URL / DOI", placeholder="https://…",
                              key=f"{prefix}_surl")
    with c4:
        as_of = st.date_input("As-of (publication date)", value=date.today(),
                              format="DD/MM/YYYY", key=f"{prefix}_asof",
                              help="When this bulletin was published. The "
                                   "versioning key for backtesting.")
    return {"source_type": s_type, "source_name": s_name,
            "source_url": s_url, "as_of_date": as_of}


_BULK_COLS = ["event_date", "confirmed", "suspected", "deaths"]
# Accepted aliases -> canonical column name (case-insensitive match).
_COL_ALIASES = {
    "event_date": "event_date", "date": "event_date", "day": "event_date",
    "confirmed": "confirmed", "confirmed_cases": "confirmed",
    "new_confirmed": "confirmed", "cumulative_confirmed": "confirmed",
    "suspected": "suspected", "suspected_cases": "suspected",
    "new_suspected": "suspected", "cumulative_suspected": "suspected",
    "deaths": "deaths", "death": "deaths",
    "new_deaths": "deaths", "cumulative_deaths": "deaths",
}


def _read_bulk_csv(uploaded) -> pd.DataFrame:
    """Read an uploaded CSV into the frame add_snapshots_bulk expects:
    columns event_date, confirmed, suspected, deaths.

    Accepts common aliases (e.g. `date`, `new_confirmed`). Parses dates
    flexibly (ISO `YYYY-MM-DD` or `DD Mmm YYYY` recommended). Missing
    count columns default to 0. Raises ValueError on bad input."""
    raw = pd.read_csv(uploaded, dtype=str)
    # map any recognised column (case/space-insensitive) to canonical name
    rename = {}
    for c in raw.columns:
        key = c.strip().lower().replace(" ", "_")
        if key in _COL_ALIASES:
            rename[c] = _COL_ALIASES[key]
    raw = raw.rename(columns=rename)

    if "event_date" not in raw.columns:
        raise ValueError(
            "CSV must have a date column (named `event_date` or `date`).")

    out = pd.DataFrame()
    # format="mixed" parses each value independently, so a file may freely mix
    # ISO (2026-05-29) and DD Mmm YYYY (29 May 2026) without the format being
    # locked in from the first row.
    try:
        out["event_date"] = pd.to_datetime(raw["event_date"], errors="coerce",
                                            format="mixed", dayfirst=False)
    except (ValueError, TypeError):
        out["event_date"] = pd.to_datetime(raw["event_date"], errors="coerce",
                                            dayfirst=False)
    bad = out["event_date"].isna()
    if bad.any():
        rows = (raw.index[bad] + 2).tolist()  # +2: header + 1-based
        raise ValueError(
            f"Unparseable date(s) on CSV row(s) {rows}. Use YYYY-MM-DD "
            "(e.g. 2026-05-29) or DD Mmm YYYY (e.g. 29 May 2026).")

    for col in ("confirmed", "suspected", "deaths"):
        if col in raw.columns:
            nums = pd.to_numeric(raw[col], errors="coerce")
            if nums.isna().any():
                rows = (raw.index[nums.isna()] + 2).tolist()
                raise ValueError(f"Non-numeric `{col}` on CSV row(s) {rows}.")
            out[col] = nums.astype(int)
        else:
            out[col] = 0  # column absent -> treat as zero
    return out[_BULK_COLS]


tab_one, tab_bulk, tab_manage = st.tabs(
    ["➕ Add one", "📋 Bulk add", "🗂️ Browse / manage"])

# ===========================================================================
# TAB 1 — single reading
# ===========================================================================
with tab_one:
    st.markdown('<div class="de-step">1 · Location</div>',
                unsafe_allow_html=True)
    location = st.selectbox("Outbreak location", fss.LOCATIONS, key="one_loc")

    src = _source_block("one")

    st.markdown('<div class="de-step">3 · Value type</div>',
                unsafe_allow_html=True)
    value_type = st.radio("Value type", fss.VALUE_TYPES, horizontal=True,
                          key="one_vtype",
                          help="cumulative = running totals (WHO DON style); "
                               "incidence = new cases that day (Africa CDC "
                               "daily SitRep style).")

    st.markdown('<div class="de-step">4 · Counts</div>', unsafe_allow_html=True)
    is_cum = value_type == "cumulative"
    lbl = "Cumulative" if is_cum else "New"
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        ev_date = st.date_input("Event date", value=date.today(),
                                format="DD/MM/YYYY", key="one_evdate")
    with c2:
        confirmed = st.number_input(f"{lbl} confirmed", min_value=0, step=1,
                                    key="one_conf")
    with c3:
        suspected = st.number_input(f"{lbl} suspected", min_value=0, step=1,
                                    key="one_susp")
    with c4:
        deaths = st.number_input(f"{lbl} deaths", min_value=0, step=1,
                                 key="one_death")
    note = st.text_input("Note (optional)", key="one_note",
                         placeholder="e.g. DON607 revised the 29 May count")

    if st.button("Save reading", type="primary", key="one_save"):
        if not src["source_name"].strip():
            st.error("Source name is required (it's part of the overwrite key).")
        else:
            try:
                did = fss.add_snapshot(
                    location=location, value_type=value_type,
                    event_date=ev_date, confirmed=confirmed,
                    suspected=suspected, deaths=deaths, note=note, **src)
                st.success(f"Saved ✓  ({location} · {value_type} · "
                           f"{ev_date:%d %b %Y} · {src['source_name']}). "
                           f"Any previous reading for this exact bulletin/date "
                           f"was overwritten.")
                st.caption(f"doc id: {did}")
            except Exception as e:
                st.error(f"Write failed: {e}")

# ===========================================================================
# TAB 2 — bulk (one bulletin, many dates)
# ===========================================================================
with tab_bulk:
    st.caption("Pick the bulletin's metadata once, then add many per-date "
               "rows — either by typing them or by uploading a CSV. All rows "
               "share the source and value type. Works for any source type "
               "(WHO DON, Africa CDC, INSP SitRep, Other).")
    st.markdown('<div class="de-step">1 · Location</div>',
                unsafe_allow_html=True)
    b_location = st.selectbox("Outbreak location", fss.LOCATIONS, key="b_loc")

    b_src = _source_block("b")

    st.markdown('<div class="de-step">3 · Value type</div>',
                unsafe_allow_html=True)
    b_vtype = st.radio("Value type", fss.VALUE_TYPES, horizontal=True,
                       key="b_vtype")
    b_note = st.text_input("Note applied to all rows (optional)", key="b_note")

    is_cum = b_vtype == "cumulative"
    lbl = "Cumulative" if is_cum else "New"

    st.markdown('<div class="de-step">4 · Rows (date + counts)</div>',
                unsafe_allow_html=True)
    b_mode = st.radio("How do you want to add rows?",
                      ["Type rows", "Upload CSV file"], horizontal=True,
                      key="b_mode")

    # rows to save get collected here, whichever mode is active
    to_save = None

    if b_mode == "Type rows":
        seed = pd.DataFrame({"event_date": [date.today()], "confirmed": [0],
                             "suspected": [0], "deaths": [0]})
        edited = st.data_editor(
            seed, num_rows="dynamic", use_container_width=True, height=320,
            key="b_editor",
            column_config={
                "event_date": st.column_config.DateColumn(
                    "Event date", format="DD MMM YYYY", required=True),
                "confirmed": st.column_config.NumberColumn(
                    f"{lbl} confirmed", min_value=0, step=1, required=True),
                "suspected": st.column_config.NumberColumn(
                    f"{lbl} suspected", min_value=0, step=1, required=True),
                "deaths": st.column_config.NumberColumn(
                    f"{lbl} deaths", min_value=0, step=1, required=True),
            })
        to_save = edited

    else:  # "Upload CSV file"
        st.caption(
            "CSV needs a **date** column (`event_date` or `date`) plus any of "
            "`confirmed`, `suspected`, `deaths` (missing ones count as 0). "
            "Aliases like `new_confirmed` / `cumulative_deaths` are accepted. "
            "**Dates:** YYYY-MM-DD (e.g. 2026-05-29) or DD Mmm YYYY "
            "(e.g. 29 May 2026). The bulletin metadata above is applied to "
            "every row — the CSV holds only dates and counts.")
        # downloadable template so users know the exact shape
        _tmpl = pd.DataFrame({
            "event_date": ["2026-05-29", "2026-05-30"],
            "confirmed": [134, 5], "suspected": [906, 0], "deaths": [241, 2]})
        st.download_button(
            "⬇️ Download CSV template", _tmpl.to_csv(index=False).encode("utf-8"),
            file_name="bulk_upload_template.csv", mime="text/csv",
            key="b_tmpl")

        b_csv = st.file_uploader("Upload CSV", type=["csv"], key="b_csv")
        if b_csv is not None:
            try:
                to_save = _read_bulk_csv(b_csv)
                st.success(f"Loaded {len(to_save)} row(s) from "
                           f"{b_csv.name}. Review below, then save.")
                preview = to_save.copy()
                preview["event_date"] = preview["event_date"].dt.strftime(
                    "%d %b %Y")
                st.dataframe(preview, use_container_width=True, height=260,
                             hide_index=True)
            except Exception as e:
                st.error(f"Could not read CSV: {e}")
                to_save = None

    if st.button("Save all rows", type="primary", key="b_save"):
        if not b_src["source_name"].strip():
            st.error("Source name is required.")
        elif to_save is None or len(to_save) == 0:
            st.error("No rows to save — add rows or upload a CSV first.")
        else:
            try:
                n = fss.add_snapshots_bulk(
                    to_save, location=b_location, value_type=b_vtype,
                    note=b_note, **b_src)
                st.success(f"Saved {n} row(s) ✓ for {b_src['source_name']} "
                           f"({b_location} · {b_vtype}). Existing readings for "
                           f"the same bulletin/dates were overwritten.")
            except Exception as e:
                st.error(f"Bulk write failed: {e}")

# ===========================================================================
# TAB 3 — browse / manage
# ===========================================================================
with tab_manage:
    top_l, top_r = st.columns([3, 1])
    with top_l:
        m_loc = st.selectbox("Location", ["(all)"] + fss.LOCATIONS,
                             key="m_loc")
    with top_r:
        st.markdown("<div style='height:1.8rem'></div>", unsafe_allow_html=True)
        if st.button("🔄 Refresh", key="m_refresh", use_container_width=True):
            st.rerun()

    df = fss.load_snapshots(None if m_loc == "(all)" else m_loc)

    if df.empty:
        st.info("No snapshots stored yet.")
    else:
        # ---- Advanced filters -------------------------------------------
        with st.expander("🔎 Advanced filters", expanded=False):
            f1, f2, f3 = st.columns(3)
            with f1:
                f_stype = st.multiselect(
                    "Source type", sorted(df["source_type"].dropna().unique()),
                    key="m_f_stype")
                f_vtype = st.multiselect(
                    "Value type", sorted(df["value_type"].dropna().unique()),
                    key="m_f_vtype")
            with f2:
                f_sname = st.multiselect(
                    "Bulletin (source name)",
                    sorted(df["source_name"].dropna().unique()),
                    key="m_f_sname")
                f_text = st.text_input(
                    "Search note / source / URL", key="m_f_text",
                    placeholder="free text…")
            with f3:
                _min_d = df["event_date"].min().date()
                _max_d = df["event_date"].max().date()
                f_dates = st.date_input(
                    "Event-date range", value=(_min_d, _max_d),
                    min_value=_min_d, max_value=_max_d,
                    format="DD/MM/YYYY", key="m_f_dates")

        fdf = df.copy()
        if st.session_state.get("m_f_stype"):
            fdf = fdf[fdf["source_type"].isin(st.session_state["m_f_stype"])]
        if st.session_state.get("m_f_vtype"):
            fdf = fdf[fdf["value_type"].isin(st.session_state["m_f_vtype"])]
        if st.session_state.get("m_f_sname"):
            fdf = fdf[fdf["source_name"].isin(st.session_state["m_f_sname"])]
        if isinstance(f_dates, (tuple, list)) and len(f_dates) == 2:
            lo, hi = pd.to_datetime(f_dates[0]), pd.to_datetime(f_dates[1])
            fdf = fdf[(fdf["event_date"] >= lo) & (fdf["event_date"] <= hi)]
        txt = st.session_state.get("m_f_text", "").strip().lower()
        if txt:
            hay = (fdf["note"].fillna("") + " " + fdf["source_name"].fillna("")
                   + " " + fdf["source_url"].fillna("")).str.lower()
            fdf = fdf[hay.str.contains(txt, regex=False)]

        # ---- Summary chips ----------------------------------------------
        st.caption(
            f"Showing **{len(fdf)}** of {len(df)} document(s) · "
            f"{fdf['source_type'].nunique()} source type(s) · "
            f"{fdf['source_name'].nunique()} bulletin(s). Every version is "
            "shown (no as-of reduction). Tick rows to delete.")

        # ---- Table with row-select for deletion -------------------------
        show = fdf.copy()
        show.insert(0, "🗑", False)
        show["event_date"] = show["event_date"].dt.strftime("%d %b %Y")
        show["as_of_date"] = show["as_of_date"].dt.strftime("%d %b %Y")
        col_order = ["🗑", "location", "value_type", "event_date",
                     "confirmed", "suspected", "deaths", "source_type",
                     "source_name", "as_of_date", "source_url", "note",
                     "document_id"]
        show = show[[c for c in col_order if c in show.columns]]
        edited_tbl = st.data_editor(
            show, use_container_width=True, hide_index=True, height=420,
            key="m_table",
            disabled=[c for c in show.columns if c != "🗑"],
            column_config={
                "🗑": st.column_config.CheckboxColumn(
                    "🗑", help="Tick to mark for deletion", width="small"),
                "document_id": None,  # hidden — kept for the delete call
                "value_type": st.column_config.TextColumn("type", width="small"),
                "source_type": st.column_config.TextColumn("source"),
                "source_name": st.column_config.TextColumn("bulletin"),
                "source_url": st.column_config.LinkColumn("url"),
                "confirmed": st.column_config.NumberColumn("conf.",
                                                           width="small"),
                "suspected": st.column_config.NumberColumn("susp.",
                                                           width="small"),
                "deaths": st.column_config.NumberColumn("deaths",
                                                        width="small"),
            })

        picked_ids = edited_tbl.loc[edited_tbl["🗑"], "document_id"].tolist()

        # ---- Delete actions ---------------------------------------------
        st.markdown('<div class="de-step">Delete</div>', unsafe_allow_html=True)
        d1, d2 = st.columns(2)

        # (a) delete the ticked rows
        with d1:
            if st.button(f"🗑 Delete selected ({len(picked_ids)})",
                         disabled=not picked_ids, use_container_width=True,
                         key="m_del_sel"):
                try:
                    n = fss.delete_snapshots(picked_ids)
                    st.success(f"Deleted {n} document(s) ✓")
                    st.rerun()
                except Exception as e:
                    st.error(f"Delete failed: {e}")

        # (b) clear all (respects the location selector), double-confirmed
        with d2:
            scope = "ALL locations" if m_loc == "(all)" else m_loc
            with st.popover(f"🧨 Clear all ({scope})",
                            use_container_width=True):
                st.warning(
                    f"This permanently deletes **every** stored document for "
                    f"**{scope}** ({len(df)} row(s)). This cannot be undone.")
                confirm = st.text_input(
                    "Type DELETE to confirm", key="m_clear_confirm",
                    placeholder="DELETE")
                if st.button("Permanently clear", type="primary",
                             disabled=(confirm != "DELETE"), key="m_clear_go"):
                    try:
                        n = fss.delete_all(
                            None if m_loc == "(all)" else m_loc)
                        st.success(f"Cleared {n} document(s) ✓")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Clear failed: {e}")
