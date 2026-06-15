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
    st.caption("Pick the bulletin's metadata once, then paste/enter many "
               "per-date rows below. All rows share the source and value type.")
    st.markdown('<div class="de-step">1 · Location</div>',
                unsafe_allow_html=True)
    b_location = st.selectbox("Outbreak location", fss.LOCATIONS, key="b_loc")

    b_src = _source_block("b")

    st.markdown('<div class="de-step">3 · Value type</div>',
                unsafe_allow_html=True)
    b_vtype = st.radio("Value type", fss.VALUE_TYPES, horizontal=True,
                       key="b_vtype")
    b_note = st.text_input("Note applied to all rows (optional)", key="b_note")

    st.markdown('<div class="de-step">4 · Rows (date + counts)</div>',
                unsafe_allow_html=True)
    is_cum = b_vtype == "cumulative"
    lbl = "Cumulative" if is_cum else "New"
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

    if st.button("Save all rows", type="primary", key="b_save"):
        if not b_src["source_name"].strip():
            st.error("Source name is required.")
        else:
            try:
                n = fss.add_snapshots_bulk(
                    edited, location=b_location, value_type=b_vtype,
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
    m_loc = st.selectbox("Filter by location", ["(all)"] + fss.LOCATIONS,
                         key="m_loc")
    if st.button("Refresh", key="m_refresh"):
        st.rerun()
    df = fss.load_snapshots(None if m_loc == "(all)" else m_loc)
    if df.empty:
        st.info("No snapshots stored yet.")
    else:
        st.caption(f"{len(df)} document(s). Every version is shown (no as-of "
                   "reduction here).")
        show = df.drop(columns=["document_id"]).copy()
        show["event_date"] = show["event_date"].dt.strftime("%d %b %Y")
        show["as_of_date"] = show["as_of_date"].dt.strftime("%d %b %Y")
        st.dataframe(show, use_container_width=True, hide_index=True)

        st.markdown('<div class="de-step">Delete a document</div>',
                    unsafe_allow_html=True)
        labels = {
            f"{r.location} · {r.value_type} · "
            f"{pd.to_datetime(r.event_date):%d %b %Y} · {r.source_name} "
            f"(as-of {pd.to_datetime(r.as_of_date):%d %b %Y})": r.document_id
            for r in df.itertuples(index=False)}
        pick = st.selectbox("Pick a document", list(labels.keys()),
                            key="m_pick")
        if st.button("Delete selected", key="m_del"):
            try:
                fss.delete_snapshot(labels[pick])
                st.success("Deleted ✓")
                st.rerun()
            except Exception as e:
                st.error(f"Delete failed: {e}")
