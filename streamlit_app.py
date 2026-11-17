import streamlit as st
import tempfile
import os
from pathlib import Path

import io
import pandas as pd

from matcher import run_matching, run_matching_multi
import storage

st.set_page_config(page_title="Center Matcher Tool", page_icon="🔎", layout="wide")

# ── Custom styling (colors work in both light and dark theme) ──
st.markdown("""
<style>
:root {
    --c-primary: #5B5BD6;
    --c-primary-dark: #4338CA;
    --c-teal: #0D9488;
}
/* Hero banner */
.hero {
    background: linear-gradient(120deg, #4338CA 0%, #6D5BE0 45%, #0D9488 100%);
    border-radius: 16px;
    padding: 26px 28px;
    margin-bottom: 18px;
    color: #FFFFFF;
    box-shadow: 0 6px 18px rgba(67, 56, 202, 0.25);
}
.hero h1 { color: #FFFFFF; margin: 0; padding: 0; font-size: 2rem; }
.hero p  { color: rgba(255,255,255,0.88); margin: 6px 0 0 0; font-size: 0.95rem; }

/* Tabs as pills (works for both old button-based and new div-based Streamlit tabs) */
div[role="tablist"], div[data-baseweb="tab-list"] { gap: 8px; border-bottom: none !important; box-shadow: none !important; }
[role="tab"] {
    border-radius: 999px !important;
    padding: 8px 18px !important;
    background: rgba(91, 91, 214, 0.12) !important;
    transition: all .15s ease;
    cursor: pointer;
}
[role="tab"]:hover { background: rgba(91, 91, 214, 0.22) !important; }
[role="tab"][aria-selected="true"] {
    background: linear-gradient(120deg, #4338CA, #6D5BE0) !important;
}
[role="tab"][aria-selected="true"] *,
[role="tab"][aria-selected="true"] p { color: #FFFFFF !important; font-weight: 600; }
.react-aria-SelectionIndicator,
div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] { display: none !important; }

/* Buttons */
.stButton > button[kind="primary"],
button[data-testid="stBaseButton-primary"] {
    background: linear-gradient(120deg, #4338CA, #0D9488);
    border: none; color: #FFFFFF; border-radius: 10px; font-weight: 600;
}
.stButton > button[kind="primary"]:hover,
button[data-testid="stBaseButton-primary"]:hover { filter: brightness(1.1); }
.stDownloadButton > button, button[data-testid="stBaseButton-secondary"] {
    border-radius: 10px; border: 1.5px solid #5B5BD6;
}

/* Section headings */
h2, h3 { border-left: 5px solid #0D9488; padding-left: 10px; }

/* Stat boxes */
.stat-box {
    border-radius: 14px;
    padding: 16px 10px;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}
.stat-green { background: linear-gradient(135deg, #D1FAE5, #A7F3D0); color: #065F46; border-bottom: 4px solid #10B981; }
.stat-red   { background: linear-gradient(135deg, #FEE2E2, #FECACA); color: #991B1B; border-bottom: 4px solid #EF4444; }
.stat-amber { background: linear-gradient(135deg, #FEF3C7, #FDE68A); color: #92400E; border-bottom: 4px solid #F59E0B; }
.stat-row-gap { margin-bottom: 18px; }
.stat-num   { font-size: 28px; font-weight: 800; }
.stat-lbl   { font-size: 12px; margin-top: 2px; font-weight: 600; letter-spacing: .4px; text-transform: uppercase; }

/* Metric cards */
div[data-testid="stMetric"] {
    background: rgba(91, 91, 214, 0.08);
    border-left: 5px solid #5B5BD6;
    border-radius: 12px;
    padding: 10px 16px;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="hero"><h1>🔎 Center Matcher Tool</h1>'
    '<p>Compare List A against List B using fuzzy name matching. '
    'You can upload one or multiple List B files.</p></div>',
    unsafe_allow_html=True,
)

tab_single, tab_multi, tab_history = st.tabs(
    ["📄 Single B file", "🗂️ Multiple B files", "📚 Saved history"]
)


# ── Permanent storage config (Streamlit secrets [github]) ──
def get_gh_conf():
    try:
        g = st.secrets["github"]
        return {
            "token": g["token"],
            "repo": g["repo"],
            "branch": g.get("branch", "main"),
            "dir": g.get("dir", "runs"),
        }
    except Exception:
        return None


GH_CONF = get_gh_conf()


@st.cache_data(ttl=300, show_spinner=False)
def cached_list_runs(_conf, _bust):
    return storage.list_runs(_conf)


@st.cache_data(show_spinner=False)
def cached_load_run(name, _conf):
    return storage.load_run(name, _conf)


def save_uploaded_file(uploaded_file, dest_dir):
    """Save a Streamlit uploaded file to disk
    (matcher.py functions expect file paths, not in-memory objects)."""
    suffix = Path(uploaded_file.name).suffix
    path = os.path.join(dest_dir, f"{uploaded_file.name}")
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


def show_stats_row(matched, unmatched_a, unmatched_b):
    c1, c2, c3 = st.columns(3)
    c1.markdown(
        f'<div class="stat-box stat-green"><div class="stat-num">{matched}</div>'
        f'<div class="stat-lbl">Matched</div></div>', unsafe_allow_html=True
    )
    c2.markdown(
        f'<div class="stat-box stat-red"><div class="stat-num">{unmatched_a}</div>'
        f'<div class="stat-lbl">Unmatched A</div></div>', unsafe_allow_html=True
    )
    c3.markdown(
        f'<div class="stat-box stat-amber"><div class="stat-num">{unmatched_b}</div>'
        f'<div class="stat-lbl">Unmatched B</div></div>', unsafe_allow_html=True
    )


# ═══════════════════════════════════════════════════════════════
#  TAB 1 — SINGLE B FILE
# ═══════════════════════════════════════════════════════════════
with tab_single:
    st.subheader("1️⃣ List A — master file")
    file_a_single = st.file_uploader(
        "List A (.xlsx / .xls)", type=["xlsx", "xls"], key="single_a"
    )

    st.subheader("2️⃣ List B — comparison file")
    file_b_single = st.file_uploader(
        "List B (.xlsx / .xls)", type=["xlsx", "xls"], key="single_b"
    )

    single_label = st.text_input(
        "Run name (optional — the run is saved in history under this name)", key="single_run_label"
    )
    run_single = st.button("▶️ Run matching", key="run_single", type="primary")

    if run_single:
        if not file_a_single or not file_b_single:
            st.error("Both List A and List B files are required.")
        else:
            with st.spinner("Matching in progress..."):
                with tempfile.TemporaryDirectory() as tmp_dir:
                    try:
                        path_a = save_uploaded_file(file_a_single, tmp_dir)
                        path_b = save_uploaded_file(file_b_single, tmp_dir)
                        output_path = os.path.join(tmp_dir, "Result.xlsx")

                        stats = run_matching(path_a, path_b, output_path)

                        with open(output_path, "rb") as f:
                            result_bytes = f.read()

                        st.session_state["single_stats"] = stats
                        st.session_state["single_result_bytes"] = result_bytes

                        # Save the result permanently (visible in Saved history)
                        st.session_state.pop("single_saved_name", None)
                        st.session_state.pop("single_save_error", None)
                        try:
                            label = "Single_" + (
                                single_label.strip()
                                or f"{Path(file_a_single.name).stem}_vs_{Path(file_b_single.name).stem}"
                            )
                            st.session_state["single_saved_name"] = storage.save_run(
                                label, result_bytes, GH_CONF
                            )
                            st.session_state["history_bust"] = st.session_state.get("history_bust", 0) + 1
                        except Exception as save_err:
                            st.session_state["single_save_error"] = str(save_err)
                    except Exception as e:
                        st.error(f"Error: {e}")
                        st.session_state.pop("single_stats", None)
                        st.session_state.pop("single_result_bytes", None)

    if "single_stats" in st.session_state:
        stats = st.session_state["single_stats"]
        st.success("✅ Matching complete")
        if st.session_state.get("single_saved_name"):
            st.info(f"💾 Saved to history: `{st.session_state['single_saved_name']}` "
                    "— you can view it anytime in the 📚 Saved history tab.")
        if st.session_state.get("single_save_error"):
            st.warning(f"⚠️ Could not save to history: {st.session_state['single_save_error']} "
                       "(Please download the result.)")
        show_stats_row(stats["matched"], stats["unmatched_a"], stats["unmatched_b"])
        st.metric("Match rate", f"{stats['match_rate']}%")

        if stats.get("unmatched_districts"):
            with st.expander(f"⚠️ {len(stats['unmatched_districts'])} district(s) not found in List B"):
                st.write(", ".join(stats["unmatched_districts"]))

        st.download_button(
            label="⬇️ Download Excel result",
            data=st.session_state["single_result_bytes"],
            file_name="Center_Matcher_Result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ═══════════════════════════════════════════════════════════════
#  TAB 2 — MULTIPLE B FILES
# ═══════════════════════════════════════════════════════════════
with tab_multi:
    st.subheader("1️⃣ List A — master file")
    file_a_multi = st.file_uploader(
        "List A (.xlsx / .xls)", type=["xlsx", "xls"], key="multi_a"
    )

    st.subheader("2️⃣ List B — multiple files")
    st.caption("Hold Ctrl (Windows) or Cmd (Mac) to select multiple files — each file gets its own tab in the output.")
    files_b_multi = st.file_uploader(
        "List B files (.xlsx / .xls)", type=["xlsx", "xls"],
        accept_multiple_files=True, key="multi_b"
    )

    if files_b_multi:
        st.write(f"**{len(files_b_multi)} file(s) selected:** " + ", ".join(f.name for f in files_b_multi))

    run_label = st.text_input(
        "Run name (optional — the run is saved in history under this name)", key="multi_run_label"
    )
    run_multi = st.button("▶️ Run matching — All B files", key="run_multi", type="primary")

    if run_multi:
        if not file_a_multi:
            st.error("List A file is required.")
        elif not files_b_multi:
            st.error("At least one List B file is required.")
        else:
            with st.spinner("Matching in progress... This may take a few moments."):
                with tempfile.TemporaryDirectory() as tmp_dir:
                    try:
                        path_a = save_uploaded_file(file_a_multi, tmp_dir)

                        b_paths_with_names = []
                        for fb in files_b_multi:
                            path_b = save_uploaded_file(fb, tmp_dir)
                            tab_name = Path(fb.name).stem[:31]
                            b_paths_with_names.append((tab_name, path_b))

                        output_path = os.path.join(tmp_dir, "Result_Multi.xlsx")

                        stats = run_matching_multi(path_a, b_paths_with_names, output_path)

                        with open(output_path, "rb") as f:
                            result_bytes = f.read()

                        st.session_state["multi_stats"] = stats
                        st.session_state["multi_result_bytes"] = result_bytes

                        # NEW: result permanently save (history me hamesha rahega)
                        st.session_state.pop("multi_saved_name", None)
                        st.session_state.pop("multi_save_error", None)
                        try:
                            label = run_label.strip() or "_".join(t for t, _ in b_paths_with_names)
                            st.session_state["multi_saved_name"] = storage.save_run(
                                label, result_bytes, GH_CONF
                            )
                            st.session_state["history_bust"] = st.session_state.get("history_bust", 0) + 1
                        except Exception as save_err:
                            st.session_state["multi_save_error"] = str(save_err)
                    except Exception as e:
                        st.error(f"Error: {e}")
                        st.session_state.pop("multi_stats", None)
                        st.session_state.pop("multi_result_bytes", None)

    if "multi_stats" in st.session_state:
        stats = st.session_state["multi_stats"]
        st.success("✅ Matching complete")
        if st.session_state.get("multi_saved_name"):
            st.info(f"💾 Saved to history: `{st.session_state['multi_saved_name']}` "
                    "— you can view it anytime in the 📚 Saved history tab.")
        if st.session_state.get("multi_save_error"):
            st.warning(f"⚠️ Could not save to history: {st.session_state['multi_save_error']} "
                       "(Please download the result.)")
        show_stats_row(stats["total_matched"], stats["total_unmatched_a"], stats["total_unmatched_b"])

        # NEW FEATURE: common centers (2+ projects mein match hue)
        st.write("")
        st.markdown("#### 🤝 Common centers (matched in multiple projects)")
        st.metric("Common centers", stats.get("common_count", 0))
        common_df = stats.get("common_df")
        if common_df is not None and not common_df.empty:
            show_cols = [c for c in common_df.columns
                         if c in ("TEST CENTER NAME", "Common In (Count)", "Project Codes")
                         or c.lower().endswith("district name")]
            st.dataframe(common_df[show_cols], use_container_width=True, hide_index=True)
            st.caption("The full record (center name + match % for each project) is in the "
                       "'COMMON_CENTERS' tab of the Excel file.")
        else:
            st.info("No center was matched in more than one project.")

        cs_df = stats.get("common_summary_df")
        if cs_df is not None and not cs_df.empty:
            st.markdown("#### 📊 Project-wise common summary")
            st.dataframe(cs_df, use_container_width=True, hide_index=True)
            st.caption("The pairwise matrix is at the bottom of the 'COMMON_SUMMARY' tab in the Excel file.")

        st.markdown("#### Per-file breakdown")
        st.dataframe(
            [
                {
                    "Tab": f["tab"],
                    "Matched": f["matched"],
                    "Unmatched A": f["unmatched_a"],
                    "Unmatched B": f["unmatched_b"],
                    "Match Rate %": f["match_rate"],
                }
                for f in stats["per_file"]
            ],
            use_container_width=True,
            hide_index=True,
        )

        if stats.get("unmatched_districts"):
            with st.expander(f"⚠️ {len(stats['unmatched_districts'])} district(s) not found in any B file"):
                st.write(", ".join(stats["unmatched_districts"]))

        st.download_button(
            label=f"⬇️ Download Excel result — {len(stats['per_file'])} tabs",
            data=st.session_state["multi_result_bytes"],
            file_name="Center_Matcher_Result_Multi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ═══════════════════════════════════════════════════════════════
#  TAB 3 — SAVED HISTORY (permanent)
# ═══════════════════════════════════════════════════════════════
def split_common_summary(raw):
    """Split the COMMON_SUMMARY sheet into its two tables (project-wise + matrix)."""
    blank = raw.index[raw.isna().all(axis=1)].tolist()
    if not blank:
        return raw, None
    first = raw.iloc[:blank[0]]
    first = pd.DataFrame(first.values[1:], columns=first.iloc[0].tolist())
    rest = raw.iloc[blank[0] + 1:].dropna(how="all")
    matrix = None
    if len(rest) > 1:
        body = rest.iloc[1:]
        matrix = pd.DataFrame(body.values[1:], columns=body.iloc[0].tolist())
        matrix = matrix.dropna(axis=1, how="all")
    return first, matrix


with tab_history:
    st.subheader("📚 Saved history")
    if GH_CONF:
        st.caption(f"✅ Permanent storage: GitHub repo `{GH_CONF['repo']}` — "
                   "data stays safe across reboots and redeploys.")
    else:
        st.warning(
            "⚠️ GitHub storage is not configured — results are being saved to a local folder. "
            "On Streamlit Cloud this is erased on reboot. Add a `[github]` section in "
            "App Settings → Secrets to keep history permanently."
        )

    if st.button("🔄 Refresh list", key="hist_refresh"):
        st.session_state["history_bust"] = st.session_state.get("history_bust", 0) + 1

    try:
        runs = cached_list_runs(GH_CONF, st.session_state.get("history_bust", 0))
    except Exception as e:
        runs = []
        st.error(f"Could not load history: {e}")

    if not runs:
        st.info("No saved runs yet. Run matching in the Multiple B files tab — "
                "the result will be saved here automatically.")
    else:
        by_name = {r["name"]: r for r in runs}
        chosen = st.selectbox(
            f"{len(runs)} saved run(s)",
            options=list(by_name.keys()),
            format_func=lambda n: f"{by_name[n]['saved_at']}  —  {by_name[n]['label']}",
            key="hist_choice",
        )
        try:
            data = cached_load_run(chosen, GH_CONF)
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None)
        except Exception as e:
            sheets = None
            st.error(f"Could not load run: {e}")

        if sheets:
            st.download_button(
                "⬇️ Download this run (Excel)", data=data, file_name=chosen,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="hist_download",
            )
            names = list(sheets.keys())
            default_idx = names.index("COMMON_SUMMARY") if "COMMON_SUMMARY" in names else 0
            sheet_pick = st.selectbox("View tab", names, index=default_idx, key="hist_sheet")
            raw = sheets[sheet_pick]
            if sheet_pick == "COMMON_SUMMARY":
                t1, t2 = split_common_summary(raw)
                st.markdown("**Project-wise common summary**")
                st.dataframe(t1, use_container_width=True, hide_index=True)
                if t2 is not None:
                    st.markdown("**Pairwise common centers**")
                    st.dataframe(t2, use_container_width=True, hide_index=True)
            else:
                df_view = pd.DataFrame(raw.values[1:], columns=raw.iloc[0].tolist()) if len(raw) else raw
                st.dataframe(df_view, use_container_width=True, hide_index=True)


st.divider()
st.caption("Center Matcher Tool — internal use only")
