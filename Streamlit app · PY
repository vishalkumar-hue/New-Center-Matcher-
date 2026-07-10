import streamlit as st
import tempfile
import os
from pathlib import Path

from matcher import run_matching, run_matching_multi

st.set_page_config(page_title="Center Matcher Tool", page_icon="🔎", layout="centered")

# ── Thoda sa custom CSS, taaki dark/light dono theme mein accent color dikhe ──
st.markdown("""
<style>
.stat-box {
    border-radius: 12px;
    padding: 14px 10px;
    text-align: center;
}
.stat-green { background: #E1F5EE; color: #085041; }
.stat-red   { background: #FCEBEB; color: #791F1F; }
.stat-amber { background: #FAEEDA; color: #633806; }
.stat-num   { font-size: 26px; font-weight: 700; }
.stat-lbl   { font-size: 12px; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

st.title("🔎 Center Matcher Tool")
st.caption(
    "List A ko List B ke saath fuzzy naam-matching se compare karo. "
    "Ek ya multiple List B files upload kar sakte ho."
)

tab_single, tab_multi = st.tabs(["📄 Single B file", "🗂️ Multiple B files"])


def save_uploaded_file(uploaded_file, dest_dir):
    """Streamlit ke uploaded file object ko disk pe save karta hai
    (matcher.py ke functions path expect karte hain, in-memory object nahi)."""
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

    run_single = st.button("▶️ Run matching", key="run_single", type="primary")

    if run_single:
        if not file_a_single or not file_b_single:
            st.error("List A aur List B dono files required hain.")
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
                    except Exception as e:
                        st.error(f"Error: {e}")
                        st.session_state.pop("single_stats", None)
                        st.session_state.pop("single_result_bytes", None)

    if "single_stats" in st.session_state:
        stats = st.session_state["single_stats"]
        st.success("✅ Matching complete")
        show_stats_row(stats["matched"], stats["unmatched_a"], stats["unmatched_b"])
        st.metric("Match rate", f"{stats['match_rate']}%")

        if stats.get("unmatched_districts"):
            with st.expander(f"⚠️ {len(stats['unmatched_districts'])} district(s) List B mein nahi mile"):
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
    st.caption("Ctrl (Windows) ya Cmd (Mac) daba ke multiple files select karo — har file output mein apni tab banayegi.")
    files_b_multi = st.file_uploader(
        "List B files (.xlsx / .xls)", type=["xlsx", "xls"],
        accept_multiple_files=True, key="multi_b"
    )

    if files_b_multi:
        st.write(f"**{len(files_b_multi)} file(s) selected:** " + ", ".join(f.name for f in files_b_multi))

    run_multi = st.button("▶️ Run matching — All B files", key="run_multi", type="primary")

    if run_multi:
        if not file_a_multi:
            st.error("List A file required hai.")
        elif not files_b_multi:
            st.error("Kam se kam ek List B file required hai.")
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
                    except Exception as e:
                        st.error(f"Error: {e}")
                        st.session_state.pop("multi_stats", None)
                        st.session_state.pop("multi_result_bytes", None)

    if "multi_stats" in st.session_state:
        stats = st.session_state["multi_stats"]
        st.success("✅ Matching complete")
        show_stats_row(stats["total_matched"], stats["total_unmatched_a"], stats["total_unmatched_b"])

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
            with st.expander(f"⚠️ {len(stats['unmatched_districts'])} district(s) kisi B file mein nahi mile"):
                st.write(", ".join(stats["unmatched_districts"]))

        st.download_button(
            label=f"⬇️ Download Excel result — {len(stats['per_file'])} tabs",
            data=st.session_state["multi_result_bytes"],
            file_name="Center_Matcher_Result_Multi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.divider()
st.caption("Center Matcher Tool — internal use only")
