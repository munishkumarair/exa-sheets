import io
import os
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
from helpers import populate_dataframe, fill_rows

st.set_page_config(page_title="Exa Sheets", layout="centered")

st.title("Exa Sheets")
st.caption("Upload an Excel template (first column companies, next columns data points).")

# Basic logging setup to console
logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO"),
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)

with st.expander("Expected Input Format", expanded=True):
    st.write("The sheet should look like this: first column is company names; subsequent columns are data points.")
    # Resolve asset path relative to project root or this file location
    possible_paths = [
        Path("assets/sample_input.png"),
        Path(__file__).resolve().parent.parent / "assets" / "sample_input.png",
    ]
    img_path = next((p for p in possible_paths if p.exists()), possible_paths[0])
    st.image(str(img_path), caption="Sample input structure", width='stretch')

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"]) 

if uploaded is not None:
    try:
        df_in = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
        st.stop()

    st.subheader("Data preview")
    st.dataframe(df_in.head(), width='stretch')

    if "df_out" not in st.session_state:
        st.session_state["df_out"] = None
    if "sample_filled_indices" not in st.session_state:
        st.session_state["sample_filled_indices"] = set()
    if "last_sample_indices" not in st.session_state:
        st.session_state["last_sample_indices"] = []

    col1, col2 = st.columns(2)

    if col1.button("Fill sample (top 5)"):
        with st.spinner("Filling top 5 rows..."):
            df_out = pd.DataFrame(columns=list(df_in.columns))
            df_out[df_in.columns[0]] = df_in[df_in.columns[0]]
            if st.session_state["df_out"] is not None:
                # Preserve any previously computed values
                df_out = st.session_state["df_out"].copy()
            top_indices = list(df_in.index[:5])
            fill_rows(df_out, df_in, top_indices)
            st.session_state["df_out"] = df_out
            st.session_state["sample_filled_indices"].update(top_indices)
            st.session_state["last_sample_indices"] = top_indices
        st.success("Sample (top 5) filled.")

    if col2.button("Proceed to full fill"):
        with st.spinner("Filling remaining rows..."):
            # Start from existing df_out if present, otherwise initialize fresh
            if st.session_state["df_out"] is None:
                df_out = pd.DataFrame(columns=list(df_in.columns))
                df_out[df_in.columns[0]] = df_in[df_in.columns[0]]
            else:
                df_out = st.session_state["df_out"].copy()

            remaining_indices = [i for i in df_in.index if i not in st.session_state["sample_filled_indices"]]
            fill_rows(df_out, df_in, remaining_indices)
            st.session_state["df_out"] = df_out

        st.success("Completed.")
        st.subheader("Results Preview")
        preview_df = st.session_state["df_out"].head() if st.session_state["df_out"] is not None else df_out.head()
        st.dataframe(preview_df, width='stretch')

        # Offer download as Excel
        buffer = io.BytesIO()
        final_df = st.session_state["df_out"] if st.session_state["df_out"] is not None else df_out
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            final_df.to_excel(writer, index=False)
        buffer.seek(0)

        timestamp = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
        filename = f"exa-sheets-output-{timestamp}.xlsx"
        st.download_button(
            label="Download results (.xlsx)",
            data=buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Always show the current sample preview (if available) below the buttons
    if st.session_state["df_out"] is not None and st.session_state["last_sample_indices"]:
        st.subheader("Sample Preview (Top 5)")
        try:
            sample_preview = st.session_state["df_out"].loc[st.session_state["last_sample_indices"]]
        except Exception:
            sample_preview = st.session_state["df_out"].head()
        st.dataframe(sample_preview, width='stretch')

st.divider()
st.caption(
    "Requires EXA_API_KEY set in environment. Your key is used only in your browser session."
)


