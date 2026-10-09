import tempfile
from pathlib import Path

import streamlit as st

from roi_excel_calculator_v2_dynamic_premises_fixed_v8 import (
    create_excel_download,
    process_file,
    read_premises_values,
)
from frame_trace import render_traced_tables
from frame_chat import render_chat_panel


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Frame ROI Calculator",
    page_icon="📊",
    layout="wide",
)


# =========================================================
# HEADER
# =========================================================

st.title("Frame Services Calculator")

st.write(
    "Upload the Services Project Excel file. "
    "The application calculates Current Services and Future Services "
    "row by row using the X/Y/Z assumptions from the Premises tab."
)

# st.info(
#     "No ROI percentage is calculated. "
#     "The output contains FTE, cost, value, yearly margin contribution, "
#     "and monthly margin contribution."
# )


# =========================================================
# INPUT
# =========================================================

st.subheader("Cost and Schedule to reach Future services")
# st.caption(
#     "Enter Total % FTE per service for Talent and AI. "
#     "Total Investment (row 3) = row 1 + row 2 and is calculated "
#     "automatically. Cost and value use the X assumptions "
#     "from the Premises tab."
# )

SCHEDULE_LABELS = [
    "Team and AI services % FTE (Tasks, Upskilling)",
    "Cost of NRE Services of Project",
    "Recurring Cost of Project at % FTE",
]

talent_col, ai_col = st.columns(2)

with talent_col:
    st.markdown("**Talent Total % FTE per service**")
    talent_fte = [
        st.number_input(
            label, min_value=0.0, value=0.0, step=0.1, format="%.2f",
            key=f"talent_fte_{i}",
        )
        for i, label in enumerate(SCHEDULE_LABELS)
    ]

with ai_col:
    st.markdown("**AI Total % FTE per service**")
    ai_fte = [
        st.number_input(
            label, min_value=0.0, value=0.0, step=0.1, format="%.2f",
            key=f"ai_fte_{i}",
        )
        for i, label in enumerate(SCHEDULE_LABELS)
    ]

schedule_fte = {"talent": talent_fte, "ai": ai_fte}


uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"],
    help="Upload the Excel workbook containing Services Project and Premises tabs.",
)


# =========================================================
# PROCESS
# =========================================================

if uploaded_file is not None:

    st.success(f"Uploaded: {uploaded_file.name}")

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=Path(uploaded_file.name).suffix,
    ) as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        input_path = temp_file.name

    try:
        with st.spinner("Calculating Current and Future Services..."):
            (
                current_df,
                future_df,
                combined_df,
                summary_df,
                future_schedule_df,
                future_current_margin_df,
            ) = process_file(input_path, schedule_fte_inputs=schedule_fte)

        # -----------------------------------------------------
        # RESULT TABLES (hover a cell to see its formula + inputs used)
        # -----------------------------------------------------

        render_traced_tables(
            current_df=current_df,
            future_df=future_df,
            schedule_df=future_schedule_df,
            margin_df=future_current_margin_df,
            premises=read_premises_values(input_path),
        )

        # -----------------------------------------------------
        # DOWNLOAD
        # -----------------------------------------------------

        excel_data = create_excel_download(
            current_df=current_df,
            future_df=future_df,
            combined_df=combined_df,
            summary_df=summary_df,
            future_schedule_df=future_schedule_df,
            future_current_margin_df=future_current_margin_df,
        )

        st.download_button(
            label="Download Calculated Excel",
            data=excel_data,
            file_name="frame_roi_calculated_output.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
        )

        st.caption(
            "The downloaded workbook contains Summary, Current Services, "
            "Future Services, All Calculations, and Premises Used sheets."
        )

        # -----------------------------------------------------
        # CHAT WITH RESULTS
        # -----------------------------------------------------

        render_chat_panel(
            input_path=input_path,
            file_key=(
                f"{uploaded_file.name}:{uploaded_file.size}:"
                f"{schedule_fte}"
            ),
            current_df=current_df,
            future_df=future_df,
            schedule_df=future_schedule_df,
            margin_df=future_current_margin_df,
            schedule_fte=schedule_fte,
        )

    except Exception as exc:
        st.error(f"Error processing Excel file: {exc}")

    finally:
        try:
            Path(input_path).unlink(missing_ok=True)
        except Exception:
            pass