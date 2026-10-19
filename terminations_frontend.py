import pandas as pd
import streamlit as st
from termination_backend import run_pipeline, to_excel_bytes
from datetime import date

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


st.set_page_config(
    page_title="Termination Forms Report",
    page_icon="📋",
    layout="wide",
)

st.title("Termination Forms Report")
st.caption(
    "Upload the Sage file and the Form responses file to generate the report. "
    "Only End of Assignment terminations are included."
)


col1, col2 = st.columns(2)

with col1:
    terminations_file = st.file_uploader(
        "Sage file (All Pending HR Requests)", type=["xlsx"]
    )

with col2:
    responses_file = st.file_uploader(
        "Form responses (TTL Expert / Lead)", type=["xlsx"]
    )

if terminations_file is None or responses_file is None:
    st.info("Upload both files to see the report.")
    st.stop()

try:
    results = run_pipeline(terminations_file, responses_file)
except ValueError as e:
    st.error(str(e))
    st.stop()

kpis = results['kpis']

st.divider()

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("Total terminations", kpis['total_terminations'], border=True,
          help="End of Assignment terminations in the Sage file")

k2.metric("Completed forms", kpis['completed_forms'], border=True)

k3.metric("Pending forms", kpis['incomplete_forms'], border=True)

k4.metric("Completion rate", f"{kpis['completion_rate']}%", border=True)

k5.metric("Pending priority", kpis['incomplete_priority'], border=True,
          help="Pending forms where Priority State is Yes")

st.divider()


summary_display = results['summary'].rename(columns={
    'HR Request: Created By': 'Manager',
    'completed_forms': 'Completed',
    'incomplete_forms': 'Pending',
    'incomplete_priority': 'Pending priority',
    'total_forms': 'Total',
    'completion_rate': 'Completion rate',
})

today = date.today().isoformat()

# Botón 1: libro completo, antes de las pestañas
st.download_button(
    "Download full report",
    data=to_excel_bytes({
        "Summary": summary_display,
        "Pending": results['incomplete'],
        "Completed": results['completed'],
    }),
    file_name=f"termination_forms_report_{today}.xlsx",
    mime=XLSX_MIME,
    type="primary",
    on_click="ignore",
)

tab1, tab2, tab3 = st.tabs([
    "Summary by manager",
    f"Pending forms ({kpis['incomplete_forms']})",
    f"Completed forms ({kpis['completed_forms']})",
])

with tab1:
    st.dataframe(
        summary_display,
        hide_index=True,
        column_config={
            "Completion rate": st.column_config.ProgressColumn(
                format="%.1f%%", min_value=0, max_value=100,
            ),
        },
    )
    # Botón 2: resumen, debajo de su tabla
    st.download_button(
        "Download summary",
        data=to_excel_bytes({"Summary": summary_display}),
        file_name=f"summary_by_manager_{today}.xlsx",
        mime=XLSX_MIME,
        on_click="ignore",
    )

with tab2:
    st.dataframe(results['incomplete'], hide_index=True)
    # Botón 3: pendientes, debajo de su tabla
    st.download_button(
        "Download pending forms",
        data=to_excel_bytes({"Pending": results['incomplete']}),
        file_name=f"pending_forms_{today}.xlsx",
        mime=XLSX_MIME,
        on_click="ignore",
    )

with tab3:
    st.dataframe(results['completed'], hide_index=True)
    # Botón 4: completados, debajo de su tabla
    st.download_button(
        "Download completed forms",
        data=to_excel_bytes({"Completed": results['completed']}),
        file_name=f"completed_forms_{today}.xlsx",
        mime=XLSX_MIME,
        on_click="ignore",
    )