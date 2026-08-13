from datetime import datetime

import pandas as pd
import streamlit as st

from outlier_engine import analyze_outliers, build_excel, parse_pasted_data


st.set_page_config(
    page_title="Housing Sale Outlier Highlighter",
    page_icon="🏠",
    layout="centered",
)

st.title("Housing Sale Outlier Highlighter")
st.caption(
    "Paste Close Price and Living Area directly from Excel, analyze the listings, "
    "then download the same two columns with statistical outliers highlighted."
)


METHOD_LABELS = {
    "Either method": "either",
    "Price vs. living area only": "robust_regression",
    "Price per sq. ft. only": "price_per_sqft_iqr",
}

FIRST_ROW_LABELS = {
    "Auto-detect": "auto",
    "First row is headers": "header",
    "First row is data": "data",
}


with st.form("outlier_form"):
    pasted_data = st.text_area(
        "Paste both Excel columns here",
        height=300,
        placeholder=(
            "Close Price\tLiving Area\n"
            "841,990\t1818\n"
            "800,000\t1818\n"
            "792,000\t1818"
        ),
        help="In Excel, select the Close Price and Living Area columns together, copy, and paste here.",
    )

    with st.expander("Advanced settings"):
        method_label = st.selectbox(
            "Outlier rule",
            options=list(METHOD_LABELS.keys()),
            index=0,
            help=(
                "Either method flags a listing when either the robust price-vs-size model "
                "or the price-per-square-foot IQR test identifies it."
            ),
        )

        robust_z_threshold = st.number_input(
            "Robust Z threshold",
            min_value=0,
            max_value=8.0,
            value=3.5,
            step=0.1,
            help="Lower values flag more price-vs-size outliers.",
        )

        iqr_multiplier = st.number_input(
            "Price/sq.-ft. IQR multiplier",
            min_value=0.5,
            max_value=5.0,
            value=1.5,
            step=0.1,
            help="Lower values flag more price-per-square-foot outliers.",
        )

        flag_extreme_living_area = st.checkbox(
            "Also flag unusually large or small living areas",
            value=False,
        )

        first_row_label = st.selectbox(
            "First pasted row",
            options=list(FIRST_ROW_LABELS.keys()),
            index=0,
        )

    submitted = st.form_submit_button(
        "Analyze outliers",
        type="primary",
        width="stretch",
    )


if submitted:
    try:
        df, headers = parse_pasted_data(
            pasted_data,
            first_row_mode=FIRST_ROW_LABELS[first_row_label],
        )

        results, diagnostics = analyze_outliers(
            df,
            robust_z_threshold=float(robust_z_threshold),
            iqr_multiplier=float(iqr_multiplier),
            outlier_method=METHOD_LABELS[method_label],
            flag_extreme_living_area=flag_extreme_living_area,
        )

        excel_bytes = build_excel(results, headers)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Housing_Outliers_Highlighted_{timestamp}.xlsx"

        st.session_state["analysis_results"] = results
        st.session_state["analysis_headers"] = headers
        st.session_state["analysis_diagnostics"] = diagnostics
        st.session_state["analysis_excel"] = excel_bytes
        st.session_state["analysis_filename"] = filename

    except ValueError as exc:
        st.error(str(exc))
        for key in (
            "analysis_results",
            "analysis_headers",
            "analysis_diagnostics",
            "analysis_excel",
            "analysis_filename",
        ):
            st.session_state.pop(key, None)


if "analysis_results" in st.session_state:
    results = st.session_state["analysis_results"]
    headers = st.session_state["analysis_headers"]
    diagnostics = st.session_state["analysis_diagnostics"]

    st.subheader("Output")

    metric_col1, metric_col2 = st.columns(2)
    metric_col1.metric("Listings analyzed", f"{diagnostics['listings']:,}")
    metric_col2.metric("Outliers highlighted", f"{diagnostics['outliers']:,}")

    preview = results[["Close Price", "Living Area"]].round(0).astype(int).copy()
    preview.columns = headers

    def highlight_outlier_cells(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        outlier_rows = results.loc[data.index, "Outlier"]
        styles.loc[outlier_rows, :] = "background-color: #00B0F0; color: black;"
        return styles

    styled_preview = (
        preview.style
        .format({headers[0]: "{:,.0f}", headers[1]: "{:.0f}"})
        .apply(highlight_outlier_cells, axis=None)
    )

    st.dataframe(
        styled_preview,
        hide_index=True,
        width="stretch",
        height=520,
    )

    if diagnostics["outliers"] == 0:
        st.info(
            "No listings were flagged with the current settings. "
            "You can lower the thresholds in Advanced settings and analyze again if needed."
        )

    st.download_button(
        "Download highlighted Excel file",
        data=st.session_state["analysis_excel"],
        file_name=st.session_state["analysis_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
        on_click="ignore",
        icon=":material/download:",
    )

    with st.expander("Current detection details"):
        st.write(
            f"Regression outliers: **{diagnostics['regression_outliers']}**  \n"
            f"Price-per-square-foot outliers: **{diagnostics['ppsf_outliers']}**"
        )
        st.caption(
            "The downloaded workbook contains one worksheet only. Every listing is retained "
            "in its original order. Outlier rows receive only a blue fill; data cells are "
            "Aptos Narrow 11 pt and unbolded."
        )
