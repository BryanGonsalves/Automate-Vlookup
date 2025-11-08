import io
from typing import Optional, Tuple

import pandas as pd
import streamlit as st


def load_spreadsheet(uploaded_file: Optional[st.runtime.uploaded_file_manager.UploadedFile]) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """Load CSV or Excel file into DataFrame; return (df, error_message)."""
    if uploaded_file is None:
        return None, None

    try:
        filename = uploaded_file.name.lower()
        if filename.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            return None, "Unsupported file type. Please upload CSV or Excel files."

        if df.empty:
            return None, "Uploaded file has no rows."

        # Drop columns with duplicated names to avoid ambiguous selectors later.
        df = df.loc[:, ~df.columns.duplicated()]
        return df, None
    except Exception as exc:  # pragma: no cover - surfaced to UI
        return None, f"Could not read file: {exc}"


def merge_frames(left_df: pd.DataFrame, right_df: pd.DataFrame, left_key: str, right_key: str, value_cols: list[str]) -> pd.DataFrame:
    """Perform a left join mimicking VLOOKUP/XLOOKUP semantics."""
    lookup_cols = [col for col in value_cols if col in right_df.columns and col != right_key]
    if not lookup_cols:
        raise ValueError("No valid lookup columns selected.")

    trimmed_right = right_df[[right_key] + lookup_cols].drop_duplicates(subset=right_key, keep="first")
    merged = left_df.merge(trimmed_right, how="left", left_on=left_key, right_on=right_key, suffixes=("", "_lookup"))
    if right_key not in lookup_cols:
        merged = merged.drop(columns=[right_key])
    return merged


def order_columns(base_cols: list[str], new_cols: list[str], insert_after: str) -> list[str]:
    """Return column ordering where new columns are inserted after the chosen base column."""
    if insert_after == "(At beginning)":
        idx = 0
    elif insert_after == "(At end)":
        idx = len(base_cols)
    else:
        idx = base_cols.index(insert_after) + 1
    return base_cols[:idx] + new_cols + base_cols[idx:]


def main() -> None:
    st.set_page_config(page_title="Lookup Automator", page_icon="🔍", layout="wide")
    st.title("Lookup Automator")
    st.write(
        "Upload two spreadsheets, choose the key columns, and select which values to bring over. "
        "This replicates Excel's VLOOKUP/XLOOKUP with a friendlier workflow."
    )

    st.sidebar.header("Upload source files")
    file_a = st.sidebar.file_uploader("Base file (receives the lookup values)", type=["csv", "xlsx", "xls"], key="file_a")
    file_b = st.sidebar.file_uploader("Lookup file (contains the reference values)", type=["csv", "xlsx", "xls"], key="file_b")

    df_a, err_a = load_spreadsheet(file_a)
    df_b, err_b = load_spreadsheet(file_b)

    for label, err in (("Base file", err_a), ("Lookup file", err_b)):
        if err:
            st.error(f"{label}: {err}")

    if df_a is None or df_b is None:
        st.info("Upload both files to continue.")
        return

    st.subheader("Column selection")
    col1, col2 = st.columns(2)
    with col1:
        left_key = st.selectbox("Primary key in base file", options=df_a.columns, help="Rows are matched using this column.")
    with col2:
        right_key = st.selectbox("Primary key in lookup file", options=df_b.columns, help="Must contain the keys present in the base file.")

    value_cols = st.multiselect(
        "Columns to fetch from lookup file",
        options=[col for col in df_b.columns if col != right_key],
        help="Select one or more columns to append to the base file.",
    )

    insert_after_options = ["(At beginning)"] + list(df_a.columns) + ["(At end)"]
    insert_after = st.selectbox(
        "Insert fetched columns after",
        options=insert_after_options,
        index=len(insert_after_options) - 1,
        help="Choose where the fetched columns should appear inside the base file.",
    )

    if st.button("Run lookup", type="primary", disabled=not value_cols):
        try:
            merged = merge_frames(df_a, df_b, left_key, right_key, value_cols)
            base_columns = list(df_a.columns)
            new_columns = [col for col in merged.columns if col not in base_columns]
            ordered_columns = order_columns(base_columns, new_columns, insert_after)
            merged = merged[ordered_columns]

            st.success(f"Lookup completed. {merged.shape[0]} rows processed.")
            st.dataframe(merged, use_container_width=True)

            csv_buffer = io.StringIO()
            merged.to_csv(csv_buffer, index=False)
            st.download_button(
                label="Download results as CSV",
                data=csv_buffer.getvalue(),
                file_name="lookup_results.csv",
                mime="text/csv",
            )
        except Exception as exc:
            st.error(f"Lookup failed: {exc}")
    else:
        st.caption("Select at least one column to fetch and click 'Run lookup'.")


if __name__ == "__main__":
    main()
