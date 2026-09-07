import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# PAGE CONFIG
st.set_page_config(
    page_title="DB Reconciliation System",
    page_icon="📊",
    layout="wide"
)


# FUNCTIONS
def read_file(uploaded_file):
    """
    Read CSV or Excel file.

    If the first row is completely empty,
    skip it and use the next row as the header.
    """

    if uploaded_file.name.lower().endswith(".csv"):

        df = pd.read_csv(
            uploaded_file,
            header=None
        )

    elif uploaded_file.name.lower().endswith(
        (".xlsx", ".xls")
    ):

        df = pd.read_excel(
            uploaded_file,
            header=None
        )

    else:

        raise ValueError(
            "Unsupported file format"
        )

    # Check if first row is completely empty
    if len(df) > 0:

        first_row = df.iloc[0]

        first_row_empty = (
            first_row.isna()
            |
            first_row.astype(str)
            .str.strip()
            .eq("")
        ).all()

        if first_row_empty:

            df = df.iloc[1:].reset_index(
                drop=True
            )

    # First remaining row becomes header
    df.columns = (
        df.iloc[0]
        .astype(str)
        .str.strip()
    )

    # Remove header row
    df = df.iloc[1:].reset_index(
        drop=True
    )

    return df


#============================================================

def find_duplicate_ids(df, id_column):
    """
    Find all rows belonging to duplicate IDs.
    """

    duplicate_mask = df[id_column].duplicated(
        keep=False
    )

    duplicate_data = df[
        duplicate_mask
    ].copy()

    duplicate_ids = (
        duplicate_data[id_column]
        .dropna()
        .unique()
    )

    return duplicate_data, duplicate_ids


#============================================================

def delete_duplicate_ids(df, id_column):
    """
    Delete ALL rows belonging to duplicated IDs.

    Example:

    A001
    A001
    A002

    becomes:

    A002
    """

    duplicate_mask = df[id_column].duplicated(
        keep=False
    )

    cleaned_df = df[
        ~duplicate_mask
    ].copy()

    cleaned_df.reset_index(
        drop=True,
        inplace=True
    )

    return cleaned_df


#============================================================

def create_comparison(
    df1,
    df2,
    id_col_df1,
    value_col_df1,
    id_col_df2,
    value_col_df2
):
    """
    Create reconciliation comparison.
    """

    # Select ID and Value
    df1_compare = df1[
        [id_col_df1, value_col_df1]
    ].copy()

    df2_compare = df2[
        [id_col_df2, value_col_df2]
    ].copy()

    # Standardize column names
    df1_compare.rename(
        columns={
            id_col_df1: "ID",
            value_col_df1: "Value_df1"
        },
        inplace=True
    )

    df2_compare.rename(
        columns={
            id_col_df2: "ID",
            value_col_df2: "Value_df2"
        },
        inplace=True
    )

    # Clean IDs
    df1_compare["ID"] = (
        df1_compare["ID"]
        .astype(str)
        .str.strip()
    )

    df2_compare["ID"] = (
        df2_compare["ID"]
        .astype(str)
        .str.strip()
    )

    # Empty ID -> NaN
    df1_compare["ID"] = (
        df1_compare["ID"]
        .replace(
            ["", "nan", "None"],
            np.nan
        )
    )

    df2_compare["ID"] = (
        df2_compare["ID"]
        .replace(
            ["", "nan", "None"],
            np.nan
        )
    )

    # Remove rows without ID
    df1_compare.dropna(
        subset=["ID"],
        inplace=True
    )

    df2_compare.dropna(
        subset=["ID"],
        inplace=True
    )

    # Convert Value to numeric
    df1_compare["Value_df1"] = pd.to_numeric(
        df1_compare["Value_df1"],
        errors="coerce"
    )

    df2_compare["Value_df2"] = pd.to_numeric(
        df2_compare["Value_df2"],
        errors="coerce"
    )

    # OUTER MERGE
    comparison = df1_compare.merge(
        df2_compare,
        on="ID",
        how="outer"
    )

    # GAP
    comparison["Gap"] = (
        comparison["Value_df1"].fillna(0)
        -
        comparison["Value_df2"].fillna(0)
    )

    # NOTE
    comparison["Note"] = np.where(
        comparison["Gap"] == 0,
        "Settle",
        "Unsettle"
    )

    # Column order
    comparison = comparison[
        [
            "ID",
            "Value_df1",
            "Value_df2",
            "Gap",
            "Note"
        ]
    ]

    return comparison

#============================================================

def create_summary(comparison):

    total_trx = comparison["ID"].nunique()

    total_value_df1 = (
        comparison["Value_df1"]
        .fillna(0)
        .sum()
    )

    total_value_df2 = (
        comparison["Value_df2"]
        .fillna(0)
        .sum()
    )

    total_gap = (
        comparison["Gap"]
        .fillna(0)
        .sum()
    )

    settle = (
        comparison["Note"] == "Settle"
    ).sum()

    unsettle = (
        comparison["Note"] == "Unsettle"
    ).sum()

    summary = pd.DataFrame({
        "Summary": [
            "Total Trx",
            "Total Value - DB A",
            "Total Value - DB B",
            "Total Gap",
            "Settle",
            "Unsettle"
        ],
        "Value": [
            total_trx,
            total_value_df1,
            total_value_df2,
            total_gap,
            settle,
            unsettle
        ]
    })

    return summary

# ============================================================

def create_excel(comparison, summary):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # Sheet 1
        comparison.to_excel(
            writer,
            sheet_name="Comparison",
            index=False
        )


        # Sheet 2
        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        workbook = writer.book

        # Comparison formatting
        ws = workbook["Comparison"]

        ws.freeze_panes = "A2"

        ws.auto_filter.ref = (
            ws.dimensions
        )

        widths = {
            "A": 25,
            "B": 18,
            "C": 18,
            "D": 18,
            "E": 15
        }

        for column, width in widths.items():

            ws.column_dimensions[
                column
            ].width = width

        for row in ws.iter_rows(
            min_row=2,
            min_col=2,
            max_col=4
        ):

            for cell in row:

                cell.number_format = (
                    '#,##0.00'
                )

        # Summary formatting
        ws_summary = workbook["Summary"]

        ws_summary.column_dimensions[
            "A"
        ].width = 30

        ws_summary.column_dimensions[
            "B"
        ].width = 25

    output.seek(0)

    return output


# SESSION STATE
if "df1" not in st.session_state:
    st.session_state["df1"] = None

if "df2" not in st.session_state:
    st.session_state["df2"] = None

if "mapping_done" not in st.session_state:
    st.session_state["mapping_done"] = False

if "comparison" not in st.session_state:
    st.session_state["comparison"] = None

if "summary" not in st.session_state:
    st.session_state["summary"] = None

if "file_a_name" not in st.session_state:
    st.session_state["file_a_name"] = None

if "file_b_name" not in st.session_state:
    st.session_state["file_b_name"] = None


# TITLE
st.title(
    "📊 DB Reconciliation System"
)

st.write(
    "Upload two databases, map the ID and Value columns, "
    "clean duplicate IDs, and reconcile the data."
)


# 1. UPLOAD DATABASE
st.subheader(
    "Upload Database"
)

col1, col2 = st.columns(2)


with col1:

    st.markdown(
        "### Database A"
    )

    file_df1 = st.file_uploader(
        "Upload Database A",
        type=[
            "xlsx",
            "xls",
            "csv"
        ],
        key="upload_df1"
    )


with col2:

    st.markdown(
        "### Database B"
    )

    file_df2 = st.file_uploader(
        "Upload Database B",
        type=[
            "xlsx",
            "xls",
            "csv"
        ],
        key="upload_df2"
    )


# START PROCESS WHEN BOTH FILES ARE UPLOADED
if file_df1 and file_df2:

    # READ DATABASE A
    if (
        st.session_state["file_a_name"]
        != file_df1.name
    ):

        st.session_state["df1"] = (
            read_file(file_df1)
        )

        st.session_state[
            "file_a_name"
        ] = file_df1.name

        # New upload = reset mapping
        st.session_state[
            "mapping_done"
        ] = False

        st.session_state[
            "comparison"
        ] = None

        st.session_state[
            "summary"
        ] = None


    # READ DATABASE B
    if (
        st.session_state["file_b_name"]
        != file_df2.name
    ):

        st.session_state["df2"] = (
            read_file(file_df2)
        )

        st.session_state[
            "file_b_name"
        ] = file_df2.name

        # New upload = reset mapping
        st.session_state[
            "mapping_done"
        ] = False

        st.session_state[
            "comparison"
        ] = None

        st.session_state[
            "summary"
        ] = None


    # Get current data
    df1 = st.session_state["df1"]
    df2 = st.session_state["df2"]

    # 2. CURRENT DATA PREVIEW
    st.subheader(
        "Current Data Preview"
    )

    st.write(
        "This preview always shows the newest version "
        "of the uploaded data."
    )

    col1, col2 = st.columns(2)


    with col1:

        st.markdown(
            "### Database A"
        )

        st.write(
            f"Rows: **{len(df1):,}**"
        )

        st.write(
            "Available fields:"
        )

        st.code(
            ", ".join(
                df1.columns.astype(str)
            )
        )

        st.dataframe(
            df1.head(),
            use_container_width=True
        )


    with col2:

        st.markdown(
            "### Database B"
        )

        st.write(
            f"Rows: **{len(df2):,}**"
        )

        st.write(
            "Available fields:"
        )

        st.code(
            ", ".join(
                df2.columns.astype(str)
            )
        )

        st.dataframe(
            df2.head(),
            use_container_width=True
        )

    # 3. MAPPING
    st.subheader(
        "Map ID and Value"
    )
    if not st.session_state[
        "mapping_done"
    ]:

        col1, col2 = st.columns(2)

        # DB A
        with col1:

            st.markdown(
                "### Database A"
            )

            id_col_df1 = st.selectbox(
                "ID Column",
                df1.columns.tolist(),
                key="map_id_df1"
            )

            value_col_df1 = st.selectbox(
                "Value Column",
                df1.columns.tolist(),
                key="map_value_df1"
            )

        # DB B
        with col2:

            st.markdown(
                "### Database B"
            )

            id_col_df2 = st.selectbox(
                "ID Column",
                df2.columns.tolist(),
                key="map_id_df2"
            )

            value_col_df2 = st.selectbox(
                "Value Column",
                df2.columns.tolist(),
                key="map_value_df2"
            )

        # Confirm Mapping
        if st.button(
            "✅ Confirm Mapping",
            type="primary",
            use_container_width=True
        ):

            st.session_state[
                "id_col_df1"
            ] = id_col_df1

            st.session_state[
                "value_col_df1"
            ] = value_col_df1

            st.session_state[
                "id_col_df2"
            ] = id_col_df2

            st.session_state[
                "value_col_df2"
            ] = value_col_df2

            st.session_state[
                "mapping_done"
            ] = True

            st.rerun()

    # AFTER MAPPING
    if st.session_state[
        "mapping_done"
    ]:

        id_col_df1 = st.session_state[
            "id_col_df1"
        ]

        value_col_df1 = st.session_state[
            "value_col_df1"
        ]

        id_col_df2 = st.session_state[
            "id_col_df2"
        ]

        value_col_df2 = st.session_state[
            "value_col_df2"
        ]

        # SHOW CURRENT MAPPING
        st.success(
            "Mapping confirmed."
        )

        col1, col2 = st.columns(2)


        with col1:

            st.markdown(
                "### Database A"
            )

            st.write(
                f"**ID:** `{id_col_df1}`"
            )

            st.write(
                f"**Value:** `{value_col_df1}`"
            )


        with col2:

            st.markdown(
                "### Database B"
            )

            st.write(
                f"**ID:** `{id_col_df2}`"
            )

            st.write(
                f"**Value:** `{value_col_df2}`"
            )

        # 4. DUPLICATE CHECK
        st.subheader(
            "Duplicate Check"
        )


        duplicate_data_df1, duplicate_ids_df1 = (
            find_duplicate_ids(
                df1,
                id_col_df1
            )
        )


        duplicate_data_df2, duplicate_ids_df2 = (
            find_duplicate_ids(
                df2,
                id_col_df2
            )
        )


        duplicate_count_df1 = (
            len(duplicate_ids_df1)
        )

        duplicate_count_df2 = (
            len(duplicate_ids_df2)
        )

        # DUPLICATE STATUS
        col1, col2 = st.columns(2)
        with col1:

            st.markdown(
                "### Database A"
            )

            if duplicate_count_df1 > 0:

                st.warning(
                    f"⚠️ "
                    f"**{duplicate_count_df1:,} "
                    f"duplicate IDs** found."
                )

                st.write(
                    f"Rows affected: "
                    f"**{len(duplicate_data_df1):,}**"
                )

            else:

                st.success(
                    "✅ No duplicate IDs found."
                )


        with col2:

            st.markdown(
                "### Database B"
            )

            if duplicate_count_df2 > 0:

                st.warning(
                    f"⚠️ "
                    f"**{duplicate_count_df2:,} "
                    f"duplicate IDs** found."
                )

                st.write(
                    f"Rows affected: "
                    f"**{len(duplicate_data_df2):,}**"
                )

            else:

                st.success(
                    "✅ No duplicate IDs found."
                )

        # SHOW DUPLICATE DATA
        if duplicate_count_df1 > 0:

            with st.expander(
                "View Duplicate Data - DB A"
            ):

                st.dataframe(
                    duplicate_data_df1,
                    use_container_width=True
                )


        if duplicate_count_df2 > 0:

            with st.expander(
                "View Duplicate Data - DB B"
            ):

                st.dataframe(
                    duplicate_data_df2,
                    use_container_width=True
                )

        # DELETE ALL DUPLICATES
        duplicates_exist = (
            duplicate_count_df1 > 0
            or
            duplicate_count_df2 > 0
        )


        if duplicates_exist:

            st.warning(
                "Duplicate IDs must be cleaned "
                "before reconciliation."
            )


            if st.button(
                "🗑️ Delete All Duplicates",
                type="primary",
                use_container_width=True
            ):

                # Clean DB A if duplicates exist
                if duplicate_count_df1 > 0:

                    cleaned_df1 = (
                        delete_duplicate_ids(
                            df1,
                            id_col_df1
                        )
                    )

                    st.session_state[
                        "df1"
                    ] = cleaned_df1


                # Clean DB B if duplicates exist
                if duplicate_count_df2 > 0:

                    cleaned_df2 = (
                        delete_duplicate_ids(
                            df2,
                            id_col_df2
                        )
                    )

                    st.session_state[
                        "df2"
                    ] = cleaned_df2

                # Remove previous reconciliation
                st.session_state[
                    "comparison"
                ] = None

                st.session_state[
                    "summary"
                ] = None

                # Refresh
                st.rerun()

        # 6. RECONCILIATION
        if not duplicates_exist:

            st.subheader(
                "Reconciliation"
            )

            st.success(
                "✅ Both databases are clean. "
                "Ready for reconciliation."
            )


            if st.button(
                "🔍 Reconcile Databases",
                type="primary",
                use_container_width=True
            ):

                comparison = create_comparison(
                    df1,
                    df2,
                    id_col_df1,
                    value_col_df1,
                    id_col_df2,
                    value_col_df2
                )

                summary = create_summary(
                    comparison
                )

                st.session_state[
                    "comparison"
                ] = comparison

                st.session_state[
                    "summary"
                ] = summary

                st.rerun()

        # 7. RECONCILIATION RESULT
        if (
            st.session_state[
                "comparison"
            ] is not None
        ):

            comparison = (
                st.session_state[
                    "comparison"
                ]
            )

            summary = (
                st.session_state[
                    "summary"
                ]
            )

            # SUMMARY
            st.subheader(
                "Reconciliation Summary"
            )


            total_trx = summary.loc[
                summary["Summary"]
                == "Total Trx",
                "Value"
            ].iloc[0]


            total_value_df1 = summary.loc[
                summary["Summary"]
                == "Total Value - DB A",
                "Value"
            ].iloc[0]


            total_value_df2 = summary.loc[
                summary["Summary"]
                == "Total Value - DB B",
                "Value"
            ].iloc[0]


            total_gap = summary.loc[
                summary["Summary"]
                == "Total Gap",
                "Value"
            ].iloc[0]


            settle_count = summary.loc[
                summary["Summary"]
                == "Settle",
                "Value"
            ].iloc[0]


            unsettle_count = summary.loc[
                summary["Summary"]
                == "Unsettle",
                "Value"
            ].iloc[0]

            # Metrics
            col1, col2, col3, col4 = (
                st.columns(4)
            )


            col1.metric(
                "Total Trx",
                f"{total_trx:,.0f}"
            )


            col2.metric(
                "Settle",
                f"{settle_count:,.0f}"
            )


            col3.metric(
                "Unsettle",
                f"{unsettle_count:,.0f}"
            )


            col4.metric(
                "Total Gap",
                f"{total_gap:,.2f}"
            )

            # Value Summary
            st.write(
                "### Value Summary"
            )


            col1, col2 = st.columns(2)


            col1.metric(
                "Total Value - DB A",
                f"{total_value_df1:,.2f}"
            )


            col2.metric(
                "Total Value - DB B",
                f"{total_value_df2:,.2f}"
            )

            # COMPARISON RESULT
            st.subheader(
                "Reconciliation Result"
            )
            
            st.dataframe(
                comparison,
                use_container_width=True,
                height=500
            )

            # EXPORT
            st.subheader("Export Result")
            
            excel_file = create_excel(
                comparison,
                summary
            )
    
            st.download_button(
                label="📥 Download Excel Result",
                data=excel_file,
                file_name="DB_Reconciliation_Result.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )