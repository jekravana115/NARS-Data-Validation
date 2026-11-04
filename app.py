from __future__ import annotations

import re
import zipfile
from datetime import date
from pathlib import Path

import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException

from nars_validation import (
    MAX_FILE_BYTES,
    MAX_TOTAL_RECORDS,
    RULESET_VERSION,
    VALIDITY_FIELDS_BY_BASIS,
    ValidationInputError,
    analyze_workbook,
    build_export,
    duplicate_filenames,
    load_lookups,
    validity_failure_counts,
)


st.set_page_config(
    page_title="NARS Data Validation",
    page_icon="✓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

RESULTS_SCHEMA_VERSION = "validity-bases-v2"
if st.session_state.get("results_schema_version") != RESULTS_SCHEMA_VERSION:
    for key in (
        "validation_results",
        "validation_errors",
        "validation_exports",
        "validation_export",
    ):
        st.session_state.pop(key, None)
    st.session_state["results_schema_version"] = RESULTS_SCHEMA_VERSION
if "reference_date" not in st.session_state:
    st.session_state["reference_date"] = date.today()
if "editing_reference_date" not in st.session_state:
    st.session_state["editing_reference_date"] = False

st.markdown(
    """
    <style>
    .block-container { max-width: 1440px; padding-top: 4rem; }
    .st-key-nars_header {
      background: #173b5e; border-radius: 12px;
      margin-bottom: 1.5rem; padding: 0.8rem 1.5rem;
    }
    .st-key-nars_header h1 { color: #ffffff; margin-bottom: 0; }
    .st-key-nars_header .stPopover [data-testid="stPopoverButton"] {
      background: transparent; border: 0; box-shadow: none; color: #ffffff;
    }
    .st-key-nars_header .stPopover [data-testid="stPopoverButton"]:hover {
      background: rgba(255, 255, 255, 0.12); border: 0; box-shadow: none;
    }
    .st-key-nars_header .stPopover [data-testid="stPopoverButton"]:focus,
    .st-key-nars_header .stPopover [data-testid="stPopoverButton"]:focus-visible {
      outline: none; box-shadow: none;
    }
    div[data-testid="stFileUploader"] label[data-testid="stWidgetLabel"] > div:last-child {
      flex: 0 0 auto; width: auto; margin-left: 6px;
    }
    .stPopover [data-testid="stTooltipHoverTarget"] { width: 100%; }
    .stPopover [data-testid="stPopoverButton"] { color: #31333f; }
    .stPopover [data-testid="stPopoverButton"] > div:last-child { display: none; }
    .st-key-agency_filter { width: min(235px, 100%); }
    .st-key-agency_filter > div { width: 100% !important; }
    .st-key-agency_filter [data-baseweb="select"] { width: 100%; }
    .st-key-validate_workbooks button { width: 235px; }
    .st-key-clear_results button { width: 235px; }
    [data-testid="stHeaderActionElements"] { display: none; }
    [data-testid="stMetricDeltaIcon-Up"],
    [data-testid="stMetricDeltaIcon-Down"] { display: none; }
    div[data-testid="stMetric"] {
      background: white; border: 1px solid #dfe7ef;
      border-radius: 12px; padding: 14px 18px;
    }
    .nars-entry-metric {
      min-height: 103.5px; padding: 14px 12px; background: white;
      border: 1px solid #dfe7ef; border-radius: 12px;
    }
    .nars-entry-label { margin-bottom: 8px; color: #31333f; font-size: 14px; }
    .nars-entry-value { display: flex; align-items: baseline; gap: 4px; white-space: nowrap; }
    .nars-entry-number { color: #31333f; font-size: 1.5rem; line-height: 1.2; }
    .nars-entry-context { color: #5b6b80; font-size: .68rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.container(key="nars_header"):
    title_column, settings_column = st.columns(
        [0.85, 0.15],
        vertical_alignment="center",
    )
    with title_column:
        st.title("NARS Data Validation")
    with settings_column:
        with st.popover("", help="Validation settings", icon=":material/settings:"):
            as_of = st.session_state["reference_date"]
            st.caption(f"Insurance reference date: {as_of.isoformat()}")
            if not st.session_state["editing_reference_date"]:
                if st.button("Change reference date"):
                    st.session_state["editing_reference_date"] = True
                    st.rerun()
            else:
                as_of = st.date_input(
                    "Insurance reference date",
                    value=as_of,
                    key="reference_date_picker",
                    help="Used to identify expired insurance.",
                )
                st.session_state["reference_date"] = as_of
                if st.button("Done"):
                    st.session_state["editing_reference_date"] = False
                    st.rerun()
            st.divider()
            st.caption(f"Rule set as of: {RULESET_VERSION}")

uploaded_files = st.file_uploader(
    "Upload NARS workbook(s)",
    type=["xlsx"],
    accept_multiple_files=True,
    help=(
        "Upload workbook or a batch from different agencies. "
        "Use agency names in the filenames (e.g., DOH.xlsx). "
        "Each workbook must use the default NARS data fields."
    ),
)

if uploaded_files:
    total_bytes = sum(file.size for file in uploaded_files)
    duplicate_names = duplicate_filenames([file.name for file in uploaded_files])
    if duplicate_names:
        duplicate_list = ", ".join(f"`{name}`" for name in duplicate_names)
        st.error(f"Duplicate workbook filename(s): {duplicate_list}.")
    if any(file.size > MAX_FILE_BYTES for file in uploaded_files):
        st.error("Each workbook must be 25 MiB or smaller.")
    elif total_bytes > 100 * 1024 * 1024:
        st.error("The total upload size must be 100 MiB or smaller.")
    else:
        st.caption(
            f"{len(uploaded_files)} workbook(s) selected · "
            f"{total_bytes / (1024 * 1024):.1f} MiB total"
        )

if st.button(
    "Validate uploaded workbook(s)",
    key="validate_workbooks",
    type="primary",
    disabled=not uploaded_files
    or bool(duplicate_filenames([file.name for file in uploaded_files]))
    or any(file.size > MAX_FILE_BYTES for file in uploaded_files)
    or sum(file.size for file in uploaded_files) > 100 * 1024 * 1024,
):
    results = []
    errors = []
    try:
        lookups = load_lookups()
    except (OSError, ValueError, KeyError) as error:
        st.error(f"Could not load the validation codebooks: {error}")
        st.stop()

    progress = st.progress(0, text="Preparing validation…")
    total_processed_records = 0
    for index, file in enumerate(uploaded_files, start=1):
        file_name = Path(file.name).name
        agency = re.sub(r"[^A-Za-z0-9_-]+", "_", Path(file_name).stem).strip("_")
        if not agency:
            agency = f"Dataset_{index}"
        try:
            result = analyze_workbook(
                file.getvalue(),
                agency=agency,
                source_file=file_name,
                as_of=as_of,
                lookups=lookups,
            )
            if total_processed_records + result["summary"]["records"] > MAX_TOTAL_RECORDS:
                errors.append(
                    (
                        file_name,
                        f"Validating this workbook would exceed the "
                        f"{MAX_TOTAL_RECORDS:,}-record batch limit.",
                    )
                )
            else:
                total_processed_records += result["summary"]["records"]
                results.append(result)
        except (
            ValidationInputError,
            InvalidFileException,
            zipfile.BadZipFile,
            OSError,
            ValueError,
        ) as error:
            errors.append((file_name, str(error)))
        progress.progress(
            index / len(uploaded_files),
            text=f"Checked {index} of {len(uploaded_files)} workbook(s)…",
        )
    progress.empty()

    st.session_state["validation_results"] = results
    st.session_state["validation_errors"] = errors
    st.session_state["validation_exports"] = {}

results = st.session_state.get("validation_results", [])
errors = st.session_state.get("validation_errors", [])

for file_name, message in errors:
    st.error(f"{file_name}: {message}")

if results:
    validity_basis = "legacy18"
    basis_fields = VALIDITY_FIELDS_BY_BASIS[validity_basis]
    agency_names = sorted(
        {result["summary"]["agency"] for result in results},
        key=str.casefold,
    )
    agency_filter_options = ["*", *agency_names]
    if st.session_state.get("agency_filter") not in agency_filter_options:
        st.session_state["agency_filter"] = "*"
    selected_agency = st.selectbox(
        "Filter by agency",
        options=agency_filter_options,
        format_func=lambda agency: "All agencies" if agency == "*" else agency,
        key="agency_filter",
        disabled=len(agency_names) == 1,
    )
    filtered_results = (
        results
        if selected_agency == "*"
        else [
            result
            for result in results
            if result["summary"]["agency"] == selected_agency
        ]
    )
    agency_label = "all agencies" if selected_agency == "*" else selected_agency
    summaries = [result["summary"] for result in filtered_results]
    total_records = sum(row["records"] for row in summaries)
    total_valid = sum(row["valid_records"] for row in summaries)
    total_valid_field_entries = sum(
        row["validity_by_basis"][validity_basis]["valid_field_entries"]
        for row in summaries
    )
    total_field_checks = sum(
        row["validity_by_basis"][validity_basis]["field_checks"]
        for row in summaries
    )
    total_validity_rate = (
        total_valid_field_entries / total_field_checks * 100
        if total_field_checks
        else 0
    )

    st.divider()
    st.subheader("Validation overview")
    st.caption(
        f"Showing {len(filtered_results)} of {len(results)} validated workbook(s) "
        f"for {agency_label} as of {as_of.isoformat()}."
    )
    metrics = st.columns(4)
    metrics[0].metric("Assets checked", f"{total_records:,}")
    metrics[1].metric("Fully valid assets", f"{total_valid:,}")
    metrics[2].markdown(
        '<div class="nars-entry-metric">'
        '<div class="nars-entry-label">Valid field entries</div>'
        '<div class="nars-entry-value">'
        f'<span class="nars-entry-number">{total_valid_field_entries:,}</span>'
        f'<span class="nars-entry-context">out of {total_field_checks:,} entries</span>'
        "</div></div>",
        unsafe_allow_html=True,
    )
    metrics[3].metric("Validity rate", f"{total_validity_rate:.1f}%")

    st.subheader("Agency summary")
    st.dataframe(
        [
            {
                key: value
                for key, value in summary.items()
                if key not in {"validity_by_basis", "total_findings", "duplicate_records"}
            }
            | summary["validity_by_basis"][validity_basis]
            for summary in summaries
        ],
        hide_index=True,
        use_container_width=True,
        column_config={
            "agency": "Agency",
            "source_file": "Source workbook",
            "records": st.column_config.NumberColumn("Records checked", format="%d"),
            "valid_records": st.column_config.NumberColumn("Fully valid assets", format="%d"),
            "valid_field_entries": st.column_config.NumberColumn(
                "Valid field entries", format="%d"
            ),
            "field_checks": st.column_config.NumberColumn("Field checks", format="%d"),
            "validity_rate_pct": st.column_config.NumberColumn(
                "Field validity", format="%.1f%%"
            ),
            "records_to_review": st.column_config.NumberColumn("Records to review", format="%d"),
            "as_of": "Checked as of",
            "ruleset_version": "Rule set",
        },
    )

    field_validity = [
        row
        for result in filtered_results
        for row in result["field_validity_by_basis"][validity_basis]
    ]
    st.subheader("Validity by field")
    st.dataframe(
        field_validity,
        hide_index=True,
        use_container_width=True,
        column_config={
            "agency": "Agency",
            "field": "Field",
            "valid_entries": st.column_config.NumberColumn("Valid entries", format="%d"),
            "invalid_entries": st.column_config.NumberColumn("Invalid entries", format="%d"),
            "total_entries": st.column_config.NumberColumn("Assets checked", format="%d"),
            "validity_rate_pct": st.column_config.NumberColumn(
                "Validity rate", format="%.1f%%"
            ),
        },
    )

    failing_counts = validity_failure_counts(filtered_results, validity_basis)

    left, right = st.columns([1, 1])
    with left:
        st.subheader("Validation by field and status")
        if failing_counts:
            st.dataframe(
                failing_counts,
                hide_index=True,
                use_container_width=True,
                height=578,
                column_config={
                    "agency": "Agency",
                    "rule": "Field",
                    "status": "Status",
                    "count": st.column_config.NumberColumn("Records", format="%d"),
                },
            )
        else:
            st.success("No check failures were found in the selected workbooks.")

    all_records = [
        row for result in filtered_results for row in result["records"]
    ]
    issue_records = [
        row for result in filtered_results for row in result["issues"]
    ]
    with right:
        st.subheader("Record-level results")
        show_issues_only = st.toggle("Show records to review only", value=True)
        display_rows = issue_records if show_issues_only else all_records
        display_rows = [
            row | row["validity_by_basis"][validity_basis]
            for row in display_rows
        ]
        st.caption(f"Showing {len(display_rows):,} of {len(all_records):,} records.")
        display_columns = [
            "agency",
            "source_row",
            "result",
            "issue_count",
            "valid_field_entries",
            "field_checks",
            "validity_rate_pct",
            "issues",
            "assetname",
            "assetdesc",
            "assettype",
            "province",
            "region",
        ]
        st.dataframe(
            [{key: row.get(key) for key in display_columns} for row in display_rows],
            hide_index=True,
            use_container_width=True,
            height=500,
            column_config={
                "agency": "Agency",
                "source_row": "Source row",
                "result": "Result",
                "issue_count": "Findings",
                "valid_field_entries": "Valid fields",
                "field_checks": "Fields checked",
                "validity_rate_pct": st.column_config.NumberColumn(
                    "Field validity", format="%.1f%%"
                ),
                "issues": "Issues",
                "assetname": "Asset name",
                "assetdesc": "Asset description",
                "assettype": "Asset type",
                "province": "Province",
                "region": "Region",
            },
        )

    with st.expander("Show validity status for each asset and check"):
        asset_field_rows = [
            {
                "agency": row["agency"],
                "source_row": row.get("source_row"),
                "assetname": row["assetname"],
                "valid_field_entries": row["validity_by_basis"][validity_basis][
                    "valid_field_entries"
                ],
                "validity_rate_pct": row["validity_by_basis"][validity_basis][
                    "validity_rate_pct"
                ],
                **{
                    field_label: row[f"{rule}_status"]
                    for _, field_label, rule, _ in basis_fields
                },
            }
            for result in filtered_results
            for row in result["records"]
        ]
        st.dataframe(
            asset_field_rows,
            hide_index=True,
            use_container_width=True,
            height=500,
            column_config={
                "agency": "Agency",
                "source_row": "Source row",
                "assetname": "Asset name",
                "valid_field_entries": f"Valid entries (of {len(basis_fields)})",
                "validity_rate_pct": st.column_config.NumberColumn(
                    "Field validity", format="%.1f%%"
                ),
            },
        )

    st.subheader("Download results")
    st.caption(
        f"The workbook for {agency_label} contains the agency summary, "
        "validity by field and asset, "
        "check findings, all 33 app-rule counts, validated records, "
        "the original data with field statuses, and a review sheet."
    )
    export_key = (validity_basis, selected_agency)
    if export_key not in st.session_state["validation_exports"]:
        st.session_state["validation_exports"][export_key] = build_export(
            filtered_results, validity_basis
        )
    export_agency = (
        filtered_results[0]["summary"]["agency"]
        if selected_agency == "*" and len(filtered_results) == 1
        else "all_agencies"
        if selected_agency == "*"
        else selected_agency
    )
    export_scope = re.sub(r"[^A-Za-z0-9_-]+", "_", export_agency).strip("_")
    st.download_button(
        "Download validation workbook",
        data=st.session_state["validation_exports"][export_key],
        file_name=f"nars_validation_{export_scope}_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if st.button("Clear results for this session", key="clear_results"):
        st.session_state.pop("validation_results", None)
        st.session_state.pop("validation_errors", None)
        st.session_state.pop("validation_exports", None)
        st.rerun()

st.divider()
