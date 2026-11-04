from __future__ import annotations

import unittest
from datetime import date
from io import BytesIO

from openpyxl import Workbook, load_workbook

from nars_validation import (
    DETAILED_VALIDITY_FIELDS,
    FINANCIAL_FIELDS,
    FIELD_HEADERS,
    FIELD_POSITIONS,
    LEGACY_VALIDITY_FIELDS,
    RULE_LABELS,
    VALID_STATUS_CODES,
    analyze_workbook,
    build_export,
    duplicate_filenames,
    load_lookups,
    validity_failure_counts,
    validate_records,
)


def valid_record(lookups: dict[str, set[str]]) -> dict[str, object]:
    province, region = sorted(lookups["province_region"])[0].split("/", maxsplit=1)
    return {
        "assettype": "Buildings",
        "assetname": "Sample office",
        "assetdesc": "Synthetic example record",
        "region": region,
        "province": province,
        "lat": "13.2",
        "lon": "122.1",
        "owner": "Republic of the Philippines",
        "imped": "No",
        "imped_det": "",
        "acqmode": sorted(lookups["acquisition_mode"])[0],
        "ownland_str": "Yes",
        "ownland_lnd": "",
        "acq_cost": "1,000",
        "acc_dep": "100",
        "nbv": "900",
        "assetlife": "50",
        "yrs_use": "10",
        "mv_amt": "1200",
        "imp_amt": "",
        "appr_amt": "1100",
        "assess_amt": "900",
        "repl_amt": "1300",
        "disp_mode": sorted(lookups["disposal_mode"])[0],
        "insured": "Yes",
        "policytype": sorted(lookups["policy_type"])[0],
        "per_from": "01/01/2025",
        "per_to": "12/31/2099",
        "notins_val": "",
        "assetcond": "Good",
        "structmat": sorted(lookups["structural_material"])[0],
        "flooddef": sorted(lookups["flood_defense"])[0],
        "area_sqm": "125",
        "landclass": "",
        "sec_type": sorted(lookups["security_type"])[0],
    }


def workbook_bytes(record: dict[str, object]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    headers = [None] * max(FIELD_POSITIONS.values())
    values = [None] * max(FIELD_POSITIONS.values())
    for field, position in FIELD_POSITIONS.items():
        headers[position - 1] = FIELD_HEADERS[field][0]
        values[position - 1] = record[field]
    sheet.append(headers)
    sheet.append(values)
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


class NarsValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.lookups = load_lookups()

    def test_valid_record_passes_all_stata_validation_fields(self) -> None:
        result = validate_records(
            [valid_record(self.lookups)],
            self.lookups,
            agency="SYNTHETIC",
            source_file="SYNTHETIC.xlsx",
            as_of=date(2026, 10, 29),
        )

        self.assertEqual(result["summary"]["records"], 1)
        self.assertEqual(result["summary"]["valid_records"], 1)
        self.assertEqual(result["summary"]["valid_field_entries"], 18)
        self.assertEqual(result["summary"]["field_checks"], 18)
        self.assertEqual(result["summary"]["validity_rate_pct"], 100)
        self.assertEqual(result["summary"]["total_findings"], 0)
        self.assertEqual(result["records"][0]["result"], "Valid")
        self.assertEqual(result["records"][0]["valid_field_entries"], 18)
        self.assertEqual(result["records"][0]["financial_information_bad"], 1)
        self.assertEqual(
            result["records"][0]["validity_by_basis"]["legacy18"][
                "valid_field_entries"
            ],
            18,
        )
        self.assertEqual(
            result["records"][0]["validity_by_basis"]["detailed33"][
                "valid_field_entries"
            ],
            33,
        )
        for rule in RULE_LABELS:
            self.assertIn(f"{rule}_bad", result["records"][0])
        self.assertEqual(len(DETAILED_VALIDITY_FIELDS), 33)
        self.assertEqual(result["summary"]["validity_by_basis"]["detailed33"], {
            "valid_field_entries": 33,
            "field_checks": 33,
            "validity_rate_pct": 100,
        })

    def test_status_labels_containing_valid_are_counted_as_valid_only(self) -> None:
        for rule, labels in RULE_LABELS.items():
            valid_codes = VALID_STATUS_CODES.get(rule, {0})
            for code, label in labels.items():
                with self.subTest(rule=rule, status=label):
                    if "invalid" in label.casefold():
                        self.assertNotIn(code, valid_codes)
                    elif "valid" in label.casefold():
                        self.assertIn(code, valid_codes)

        self.assertIn(3, VALID_STATUS_CODES["impediment_details"])
        self.assertIn(6, VALID_STATUS_CODES["impediment_details"])
        legacy_valid_codes = {
            rule: valid_codes for _, _, rule, valid_codes in LEGACY_VALIDITY_FIELDS
        }
        self.assertIn(3, legacy_valid_codes["impediment_details"])
        self.assertIn(6, legacy_valid_codes["impediment_details"])
        self.assertIn(1, legacy_valid_codes["financial_information"])

    def test_detailed_validity_uses_all_33_rule_statuses(self) -> None:
        record = valid_record(self.lookups)
        record["acqmode"] = "not a listed acquisition mode"
        result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )

        detailed = result["summary"]["validity_by_basis"]["detailed33"]
        self.assertEqual(detailed["valid_field_entries"], 32)
        self.assertEqual(detailed["field_checks"], 33)
        self.assertAlmostEqual(detailed["validity_rate_pct"], 32 / 33 * 100)
        acquisition = next(
            row
            for row in result["field_validity_by_basis"]["detailed33"]
            if row["field"] == "Acquisition Mode"
        )
        self.assertEqual(acquisition["valid_entries"], 0)
        self.assertEqual(acquisition["total_entries"], 1)

    def test_legacy_validity_rate_counts_18_checks_and_each_field(self) -> None:
        valid = valid_record(self.lookups)
        review = dict(valid, acqmode="not a listed acquisition mode")
        result = validate_records(
            [valid, review],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )

        self.assertEqual(len(LEGACY_VALIDITY_FIELDS), 18)
        self.assertEqual(result["summary"]["valid_field_entries"], 35)
        self.assertEqual(result["summary"]["field_checks"], 36)
        self.assertAlmostEqual(result["summary"]["validity_rate_pct"], 35 / 36 * 100)
        self.assertEqual(
            [row["valid_field_entries"] for row in result["records"]],
            [18, 17],
        )
        acquisition = next(
            row for row in result["field_validity"] if row["field"] == "Acquisition mode"
        )
        self.assertEqual(acquisition["valid_entries"], 1)
        self.assertEqual(acquisition["invalid_entries"], 1)
        self.assertEqual(acquisition["total_entries"], 2)
        self.assertEqual(acquisition["validity_rate_pct"], 50)

    def test_findings_table_counts_only_failed_legacy_checks(self) -> None:
        record = valid_record(self.lookups)
        record["acqmode"] = "not a listed acquisition mode"
        record["disp_mode"] = "not a listed disposal mode"
        result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )

        failures = validity_failure_counts([result])
        self.assertEqual(
            failures,
            [
                {
                    "agency": "TEST",
                    "rule": "Acquisition mode",
                    "status": "invalid",
                    "count": 1,
                }
            ],
        )

    def test_legacy_validity_excludes_no_financial_information_status(self) -> None:
        record = valid_record(self.lookups)
        for field in FINANCIAL_FIELDS:
            record[field] = ""
        result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )

        self.assertEqual(result["records"][0]["financial_information_bad"], 0)
        financial = next(
            row
            for row in result["field_validity"]
            if row["field"] == "Financial information"
        )
        self.assertEqual(financial["valid_entries"], 0)
        self.assertEqual(financial["invalid_entries"], 1)

    def test_acquisition_cost_rejects_multiple_decimal_points(self) -> None:
        record = valid_record(self.lookups)
        for field in FINANCIAL_FIELDS:
            record[field] = ""
        record["acq_cost"] = "2,837.856.00"

        invalid_result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )
        invalid_row = invalid_result["records"][0]
        self.assertEqual(invalid_row["acquisition_cost_bad"], 1)
        self.assertEqual(invalid_row["financial_information_bad"], 2)

        record["acq_cost"] = "2,837,856.00"
        valid_result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )
        valid_row = valid_result["records"][0]
        self.assertEqual(valid_row["acquisition_cost_bad"], 0)
        self.assertEqual(valid_row["financial_information_bad"], 1)

    def test_invalid_values_and_normalized_missing_tokens_are_reported(self) -> None:
        record = valid_record(self.lookups)
        record.update(
            assettype="not a listed asset type",
            assetname="",
            province="not a province",
            lat="",
            owner="12345",
            imped="unknown",
        )
        result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )
        row = result["records"][0]

        self.assertEqual(row["asset_type_bad"], 1)
        self.assertEqual(row["asset_name_bad"], 1)
        self.assertEqual(row["province_bad"], 1)
        self.assertEqual(row["latitude_bad"], 2)
        self.assertEqual(row["owner_bad"], 1)
        self.assertEqual(row["impediment_bad"], 1)
        self.assertEqual(row["result"], "Review")

    def test_swapped_coordinates_are_flagged_as_interchanged(self) -> None:
        record = valid_record(self.lookups)
        record.update(lat="122.1", lon="13.2")
        result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )

        self.assertEqual(result["records"][0]["coordinates_bad"], 7)
        self.assertEqual(result["records"][0]["coordinates_status"], "interchanged")

    def test_outside_coordinates_are_not_interchanged_unless_swapping_is_in_ph(self) -> None:
        record = valid_record(self.lookups)
        record.update(lat="80", lon="30")
        result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )

        self.assertEqual(result["records"][0]["latitude_status"], "outside PH")
        self.assertEqual(result["records"][0]["longitude_status"], "outside PH")
        self.assertEqual(result["records"][0]["coordinates_bad"], 8)
        self.assertEqual(result["records"][0]["coordinates_status"], "Outside PH")

    def test_stata_supported_insurance_date_formats_are_parsed(self) -> None:
        for date_value in ("10-25-2024", "25-10-2024", "October 25, 2024"):
            with self.subTest(date_value=date_value):
                record = valid_record(self.lookups)
                record["per_to"] = date_value
                result = validate_records(
                    [record],
                    self.lookups,
                    agency="TEST",
                    source_file="TEST.xlsx",
                    as_of=date(2026, 10, 29),
                )

                self.assertEqual(result["records"][0]["insurance_period_bad"], 0)
                self.assertEqual(
                    result["records"][0]["insurance_status_status"],
                    "outdated insurance",
                )

    def test_duplicate_asset_descriptions_are_review_findings(self) -> None:
        record = valid_record(self.lookups)
        duplicate = dict(record)
        result = validate_records(
            [record, duplicate],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )

        self.assertEqual(result["summary"]["duplicate_records"], 2)
        self.assertEqual(result["summary"]["records_to_review"], 2)
        self.assertEqual(result["summary"]["valid_field_entries"], 36)
        self.assertEqual(result["summary"]["field_checks"], 36)
        self.assertEqual(result["summary"]["validity_rate_pct"], 100)
        self.assertEqual(result["summary"]["valid_records"], 0)
        self.assertEqual(
            [row["duplicate_name_instances"] for row in result["records"]],
            [1, 1],
        )

    def test_upload_maps_selected_columns_for_any_agency_filename(self) -> None:
        result = analyze_workbook(
            workbook_bytes(valid_record(self.lookups)),
            agency="DSWD",
            source_file="DSWD.xlsx",
            as_of=date(2026, 10, 29),
            lookups=self.lookups,
        )

        self.assertEqual(result["summary"]["agency"], "DSWD")
        self.assertEqual(result["summary"]["valid_records"], 1)
        self.assertEqual(result["records"][0]["source_row"], 2)
        self.assertEqual(result["records"][0]["assetname"], "Sample office")

    def test_duplicate_upload_filenames_are_detected_case_insensitively(self) -> None:
        self.assertEqual(
            duplicate_filenames(["DOH.xlsx", "doh.XLSX", "DSWD.xlsx"]),
            ["DOH.xlsx", "doh.XLSX"],
        )
        self.assertEqual(
            duplicate_filenames(["DOH.xlsx", "DSWD.xlsx"]),
            [],
        )

    def test_upload_maps_fields_by_heading_when_columns_are_reordered(self) -> None:
        record = valid_record(self.lookups)
        fields = list(reversed(FIELD_HEADERS))
        workbook = Workbook()
        sheet = workbook.active
        sheet.append([FIELD_HEADERS[field][0] for field in fields])
        sheet.append([record[field] for field in fields])
        stream = BytesIO()
        workbook.save(stream)

        result = analyze_workbook(
            stream.getvalue(),
            agency="DSWD",
            source_file="DSWD.xlsx",
            as_of=date(2026, 10, 29),
            lookups=self.lookups,
        )

        self.assertEqual(result["summary"]["valid_records"], 1)
        self.assertEqual(result["records"][0]["assetname"], "Sample office")

    def test_rejects_workbooks_with_unrecognized_columns(self) -> None:
        workbook = Workbook()
        workbook.active.append(["unrelated", "columns"])
        stream = BytesIO()
        workbook.save(stream)

        with self.assertRaisesRegex(ValueError, "does not match the NARS asset-register layout"):
            analyze_workbook(
                stream.getvalue(),
                agency="UNKNOWN",
                source_file="UNKNOWN.xlsx",
                as_of=date(2026, 10, 29),
                lookups=self.lookups,
            )

    def test_download_workbook_contains_summary_and_record_sheets(self) -> None:
        result = validate_records(
            [valid_record(self.lookups)],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )
        output = load_workbook(BytesIO(build_export([result])), read_only=True)
        try:
            self.assertEqual(
                output.sheetnames,
                [
                    "summary",
                    "field_validity",
                    "asset_field_validity",
                    "rule_counts",
                    "record_results",
                    "original_data",
                    "records_with_issues",
                ],
            )
            summary = list(output["summary"].values)
            self.assertIn("valid_records", summary[0])
            self.assertIn("valid_field_entries", summary[0])
            self.assertEqual(summary[1][0], "TEST")
            self.assertEqual(output["field_validity"].max_row, 19)
            self.assertEqual(output["asset_field_validity"].max_row, 2)
            self.assertIn("Asset type", list(next(output["asset_field_validity"].values)))
            for sheet_name in ("record_results", "records_with_issues"):
                record_headers = list(
                    next(
                        output[sheet_name].iter_rows(
                            min_row=1,
                            max_row=1,
                            values_only=True,
                        )
                    )
                )
                self.assertFalse(
                    any(header.endswith("_bad") for header in record_headers)
                )
                self.assertEqual(
                    record_headers.index("province_status"),
                    record_headers.index("province") + 1,
                )
                region_status_index = record_headers.index("region_status")
                self.assertEqual(
                    record_headers[region_status_index + 1],
                    "province_region_status",
                )
                for rule in RULE_LABELS:
                    self.assertIn(f"{rule}_status", record_headers)
                    self.assertEqual(record_headers.count(f"{rule}_status"), 1)
        finally:
            output.close()

    def test_download_includes_original_input_with_adjacent_status_columns(self) -> None:
        record = valid_record(self.lookups)
        fields = list(reversed(FIELD_HEADERS))
        headers = [FIELD_HEADERS[field][0] for field in fields] + ["Source Note"]
        values = [record[field] for field in fields] + ["Original source value"]
        acquisition_cost_index = fields.index("acq_cost")
        values[acquisition_cost_index] = 1234.5

        workbook = Workbook()
        sheet = workbook.active
        sheet.append(headers)
        sheet.append(values)
        stream = BytesIO()
        workbook.save(stream)

        result = analyze_workbook(
            stream.getvalue(),
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
            lookups=self.lookups,
        )
        output = load_workbook(
            BytesIO(build_export([result])),
            read_only=True,
            data_only=True,
        )
        try:
            original = output["original_data"]
            original_headers = list(
                next(original.iter_rows(min_row=1, max_row=1, values_only=True))
            )
            for rule in RULE_LABELS:
                status_header = f"{rule}_status"
                self.assertEqual(original_headers.count(status_header), 1)
                self.assertEqual(
                    original.cell(2, original_headers.index(status_header) + 1).value,
                    result["records"][0][status_header],
                )
            source_header = FIELD_HEADERS["acq_cost"][0]
            source_index = original_headers.index(source_header)
            self.assertEqual(
                original_headers[source_index + 1],
                "acquisition_cost_status",
            )
            province_header = FIELD_HEADERS["province"][0]
            self.assertEqual(
                original_headers[original_headers.index(province_header) + 1],
                "province_status",
            )
            region_header = FIELD_HEADERS["region"][0]
            region_status_index = original_headers.index(region_header) + 1
            self.assertEqual(original_headers[region_status_index], "region_status")
            self.assertEqual(
                original_headers[region_status_index + 1],
                "province_region_status",
            )
            self.assertEqual(
                original.cell(2, source_index + 1).value,
                1234.5,
            )
            self.assertEqual(
                original.cell(2, original_headers.index("Source Note") + 1).value,
                "Original source value",
            )
            self.assertEqual(
                original.cell(2, original_headers.index("source_row") + 1).value,
                2,
            )
        finally:
            output.close()

    def test_download_workbook_uses_selected_detailed_validity_basis(self) -> None:
        result = validate_records(
            [valid_record(self.lookups)],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )
        output = load_workbook(
            BytesIO(build_export([result], "detailed33")),
            read_only=True,
        )
        try:
            summary = list(output["summary"].values)
            self.assertEqual(summary[1][summary[0].index("field_checks")], 33)
            self.assertEqual(summary[1][summary[0].index("validity_basis")], "detailed33")
            self.assertEqual(output["field_validity"].max_row, 34)
            asset_headers = next(output["asset_field_validity"].values)
            self.assertEqual(len(asset_headers), 42)
            self.assertIn("Disposal Mode", asset_headers)
        finally:
            output.close()

    def test_download_escapes_spreadsheet_formulas_in_uploaded_values(self) -> None:
        record = valid_record(self.lookups)
        record["assetname"] = "=HYPERLINK(\"https://example.invalid\")"
        result = validate_records(
            [record],
            self.lookups,
            agency="TEST",
            source_file="TEST.xlsx",
            as_of=date(2026, 10, 29),
        )
        output = load_workbook(BytesIO(build_export([result])), read_only=True, data_only=False)
        try:
            sheet = output["record_results"]
            headers = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
            cell = sheet.cell(2, headers.index("assetname") + 1)
            self.assertEqual(cell.data_type, "s")
            self.assertTrue(cell.value.startswith("'="))
        finally:
            output.close()


if __name__ == "__main__":
    unittest.main()
