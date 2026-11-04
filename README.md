# NARS Data Validation Tool

A local-first, Stata-free web app for checking NARS workbooks.
Upload one or more agency `.xlsx` files, review record-level findings and
download an Excel workbook containing the validation summary and results.


## Run locally on Windows

Install Python 3.10 or later, then run these commands from this folder:

```powershell
.\setup-local.ps1
.\run-local.ps1
```

Open [http://127.0.0.1:8501](http://127.0.0.1:8501). The app binds to loopback
only. Press `Ctrl+C` in the terminal to stop it.

You can also run the setup and app explicitly:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\streamlit.exe run app.py --server.address 127.0.0.1 --server.port 8501
```

No Stata installation, database, login, API, or internet connection is
required after dependencies have been installed.

## Validate agency workbooks

1. Select one or more `.xlsx` NARS workbooks. Use an agency code
   in each filename, for example `DOH.xlsx`, `DSWD.xlsx`, or `DA_PFDA.xlsx`;
   the filename is shown as the agency/dataset name. There is no fixed
   workbook-count limit; each workbook must be 25 MiB or smaller, the total
   upload must be 100 MiB or smaller, and a validation batch is limited to
   50,000 records. Duplicate filenames (case-insensitive) are flagged and
   must be removed before validation.
2. The app maps the 35 fields used by the source validation script using their
   column headings. Their order may differ between workbooks. The first
   worksheet is used, matching the source script's first-sheet import.
3. Use the gear button to view the insurance reference date. Select **Change
   reference date** to open the calendar and update it, then choose an **as-of
   date** for the insurance expiry check. The same date can be selected in a
   later run to reproduce its result.
4. Run validation and use **Filter by agency** to view all agencies or one
   agency at a time. The selection applies to the summary, metrics, findings,
   record tables, and downloaded workbook. Review the validity rate and
   per-field/per-asset status tables, along with record-level findings. Sort
   table columns using each column's menu.
   The workbook includes `summary`, `field_validity`, `asset_field_validity`,
   `rule_counts`, `record_results`, `original_data`, and `records_with_issues`.
   The record sheets omit numeric `_bad` codes and place each status column
   beside its source field. `original_data` retains the uploaded column names,
   source values, and other columns, with each applicable status beside its
   original field.

The validity rate matches the existing summary workbook:
`valid field entries / (assets checked × 18)`. The separate **Fully valid
assets** count includes only assets with no findings under all 33 app rules
and no duplicate-name finding.

The app validates one or more agencies per run. Each workbook must still follow the NARS data layout represented
by the supplied agency files. An unrelated or incompatible workbook is rejected
with a list of the fields that could not be mapped.

Uploads are held in memory while the local Streamlit process/session is active.
The app does not save originals or results to disk; the downloaded workbook is
saved only when you choose to download it. Clearing results removes the
displayed results from the current Streamlit session.

## Validation rules and Stata updates

`nars_validation.py` implements the 33 validation status fields from the
supplied `temp_codes.do`, including the asset, location, legal, financial,
insurance, and technical checks. `config/lookups.json` contains the reference
codebooks converted from the supplied Stata lookup tables. It contains allowed
categories and geographic names, not agency asset records.
Status labels containing “valid” are counted as valid, case-insensitively,
unless they contain “invalid”; other rule-specific valid statuses remain
explicitly supported.

Stata is **not executed**. To incorporate a later change to the Stata rules:

1. Update `nars_validation.py` to reflect the Stata change and update
   `RULESET_VERSION`.
2. Add or adjust synthetic examples in `tests/test_nars_validation.py`.
3. Run `python -m unittest discover -s tests -v`.
4. If a lookup list changes, replace the local lookup `.dta` file and run:

   ```powershell
   .\.venv\Scripts\python.exe scripts\import_lookups.py --source .
   ```

5. Compare a local validation run with the corresponding Stata output before
   using the changed rules for operational reporting.

The legacy `.do`, `.dta`, and source `.xlsx` files stay local and are ignored by
Git. The app and its tests are the shareable, maintained implementation. When
sharing rule changes with other developers, include the updated Python rules,
tests, and lookup JSON, but do not add raw agency records or exports.

**Parity review required:** The source script references `fin_flag` without
creating it in the selected input fields. The app interprets that check's
apparent intent: distinguish no financial values, at least one numeric value,
and text-only values. The Stata coordinate logic assigns the “interchanged”
status and then overwrites it for all out-of-range pairs; the app preserves the
intended “interchanged” classification. These two rules should be confirmed
against the rule owner before accepting formal output parity. An interchanged
coordinate pair is reported only when swapping the values places both within
the Philippines latitude and longitude ranges. The app also rejects numeric
values containing more than one decimal point; the supplied Stata amount
checks only reject nonnumeric characters, so these stricter checks should be
mirrored in Stata if adopted as official rules.

The app also validates province/region pairs against the combined geographic
codebook. The supplied Stata script merges that codebook but bases its reported
status only on the individual province and region checks; this causes one
aggregate-category difference in the supplied DOH comparison. Confirm whether
the combined-pair check is intended.

The current result is a first local implementation of the supplied rule set;
have the business owner verify all 33 rules and compare results across agencies
before using it as an official record or deploying it.

## GitLab and safe sharing

The repository is for application code, tests, documentation and the
non-record codebooks. `.gitignore` excludes Excel/CSV exports, Stata `.dta`
data, Stata `.do` source files, local training files, secrets, logs, local
Python environments and editor state.

Clone the project from the internal GitLab server:

```powershell
git clone git@172.20.25.132:aaravana/nars-data-validation.git
cd nars-data-validation
```

Before committing, review `git status --short --ignored` and ensure the staged
files contain application code only—not uploaded workbooks, validation exports,
Stata files, local training data, or credentials. Use the organization's
approved repository access controls; a private Git repository is not a storage
location for confidential asset-register records.

## Tests

```powershell
python -m unittest discover -s tests -v
```

Tests create synthetic workbooks in memory and do not require or include the
provided agency data.
