# NARS validation rule map

**Current app ruleset:** `2026-11-02.1`
**Reference source:** the locally supplied `temp_codes.do` and its eleven
Stata lookup files
**Execution:** Python implementation; Stata is not invoked

## Input schema

The app reads the first worksheet and maps these 35 source columns by heading.
Current files use the column positions below; heading-based mapping allows
future agency workbooks to move those columns without silently validating the
wrong field.

| Stata column | App field | Source heading |
|---:|---|---|
| 4 | `assettype` | Asset Type |
| 6 | `assetname` | Asset Name |
| 7 | `assetdesc` | Asset Description |
| 8 | `region` | Region |
| 9 | `province` | Province |
| 15 | `lat` | Latitude |
| 16 | `lon` | Longitude |
| 17 | `owner` | Owner |
| 18 | `imped` | With Impediments? |
| 19 | `imped_det` | With Impediments? (details) |
| 20 | `acqmode` | Mode of Acquisition / Conveyance |
| 25 | `ownland_str` | FOR STRUCTURE: Do you own the land where the structure is situated/located? |
| 26 | `ownland_lnd` | FOR LAND: Do you own the land? |
| 29–33 | cost, depreciation, net book value, asset life, years of use | Acquisition Cost; Accumulated Depreciation; Net Book Value; Asset Life (years); Number of years of use |
| 35, 38, 41, 44, 47 | `mv_amt`, `imp_amt`, `appr_amt`, `assess_amt`, `repl_amt` | Market Value Amount; Cumulative Amount of Improvement (per books); Appraised Value Amount; Assessed Value Amount; Replacement Value Amount |
| 48 | `disp_mode` | Mode of Disposal |
| 53–54 | `insured`, `policytype` | Is the property insured?; Policy Type |
| 60–61 | `per_from`, `per_to` | Period From; Period To |
| 67–68 | `notins_val`, `assetcond` | Not Insured Value; Asset Condition |
| 71, 76 | `structmat`, `flooddef` | Structural Materials; Flood Defense |
| 78–81 | `area_sqm`, `landclass`, `sec_type` | Floor/Lot Area (in sq. m.); Land Classification; Security Type |

Other source columns are not required or included in the exported validation
records. Entirely blank asset rows are skipped.

## Implemented status fields

Every listed check is emitted as a numeric `*_bad` code and its labelled
`*_status` text. The workbook's `rule_counts` sheet includes zero-count
categories too, so different agency results have the same set of rule labels.

| Area | Validation fields |
|---|---|
| Asset and identity | `asset_type`, `asset_name`, duplicate asset name/description |
| Location | `region`, `province`, `province_region`, `latitude`, `longitude`, `coordinates` |
| Legal and acquisition | `owner`, `impediment`, `impediment_details`, `acquisition_mode` |
| Land and structure | `for_land`, `for_structure`, `land_structure` |
| Financial | `acquisition_cost`, `accumulated_depreciation`, `net_book_value`, `asset_life`, `years_of_use`, `financial_information` |
| Disposal and insurance | `disposal_mode`, `insured`, `policy_type`, `insurance_period`, `insurance_status`, `not_insured_value` |
| Technical | `asset_condition`, `structural_material`, `flood_defense`, `area`, `land_classification`, `security_type` |

## Legacy summary validity rate

The summary workbook's validity measure uses 18 checks per asset, matching the
18 status-count references in the supplied `summary_validation_results.xlsx`.
The app shows valid-entry counts and a validity rate of:

`valid entries / (number of assets × 18)`

The app currently displays the legacy 18-check measure. The 33 app-level
validation rules remain in the detailed rule and finding outputs; they include
overlapping/composite rules and are not 33 distinct source columns.

The accepted statuses for this metric are:

| Legacy check | App rule | Statuses counted as valid |
|---|---|---|
| Asset type | `asset_type` | `valid` |
| Asset name | `asset_name` | `valid` |
| Province + region | `province_region` | `valid` |
| Latitude + longitude | `coordinates` | `valid, within PH` |
| Owner | `owner` | `valid` |
| Impediment details | `impediment_details` | `valid` (all codes with this label) |
| Acquisition mode | `acquisition_mode` | `valid` |
| Acquisition cost | `acquisition_cost` | `valid` |
| Accumulated depreciation | `accumulated_depreciation` | `valid` |
| Net book value | `net_book_value` | `valid` |
| Asset life | `asset_life` | `valid` |
| Years of use | `years_of_use` | `valid` |
| Financial information | `financial_information` | `At least one numeric financial value` |
| Insured | `insured` | `valid` |
| Policy type | `policy_type` | `valid` |
| Insurance status | `insurance_status` | `valid, active insurance`; `land asset: valid entry` |
| Not-insured value | `not_insured_value` | `valid` |
| Asset condition | `asset_condition` | `valid`; `valid, asset is land` |

This deliberately follows the legacy summary's selected categories, which are
narrower than all statuses the app treats as non-error. For example, “No
financial information” is not counted as a valid entry in this rate. **Fully
valid assets** remains a separate measure: an asset must have no findings
across all 33 app rules and no duplicate-name finding. Total findings are the
sum of failing app-rule checks plus one duplicate finding for each affected
asset; they are not the complement of the 18-field valid-entry count.

## Source behavior requiring confirmation

The rules are transcribed from the supplied program, but exact business parity
has not yet been approved:

- The Stata financial-information block tests `fin_flag`, but it is neither
  created in the script nor one of the retained source fields. The app follows
  the block's apparent intent by distinguishing no financial values, at least
  one non-zero numeric value, and nonnumeric-only input.
- The Stata coordinate block assigns code 7 (`interchanged`) immediately
  before assigning code 8 to every pair of out-of-range coordinates. The app
  reports `interchanged` only when swapping the two values places them within
  the Philippine latitude and longitude ranges; other pairs outside both
  ranges remain `Outside PH`.
- The app checks whether a province/region pair exists in the combined
  geographic codebook. The Stata script merges that codebook but assigns the
  combined result from the individual province and region checks instead.
  One DOH record therefore differs between the app and the supplied Stata
  aggregate; confirm whether the combined-pair check is intended.
- Insurance-period parsing accepts the date forms used in the Stata source,
  including numeric hyphen-separated dates and month-name-first dates.
- The expiration check depends on the date of execution. The app makes this an
  explicit, downloadable as-of date rather than an implicit system date.
- Rules and codebooks remain a reviewed translation. Updating the Stata source
  does not automatically alter the Python implementation; update its tests and
  increment the ruleset version.

Validate changes using synthetic test records first and compare the results
against the relevant Stata export before relying on them.
