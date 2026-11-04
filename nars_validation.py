from __future__ import annotations

import json
import math
import re
import zipfile
from collections import Counter
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException


RULESET_VERSION = "2026-11-02.1"
MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_TOTAL_UNCOMPRESSED_BYTES = 150 * 1024 * 1024
MAX_ROWS = 50_000
MAX_TOTAL_RECORDS = 50_000
MAX_COLUMNS = 256

MISSING_VALUES = {
    "",
    "na",
    "n/a",
    "-",
    "none",
    "no data",
    "no data available",
    "blank",
}

FIELD_HEADERS: dict[str, tuple[str, ...]] = {
    "assettype": ("Asset Type",),
    "assetname": ("Asset Name",),
    "assetdesc": ("Asset Description",),
    "region": ("Region",),
    "province": ("Province",),
    "lat": ("Latitude", "Lat"),
    "lon": ("Longitude", "Long", "Lon"),
    "owner": ("Owner",),
    "imped": ("With Impediments?",),
    "imped_det": ("With Impediments? (details)",),
    "acqmode": ("Mode of Acquisition / Conveyance", "Mode of Acquisition/Conveyance"),
    "ownland_str": (
        "FOR STRUCTURE: Do you own the land where the structure is situated/located?",
        "For Structure: Do you own the land where the structure is situated/located?",
    ),
    "ownland_lnd": ("FOR LAND: Do you own the land?", "For Land: Do you own the land?"),
    "acq_cost": ("Acquisition Cost",),
    "acc_dep": ("Accumulated Depreciation",),
    "nbv": ("Net Book Value",),
    "assetlife": ("Asset Life (years)",),
    "yrs_use": ("Number of years of use", "Years of Use"),
    "mv_amt": ("Market Value Amount",),
    "imp_amt": ("Cumulative Amount of Improvement (per books)",),
    "appr_amt": ("Appraised Value Amount",),
    "assess_amt": ("Assessed Value Amount",),
    "repl_amt": ("Replacement Value Amount",),
    "disp_mode": ("Mode of Disposal",),
    "insured": ("Is the property insured?", "Is the Property Insured?"),
    "policytype": ("Policy Type",),
    "per_from": ("Period From",),
    "per_to": ("Period To",),
    "notins_val": ("Not Insured Value",),
    "assetcond": ("Asset Condition",),
    "structmat": ("Structural Materials",),
    "flooddef": ("Flood Defense", "Flood Defenses"),
    "area_sqm": ("Floor/Lot Area (in sq. m.)", "Floor/Lot Area (in sq.m.)"),
    "landclass": ("Land Classification",),
    "sec_type": ("Security Type",),
}

FIELD_POSITIONS: dict[str, int] = {
    "assettype": 4,
    "assetname": 6,
    "assetdesc": 7,
    "region": 8,
    "province": 9,
    "lat": 15,
    "lon": 16,
    "owner": 17,
    "imped": 18,
    "imped_det": 19,
    "acqmode": 20,
    "ownland_str": 25,
    "ownland_lnd": 26,
    "acq_cost": 29,
    "acc_dep": 30,
    "nbv": 31,
    "assetlife": 32,
    "yrs_use": 33,
    "mv_amt": 35,
    "imp_amt": 38,
    "appr_amt": 41,
    "assess_amt": 44,
    "repl_amt": 47,
    "disp_mode": 48,
    "insured": 53,
    "policytype": 54,
    "per_from": 60,
    "per_to": 61,
    "notins_val": 67,
    "assetcond": 68,
    "structmat": 71,
    "flooddef": 76,
    "area_sqm": 78,
    "landclass": 79,
    "sec_type": 81,
}

RULE_LABELS: dict[str, dict[int, str]] = {
    "asset_type": {
        0: "valid",
        1: "invalid",
        2: "missing",
        3: "with leading space",
        4: "with trailing space",
    },
    "asset_name": {
        0: "valid",
        1: "no asset name",
        2: "no asset description",
        3: "no asset name and description",
    },
    "region": {0: "valid", 1: "invalid", 2: "missing"},
    "province": {0: "valid", 1: "invalid", 2: "missing"},
    "province_region": {
        0: "valid",
        1: "missing province",
        2: "missing region",
        3: "missing province and region",
        4: "invalid province",
        5: "invalid region",
        6: "invalid province and region",
        7: "invalid province and missing region",
        8: "invalid region and missing province",
    },
    "latitude": {0: "valid", 1: "invalid format", 2: "missing", 3: "outside PH"},
    "longitude": {0: "valid", 1: "invalid format", 2: "missing", 3: "outside PH"},
    "coordinates": {
        0: "valid, within PH",
        1: "Lat - invalid format",
        2: "Lat - missing",
        3: "Lat - outside PH",
        4: "Long - invalid format",
        5: "Long - missing",
        6: "Long - outside PH",
        7: "interchanged",
        8: "Outside PH",
        9: "both are invalid",
        10: "both are missing",
        11: "invalid lat, missing long",
        12: "missing lat, invalid long",
    },
    "owner": {0: "valid", 1: "invalid", 2: "missing"},
    "impediment": {
        0: "valid",
        1: "impediment should be yes/no",
        2: "missing impediment (yes/no)",
    },
    "impediment_details": {
        0: "valid",
        1: "impediment should be yes/no",
        2: "missing impediment (yes/no)",
        3: "valid",
        4: "provide impediment details",
        5: "details should be blank",
        6: "valid",
    },
    "acquisition_mode": {
        0: "valid",
        1: "invalid",
        2: "missing",
        3: "with leading space",
        4: "with trailing space",
    },
    "for_land": {
        0: "valid",
        1: "invalid input, 'For Land' should be yes/no",
        2: "missing, fill 'For Land' only",
    },
    "for_structure": {
        0: "valid",
        1: "invalid input, 'For Structure' should be yes/no",
        2: "missing, fill 'For Structure' only",
    },
    "land_structure": {
        0: "valid",
        1: "invalid input, 'For Land' should be yes/no",
        2: "missing, fill 'For Land' only",
        3: "invalid input, fill 'For Land' only",
        4: "invalid input, fill 'For Structure' only",
        5: "invalid input, 'For Structure' should be yes/no",
        6: "missing, fill 'For Structure' only",
        7: "invalid input, missing asset type",
        8: "provide 'For Land' or 'For Structure' information",
    },
    "acquisition_cost": {0: "valid", 1: "invalid", 2: "missing"},
    "accumulated_depreciation": {
        0: "valid",
        1: "land asset: value requires clarification",
        2: "land asset: field should be blank",
        3: "missing value",
        4: "invalid format",
        5: "invalid input: asset type missing",
    },
    "net_book_value": {0: "valid", 1: "invalid", 2: "missing"},
    "asset_life": {
        0: "valid",
        1: "land asset: field should be blank",
        2: "invalid input: should be numeric",
        3: "provide information",
        4: "invalid input: asset type missing",
    },
    "years_of_use": {
        0: "valid",
        1: "invalid input: should be numeric",
        2: "provide information",
    },
    "financial_information": {
        0: "No financial information",
        1: "At least one numeric financial value",
        2: "Only nonnumeric information",
    },
    "disposal_mode": {0: "valid", 1: "not in the list", 2: "no information"},
    "insured": {
        0: "valid",
        1: "land asset: field should be blank or 'no'",
        2: "provide information",
        3: "invalid input: field should be 'yes' or 'no'",
        4: "invalid input: asset type missing",
    },
    "policy_type": {
        0: "valid",
        1: "land asset: field should be blank",
        2: "not in list",
        3: "provide insurance information",
        4: "asset has no insurance: field should be blank",
        5: "invalid input: insurance information missing or incorrect",
        6: "invalid input: asset type missing",
    },
    "insurance_period": {0: "valid", 1: "invalid", 2: "missing"},
    "insurance_status": {
        0: "valid, active insurance",
        1: "outdated insurance",
        2: "missing",
        3: "land asset: field should be blank",
        4: "land asset: valid entry",
        5: "invalid date format",
        6: "provide insurance period",
        7: "not insured: field should be blank",
        8: "not insured: valid entry",
        9: "invalid input: asset type missing",
        10: "invalid input: missing insurance information",
        11: "provide insurance information",
    },
    "not_insured_value": {
        0: "valid",
        1: "land asset: valid entry",
        2: "land asset: field should be blank",
        3: "asset insured: field should be blank",
        4: "provide sum insurable",
        5: "invalid input: field should be numeric",
        6: "no insurance information",
        7: "invalid input: missing asset type",
    },
    "asset_condition": {
        0: "valid",
        1: "valid, asset is land",
        2: "land asset: field should be blank",
        3: "invalid input",
        4: "provide asset condition",
        5: "invalid input: missing asset type",
    },
    "structural_material": {
        0: "valid",
        1: "invalid",
        2: "missing",
        3: "with leading space",
        4: "with trailing space",
    },
    "flood_defense": {
        0: "valid",
        1: "invalid",
        2: "missing",
        3: "with leading space",
        4: "with trailing space",
    },
    "area": {0: "valid", 1: "invalid", 2: "missing"},
    "land_classification": {
        0: "valid",
        1: "invalid",
        2: "provide land classification",
        3: "non-land asset: field should be blank",
        4: "valid, non-land asset",
        5: "asset type missing: invalid input",
        6: "missing asset type",
        7: "with leading space",
        8: "with trailing space",
    },
    "security_type": {
        0: "valid",
        1: "invalid",
        2: "missing",
        3: "with leading space",
        4: "with trailing space",
    },
}

RULE_SOURCE_FIELDS = {
    "asset_type": "assettype",
    "asset_name": "assetname / assetdesc",
    "region": "region",
    "province": "province",
    "province_region": "province / region",
    "latitude": "lat",
    "longitude": "lon",
    "coordinates": "lat / lon",
    "owner": "owner",
    "impediment": "imped",
    "impediment_details": "imped / imped_det",
    "acquisition_mode": "acqmode",
    "for_land": "assettype / ownland_lnd",
    "for_structure": "assettype / ownland_str",
    "land_structure": "assettype / ownland_lnd / ownland_str",
    "acquisition_cost": "acq_cost",
    "accumulated_depreciation": "acc_dep",
    "net_book_value": "nbv",
    "asset_life": "assettype / assetlife",
    "years_of_use": "yrs_use",
    "financial_information": "acq_cost / acc_dep / nbv / mv_amt / imp_amt / appr_amt / assess_amt / repl_amt",
    "disposal_mode": "disp_mode",
    "insured": "assettype / insured",
    "policy_type": "assettype / insured / policytype",
    "insurance_period": "assettype / per_to",
    "insurance_status": "assettype / insured / per_to",
    "not_insured_value": "assettype / insured / notins_val",
    "asset_condition": "assettype / assetcond",
    "structural_material": "structmat",
    "flood_defense": "flooddef",
    "area": "area_sqm",
    "land_classification": "assettype / landclass",
    "security_type": "sec_type",
}

VALID_STATUS_CODES = {
    "financial_information": {0, 1},
    "insurance_status": {0, 4, 8},
    "not_insured_value": {0, 1},
    "asset_condition": {0, 1},
    "land_classification": {0, 4},
}


def _status_label_is_valid(label: str) -> bool:
    normalized_label = label.casefold()
    return "valid" in normalized_label and "invalid" not in normalized_label


def _include_valid_label_codes(rule: str, existing_codes: set[int]) -> frozenset[int]:
    return frozenset(
        existing_codes
        | {
            code
            for code, label in RULE_LABELS[rule].items()
            if _status_label_is_valid(label)
        }
    )


for _rule, _labels in RULE_LABELS.items():
    VALID_STATUS_CODES[_rule] = set(
        _include_valid_label_codes(_rule, VALID_STATUS_CODES.get(_rule, {0}))
    )


LEGACY_VALIDITY_FIELDS = (
    ("asset_type", "Asset type", "asset_type", frozenset({0})),
    ("asset_name", "Asset name", "asset_name", frozenset({0})),
    ("prov_reg", "Province + region", "province_region", frozenset({0})),
    ("lat_lon", "Latitude + longitude", "coordinates", frozenset({0})),
    ("owner", "Owner", "owner", frozenset({0})),
    ("imped_details", "Impediment details", "impediment_details", frozenset({0, 3, 6})),
    ("acqmode", "Acquisition mode", "acquisition_mode", frozenset({0})),
    ("acqcost", "Acquisition cost", "acquisition_cost", frozenset({0})),
    ("accdep", "Accumulated depreciation", "accumulated_depreciation", frozenset({0})),
    ("nbv", "Net book value", "net_book_value", frozenset({0})),
    ("assetlife", "Asset life", "asset_life", frozenset({0})),
    ("yrsuse", "Years of use", "years_of_use", frozenset({0})),
    ("fin_flag", "Financial information", "financial_information", frozenset({1})),
    ("insured", "Insured", "insured", frozenset({0})),
    ("poltype", "Policy type", "policy_type", frozenset({0})),
    ("status_ins", "Insurance status", "insurance_status", frozenset({0, 4})),
    ("notins_val", "Not-insured value", "not_insured_value", frozenset({0})),
    ("assetcond", "Asset condition", "asset_condition", frozenset({0, 1})),
)
LEGACY_VALIDITY_FIELDS = tuple(
    (field, label, rule, _include_valid_label_codes(rule, set(valid_codes)))
    for field, label, rule, valid_codes in LEGACY_VALIDITY_FIELDS
)
_DETAIL_FIELD_LABELS = {
    "province_region": "Province + region",
    "impediment_details": "Impediment details",
    "financial_information": "Financial information",
    "insurance_status": "Insurance status",
    "not_insured_value": "Not-insured value",
    "asset_condition": "Asset condition",
    "structural_material": "Structural material",
    "flood_defense": "Flood defense",
    "land_classification": "Land classification",
    "security_type": "Security type",
}
DETAILED_VALIDITY_FIELDS = tuple(
    (
        rule,
        _DETAIL_FIELD_LABELS.get(rule, rule.replace("_", " ").title()),
        rule,
        frozenset(VALID_STATUS_CODES.get(rule, {0})),
    )
    for rule in RULE_LABELS
)
VALIDITY_FIELDS_BY_BASIS = {
    "legacy18": LEGACY_VALIDITY_FIELDS,
    "detailed33": DETAILED_VALIDITY_FIELDS,
}
LOOKUP_FIELDS = {
    "asset_type": ("assettype", "asset_type"),
    "region": ("region", "region"),
    "province": ("province", "province"),
    "acquisition_mode": ("acqmode", "acquisition_mode"),
    "disposal_mode": ("disp_mode", "disposal_mode"),
    "policy_type": ("policytype", "policy_type"),
    "structural_material": ("structmat", "structural_material"),
    "flood_defense": ("flooddef", "flood_defense"),
    "land_classification": ("landclass", "land_classification"),
    "security_type": ("sec_type", "security_type"),
}
FINANCIAL_FIELDS = (
    "acq_cost",
    "acc_dep",
    "nbv",
    "mv_amt",
    "imp_amt",
    "appr_amt",
    "assess_amt",
    "repl_amt",
)


class ValidationInputError(ValueError):
    """Raised when a workbook cannot be safely mapped to the NARS asset layout."""


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _clean(value: Any) -> str:
    text = _cell_text(value)
    if text.strip(" ").lower() in MISSING_VALUES:
        return ""
    return text


def _trimmed(value: Any) -> str:
    return _clean(value).strip(" ").lower()


def _numeric(value: Any) -> float | None:
    text = _clean(value).replace(",", "")
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return number if math.isfinite(number) else None


def _has_bad_numeric_characters(value: Any, *, allow_comma: bool = True) -> bool:
    text = _clean(value)
    pattern = r"[^0-9.,]" if allow_comma else r"[^0-9.]"
    return text.count(".") > 1 or bool(re.search(pattern, text))


def _date_value(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = _clean(value).strip(" ")
    if not text:
        return None
    formats = (
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%d-%m-%Y",
        "%d-%b-%Y",
        "%d %B %Y",
        "%d-%B-%Y",
        "%d %b %Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%m/%d/%y",
        "%d/%m/%y",
    )
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _status(rule: str, code: int) -> tuple[int, str]:
    return code, RULE_LABELS[rule][code]


def _flag(raw: str, normalized: str, known: set[str], rule: str) -> tuple[int, str]:
    code = 0 if normalized in known else (2 if not raw else 1)
    leading = raw != raw.lstrip(" ")
    trailing = raw != raw.rstrip(" ")
    if code == 0 and leading:
        code = 3
    if code == 0 and trailing:
        code = 4
    return _status(rule, code)


def _validate_record(
    source: dict[str, Any],
    lookups: dict[str, set[str]],
    as_of: date,
) -> dict[str, Any]:
    row = {field: _clean(source.get(field)) for field in FIELD_HEADERS}
    norm = {field: value.strip(" ").lower() for field, value in row.items()}
    result: dict[str, Any] = dict(row)

    def put(rule: str, code: int) -> None:
        status_code, label = _status(rule, code)
        result[f"{rule}_bad"] = status_code
        result[f"{rule}_status"] = label

    # Asset type and identifying description.
    put("asset_type", _flag(row["assettype"], norm["assettype"], lookups["asset_type"], "asset_type")[0])
    missing_name = not row["assetname"]
    missing_description = not row["assetdesc"]
    put(
        "asset_name",
        3 if missing_name and missing_description else 1 if missing_name else 2 if missing_description else 0,
    )
    for rule in ("region", "province"):
        lookup_key = "region" if rule == "region" else "province"
        put(rule, 0 if norm[rule] in lookups[lookup_key] else 2 if not row[rule] else 1)

    province_known = norm["province"] in lookups["province"]
    region_known = norm["region"] in lookups["region"]
    province_missing = not row["province"]
    region_missing = not row["region"]
    pair_known = (
        f"{norm['province']}/{norm['region']}" in lookups["province_region"]
    )
    if province_missing and region_missing:
        pair_code = 3
    elif province_missing:
        pair_code = 1 if region_known else 8
    elif region_missing:
        pair_code = 2 if province_known else 7
    elif not province_known and not region_known:
        pair_code = 6
    elif not province_known:
        pair_code = 4
    elif not region_known:
        pair_code = 5
    else:
        pair_code = 0 if pair_known else 6
    put("province_region", pair_code)

    coordinate_codes: dict[str, int] = {}
    coordinate_values: dict[str, float | None] = {}
    for field, rule, minimum, maximum in (
        ("lat", "latitude", 4.5, 21.5),
        ("lon", "longitude", 116.0, 127.0),
    ):
        value = row[field]
        number = _numeric(value)
        coordinate_values[field] = number
        code = 1 if _has_bad_numeric_characters(value, allow_comma=False) else 0
        if not value:
            code = 2
        elif number is not None and not minimum <= number <= maximum:
            code = 3
        coordinate_codes[field] = code
        put(rule, code)

    lat_code, lon_code = coordinate_codes["lat"], coordinate_codes["lon"]
    combined = lat_code
    if lon_code != 0:
        combined = lon_code + 3
    if lat_code == 1 and lon_code == 1:
        combined = 9
    if lat_code == 2 and lon_code == 2:
        combined = 10
    if lat_code == 1 and lon_code == 2:
        combined = 11
    if lat_code == 2 and lon_code == 1:
        combined = 12
    if lat_code == 3 and lon_code == 3:
        latitude = coordinate_values["lat"]
        longitude = coordinate_values["lon"]
        swapped_coordinates_within_ph = (
            latitude is not None
            and longitude is not None
            and 4.5 <= longitude <= 21.5
            and 116.0 <= latitude <= 127.0
        )
        combined = (
            7
            if swapped_coordinates_within_ph
            else 8
        )
    put("coordinates", combined)

    owner_number = _numeric(row["owner"])
    put("owner", 2 if not row["owner"] else 1 if owner_number is not None else 0)

    imped = norm["imped"]
    imped_code = 0 if imped in {"yes", "no"} else 2 if not row["imped"] else 1
    put("impediment", imped_code)
    has_details = bool(row["imped_det"])
    imped_details = 0
    if imped_code == 1:
        imped_details = 1
    if imped_code == 2:
        imped_details = 2
    if imped == "yes" and not has_details:
        imped_details = 4
    if imped != "yes" and has_details:
        imped_details = 5
    if imped == "no" and not has_details:
        imped_details = 0
    put("impediment_details", imped_details)

    put(
        "acquisition_mode",
        _flag(
            row["acqmode"],
            norm["acqmode"],
            lookups["acquisition_mode"],
            "acquisition_mode",
        )[0],
    )

    land = 1 if norm["assettype"] == "land" else 2 if not row["assettype"] else 0
    for field, rule, applies_to in (
        ("ownland_lnd", "for_land", 1),
        ("ownland_str", "for_structure", 0),
    ):
        answer = norm[field]
        code = 0
        if land == applies_to and answer not in {"yes", "no"}:
            code = 1
        if land == applies_to and not answer:
            code = 2
        put(rule, code)

    land_answer = norm["ownland_lnd"]
    structure_answer = norm["ownland_str"]
    land_code = result["for_land_bad"]
    structure_code = result["for_structure_bad"]
    combined_land_code = 0
    if land == 1 and land_code == 1 and not structure_answer:
        combined_land_code = 1
    if land == 1 and land_code == 2 and not structure_answer:
        combined_land_code = 2
    if land == 1 and structure_answer:
        combined_land_code = 3
    if land == 0 and land_answer:
        combined_land_code = 4
    if land == 0 and structure_code == 1 and not land_answer:
        combined_land_code = 5
    if land == 0 and structure_code == 2 and not land_answer:
        combined_land_code = 6
    if land == 2 and (land_answer or structure_answer):
        combined_land_code = 7
    if (
        (land == 2 or land == 1 or land == 0)
        and not land_answer
        and not structure_answer
    ):
        combined_land_code = 8
    put("land_structure", combined_land_code)

    for field, rule, allow_comma in (
        ("acq_cost", "acquisition_cost", True),
        ("nbv", "net_book_value", True),
    ):
        value = row[field]
        put(
            rule,
            2 if not value else 1 if _has_bad_numeric_characters(value, allow_comma=allow_comma) else 0,
        )

    depreciation = row["acc_dep"]
    depreciation_bad = _has_bad_numeric_characters(depreciation)
    if land == 1:
        depreciation_code = (
            0 if not depreciation else 2 if depreciation_bad else 1
        )
    elif land == 0:
        depreciation_code = 3 if not depreciation else 4 if depreciation_bad else 0
    else:
        depreciation_code = 3 if not depreciation else 5
    put("accumulated_depreciation", depreciation_code)

    asset_life = row["assetlife"]
    asset_life_code = 0
    if land == 1 and asset_life:
        asset_life_code = 1
    if land == 0 and asset_life and _has_bad_numeric_characters(asset_life):
        asset_life_code = 2
    if land in {0, 2} and not asset_life:
        asset_life_code = 3
    if land == 2 and asset_life:
        asset_life_code = 4
    put("asset_life", asset_life_code)

    years = row["yrs_use"]
    put(
        "years_of_use",
        2 if not years else 1 if _has_bad_numeric_characters(years) else 0,
    )

    financial_values = [row[field] for field in FINANCIAL_FIELDS]
    has_numeric_financial = any(
        value
        and not _has_bad_numeric_characters(value)
        and (_numeric(value) or 0) != 0
        for value in financial_values
    )
    has_text_financial = any(value and _numeric(value) is None for value in financial_values)
    financial_code = 1 if has_numeric_financial else 2 if has_text_financial else 0
    put("financial_information", financial_code)

    put(
        "disposal_mode",
        0 if norm["disp_mode"] in lookups["disposal_mode"] else 2 if not row["disp_mode"] else 1,
    )

    insured = norm["insured"]
    insured_code = 0
    if land == 1 and insured not in {"", "no"}:
        insured_code = 1
    if land == 0 and not insured:
        insured_code = 2
    if land == 0 and insured not in {"yes", "no"} and insured:
        insured_code = 3
    if land == 2 and not insured:
        insured_code = 2
    if land == 2 and insured:
        insured_code = 4
    put("insured", insured_code)

    policy = norm["policytype"]
    policy_known = policy in lookups["policy_type"]
    policy_code = 0
    if land == 1 and policy:
        policy_code = 1
    if land == 0 and insured == "yes" and policy and not policy_known:
        policy_code = 2
    if land == 0 and insured == "yes" and not policy:
        policy_code = 3
    if land == 0 and insured == "no" and policy:
        policy_code = 4
    if land == 0 and insured not in {"yes", "no"} and policy:
        policy_code = 5
    if land == 0 and insured not in {"yes", "no"} and not policy:
        policy_code = 3
    if land == 2:
        policy_code = 6
    put("policy_type", policy_code)

    expiry_raw = row["per_to"]
    expiry = _date_value(source.get("per_to"))
    expiry_code = 0
    if land == 2 and expiry is not None:
        expiry_code = 1
    elif land == 2 and expiry is None:
        expiry_code = 2
    elif expiry_raw and expiry is None:
        expiry_code = 1
    elif not expiry_raw and land == 0:
        expiry_code = 2
    put("insurance_period", expiry_code)

    if expiry is not None and expiry < as_of:
        insurance_status = 1
    elif expiry is None:
        insurance_status = 2
    else:
        insurance_status = 0
    if land == 1 and expiry is not None:
        insurance_status = 3
    if land == 1 and expiry is None:
        insurance_status = 4
    if expiry_raw and expiry is None and insured == "yes" and land == 0:
        insurance_status = 5
    if not expiry_raw and insured == "yes" and land in {0, 2}:
        insurance_status = 6
    if expiry_raw and insured == "no" and land == 0:
        insurance_status = 7
    if not expiry_raw and insured == "no" and land == 0:
        insurance_status = 8
    if expiry_raw and land == 2:
        insurance_status = 9
    if not expiry_raw and land == 2:
        insurance_status = 6
    if insured not in {"yes", "no"} and expiry_raw and land == 0:
        insurance_status = 10
    if insured not in {"yes", "no"} and not expiry_raw and land == 0:
        insurance_status = 11
    put("insurance_status", insurance_status)

    not_insured = row["notins_val"]
    not_insured_code = 0
    insured_code_value = 1 if insured == "yes" else 0 if insured == "no" else 2
    if land == 1 and not not_insured:
        not_insured_code = 1
    if land == 1 and not_insured:
        not_insured_code = 2
    if land == 0 and insured_code_value == 1 and not_insured:
        not_insured_code = 3
    if land == 0 and insured_code_value == 0 and not not_insured:
        not_insured_code = 4
    if (
        land == 0
        and insured_code_value == 0
        and not_insured
        and _has_bad_numeric_characters(not_insured, allow_comma=False)
    ):
        not_insured_code = 5
    if land == 0 and insured_code_value == 2:
        not_insured_code = 6
    if land == 2:
        not_insured_code = 7
    put("not_insured_value", not_insured_code)

    condition = norm["assetcond"]
    if land == 1:
        condition_code = 1 if not condition else 2
    elif land == 2:
        condition_code = 4 if not condition else 5
    elif condition not in {"good", "fair", "poor", "bad"}:
        condition_code = 4 if not condition else 3
    else:
        condition_code = 0
    put("asset_condition", condition_code)

    for field, rule, lookup in (
        ("structmat", "structural_material", "structural_material"),
        ("flooddef", "flood_defense", "flood_defense"),
        ("sec_type", "security_type", "security_type"),
    ):
        put(rule, _flag(row[field], norm[field], lookups[lookup], rule)[0])

    area = row["area_sqm"]
    area_code = 2 if not area else 1 if not re.search(r"[0-9]", area) else 0
    put("area", area_code)

    classification = norm["landclass"]
    classification_code = 0
    if land == 1 and classification and classification not in lookups["land_classification"]:
        classification_code = 1
    if land == 1 and not classification:
        classification_code = 2
    if land == 0 and classification:
        classification_code = 3
    if land == 0 and not classification:
        classification_code = 0
    if land == 2 and classification:
        classification_code = 5
    if land == 2 and not classification:
        classification_code = 6
    leading = row["landclass"] != row["landclass"].lstrip(" ")
    trailing = row["landclass"] != row["landclass"].rstrip(" ")
    if leading and classification_code == 0:
        classification_code = 7
    if trailing and classification_code == 0:
        classification_code = 8
    put("land_classification", classification_code)

    name_key = f"{norm['assetname']}/{norm['assetdesc']}"
    result["_duplicate_key"] = name_key
    result["_source_values"] = source
    return result


def load_lookups(path: Path | None = None) -> dict[str, set[str]]:
    lookup_path = path or Path(__file__).parent / "config" / "lookups.json"
    if not lookup_path.is_file():
        raise FileNotFoundError(
            f"Lookup codebooks are missing: {lookup_path}. "
            "Regenerate them from the supplied local Stata lookup files."
        )
    raw = json.loads(lookup_path.read_text(encoding="utf-8"))
    required = {lookup for _, lookup in LOOKUP_FIELDS.values()} | {"province_region"}
    missing = required - raw.keys()
    if missing:
        raise ValidationInputError(
            "Lookup codebooks are incomplete: " + ", ".join(sorted(missing))
        )
    return {
        name: {str(value).strip().lower() for value in values}
        for name, values in raw.items()
    }


def _normalized_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", _cell_text(value).lower())


def duplicate_filenames(filenames: list[str]) -> list[str]:
    counts = Counter(Path(filename).name.casefold() for filename in filenames)
    return sorted(
        {
            Path(filename).name
            for filename in filenames
            if counts[Path(filename).name.casefold()] > 1
        },
        key=lambda filename: (filename.casefold(), filename),
    )


def _field_mapping(headers: list[Any]) -> dict[str, int]:
    by_name: dict[str, list[int]] = {}
    for index, value in enumerate(headers):
        normalized = _normalized_header(value)
        if normalized:
            by_name.setdefault(normalized, []).append(index)

    mapping: dict[str, int] = {}
    duplicated: list[str] = []
    for field, aliases in FIELD_HEADERS.items():
        matching = {
            index
            for alias in aliases
            for index in by_name.get(_normalized_header(alias), [])
        }
        if len(matching) > 1:
            duplicated.append(field)
        elif matching:
            mapping[field] = next(iter(matching))
    if len(mapping) == len(FIELD_HEADERS) and not duplicated:
        return mapping

    # Stata's source keeps fixed Excel columns; only accept that fallback when
    # those columns still have the expected field headings.
    positional: dict[str, int] = {}
    for field, position in FIELD_POSITIONS.items():
        if position > len(headers):
            positional.clear()
            break
        expected = {_normalized_header(alias) for alias in FIELD_HEADERS[field]}
        if _normalized_header(headers[position - 1]) not in expected:
            positional.clear()
            break
        positional[field] = position - 1
    if len(positional) == len(FIELD_HEADERS):
        return positional

    missing = sorted(set(FIELD_HEADERS) - mapping.keys())
    duplicate_note = (
        " Duplicate or ambiguous column headings: " + ", ".join(duplicated) + "."
        if duplicated
        else ""
    )
    raise ValidationInputError(
        "This workbook does not match the NARS asset-register layout. "
        "Missing recognizable columns: "
        + ", ".join(missing)
        + ". Confirm it is an asset-register workbook with the required source headings."
        + duplicate_note
    )


def read_workbook(payload: bytes) -> tuple[str, list[dict[str, Any]]]:
    if not payload:
        raise ValidationInputError("The uploaded workbook is empty.")
    if len(payload) > MAX_FILE_BYTES:
        raise ValidationInputError("Each workbook must be 25 MiB or smaller.")
    try:
        with zipfile.ZipFile(BytesIO(payload)) as archive:
            if sum(item.file_size for item in archive.infolist()) > MAX_TOTAL_UNCOMPRESSED_BYTES:
                raise ValidationInputError(
                    "The expanded workbook is too large to process safely (150 MiB limit)."
                )
    except zipfile.BadZipFile as error:
        raise ValidationInputError("The upload is not a valid .xlsx workbook.") from error

    workbook = load_workbook(
        BytesIO(payload),
        read_only=True,
        data_only=True,
        keep_links=False,
    )
    try:
        if not workbook.worksheets:
            raise ValidationInputError("The workbook has no worksheets.")
        sheet = workbook.worksheets[0]
        if sheet.max_column > MAX_COLUMNS:
            raise ValidationInputError("The worksheet exceeds the 256-column limit.")
        if sheet.max_row > MAX_ROWS + 1:
            raise ValidationInputError("The worksheet exceeds the 50,000-record limit.")

        row_iterator = sheet.iter_rows(values_only=True)
        headers: tuple[Any, ...] | None = None
        header_row = 0
        for index, values in enumerate(row_iterator, start=1):
            if any(value is not None and _cell_text(value).strip() for value in values):
                headers = values
                header_row = index
                break
        if headers is None:
            raise ValidationInputError("The first worksheet is empty.")

        mapping = _field_mapping(list(headers))
        records: list[dict[str, Any]] = []
        for excel_row, values in enumerate(row_iterator, start=header_row + 1):
            if len(records) >= MAX_ROWS:
                raise ValidationInputError("The worksheet exceeds the 50,000-record limit.")
            if all(
                not _clean(values[index] if index < len(values) else None)
                for index in mapping.values()
            ):
                continue
            records.append(
                {
                    field: values[index] if index < len(values) else None
                    for field, index in mapping.items()
                }
                | {
                    "source_row": excel_row,
                    "source_sheet": sheet.title,
                    "_original_headers": tuple(headers),
                    "_original_values": tuple(
                        values[index] if index < len(values) else None
                        for index in range(len(headers))
                    ),
                    "_original_field_indexes": mapping,
                }
            )
        return sheet.title, records
    except InvalidFileException as error:
        raise ValidationInputError("The upload is not a supported Excel workbook.") from error
    finally:
        workbook.close()


def validate_records(
    records: list[dict[str, Any]],
    lookups: dict[str, set[str]],
    agency: str,
    source_file: str,
    as_of: date,
) -> dict[str, Any]:
    validated = [_validate_record(source, lookups, as_of) for source in records]
    duplicate_counts = Counter(
        row["_duplicate_key"]
        for row in validated
        if row["_duplicate_key"]
        and row["asset_name_bad"] == 0
    )
    for index, row in enumerate(validated):
        if "source_row" in records[index]:
            row["source_row"] = records[index]["source_row"]
        if "source_sheet" in records[index]:
            row["source_sheet"] = records[index]["source_sheet"]
        for key in (
            "_original_headers",
            "_original_values",
            "_original_field_indexes",
        ):
            if key in records[index]:
                row[key] = records[index][key]
        key = row.pop("_duplicate_key")
        row.pop("_source_values")
        row["agency"] = agency
        row["source_file"] = source_file
        row["duplicate_name_instances"] = (
            duplicate_counts.get(key, 0) - 1 if row["asset_name_bad"] == 0 else 0
        )
        failing_rules = [
            rule
            for rule in RULE_LABELS
            if row[f"{rule}_bad"] not in VALID_STATUS_CODES.get(rule, {0})
        ]
        row["issue_count"] = len(failing_rules) + int(
            row["duplicate_name_instances"] > 0
        )
        row["issues"] = "; ".join(
            f"{RULE_LABELS[rule][row[f'{rule}_bad']]}"
            for rule in failing_rules
        )
        if row["duplicate_name_instances"] > 0:
            duplicate_message = "duplicate asset name/description"
            row["issues"] = (
                f"{row['issues']}; {duplicate_message}"
                if row["issues"]
                else duplicate_message
            )
        row["result"] = "Valid" if row["issue_count"] == 0 else "Review"
    issue_rows = [row for row in validated if row["issue_count"] > 0]
    rule_counts: list[dict[str, Any]] = []
    for rule, labels in RULE_LABELS.items():
        counts = Counter(row[f"{rule}_status"] for row in validated)
        valid_codes = VALID_STATUS_CODES.get(rule, {0})
        for code, label in labels.items():
            count = counts.get(label, 0)
            rule_counts.append(
                {
                    "agency": agency,
                    "rule": rule.replace("_", " ").title(),
                    "status": label,
                    "count": count,
                    "valid": code in valid_codes,
                }
            )

    validity_by_basis: dict[str, dict[str, Any]] = {}
    field_validity_by_basis: dict[str, list[dict[str, Any]]] = {}
    for basis, fields in VALIDITY_FIELDS_BY_BASIS.items():
        valid_entries_per_record = [
            sum(row[f"{rule}_bad"] in valid_codes for _, _, rule, valid_codes in fields)
            for row in validated
        ]
        check_count = len(fields)
        total_checks = len(validated) * check_count
        valid_entries = sum(valid_entries_per_record)
        validity_by_basis[basis] = {
            "valid_field_entries": valid_entries,
            "field_checks": total_checks,
            "validity_rate_pct": (
                valid_entries / total_checks * 100 if total_checks else 0
            ),
        }
        for row, row_valid_entries in zip(validated, valid_entries_per_record):
            row.setdefault("validity_by_basis", {})[basis] = {
                "valid_field_entries": row_valid_entries,
                "field_checks": check_count,
                "validity_rate_pct": row_valid_entries / check_count * 100,
            }

        field_validity: list[dict[str, Any]] = []
        for _, field_label, rule, valid_codes in fields:
            field_valid = sum(row[f"{rule}_bad"] in valid_codes for row in validated)
            field_validity.append(
                {
                    "agency": agency,
                    "field": field_label,
                    "valid_entries": field_valid,
                    "invalid_entries": len(validated) - field_valid,
                    "total_entries": len(validated),
                    "validity_rate_pct": (
                        field_valid / len(validated) * 100 if validated else 0
                    ),
                }
            )
        field_validity_by_basis[basis] = field_validity

    for row in validated:
        row["valid_field_entries"] = row["validity_by_basis"]["legacy18"]["valid_field_entries"]
        row["field_checks"] = row["validity_by_basis"]["legacy18"]["field_checks"]
        row["validity_rate_pct"] = row["validity_by_basis"]["legacy18"]["validity_rate_pct"]

    review_count = len(issue_rows)
    summary = {
        "agency": agency,
        "source_file": source_file,
        "records": len(validated),
        "valid_records": len(validated) - review_count,
        "valid_field_entries": validity_by_basis["legacy18"]["valid_field_entries"],
        "field_checks": validity_by_basis["legacy18"]["field_checks"],
        "validity_rate_pct": validity_by_basis["legacy18"]["validity_rate_pct"],
        "validity_by_basis": validity_by_basis,
        "records_to_review": review_count,
        "total_findings": sum(row["issue_count"] for row in validated),
        "duplicate_records": sum(
            row["duplicate_name_instances"] > 0 for row in validated
        ),
        "as_of": as_of.isoformat(),
        "ruleset_version": RULESET_VERSION,
    }
    return {
        "summary": summary,
        "validity_by_basis": validity_by_basis,
        "field_validity_by_basis": field_validity_by_basis,
        "field_validity": field_validity_by_basis["legacy18"],
        "rule_counts": rule_counts,
        "records": validated,
        "issues": issue_rows,
    }


def validity_failure_counts(
    results: list[dict[str, Any]],
    validity_basis: str = "legacy18",
) -> list[dict[str, Any]]:
    if validity_basis not in VALIDITY_FIELDS_BY_BASIS:
        raise ValueError(f"Unknown validity basis: {validity_basis}")

    failure_rows: list[dict[str, Any]] = []
    for result in results:
        records = result["records"]
        agency = result["summary"]["agency"]
        for _, field_label, rule, valid_codes in VALIDITY_FIELDS_BY_BASIS[validity_basis]:
            status_counts = Counter(
                row[f"{rule}_status"]
                for row in records
                if row[f"{rule}_bad"] not in valid_codes
            )
            failure_rows.extend(
                {
                    "agency": agency,
                    "rule": field_label,
                    "status": status,
                    "count": count,
                }
                for status, count in status_counts.items()
                if count > 0
            )
    return failure_rows


def analyze_workbook(
    payload: bytes,
    agency: str,
    source_file: str,
    as_of: date,
    lookups: dict[str, set[str]] | None = None,
) -> dict[str, Any]:
    _, records = read_workbook(payload)
    if not records:
        raise ValidationInputError("No non-empty asset records were found.")
    return validate_records(
        records,
        lookups or load_lookups(),
        agency=agency,
        source_file=source_file,
        as_of=as_of,
    )


def _excel_safe(value: Any) -> Any:
    if isinstance(value, str):
        text = ILLEGAL_CHARACTERS_RE.sub("", value)
        if text.lstrip(" \t\r\n")[:1] in {"=", "+", "-", "@"}:
            return "'" + text
        return text
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return value


def _write_sheet(
    workbook: Workbook,
    name: str,
    rows: list[dict[str, Any]],
    columns: list[str] | None = None,
) -> None:
    sheet = workbook.create_sheet(name)
    if columns is None:
        columns = list(rows[0].keys()) if rows else []
    sheet.append(columns)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="173B5E")
        cell.alignment = Alignment(vertical="center")
    for row in rows:
        sheet.append([_excel_safe(row.get(column)) for column in columns])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    if columns:
        for index, column in enumerate(columns, start=1):
            width = min(max(len(str(column)) + 2, 14), 38)
            sheet.column_dimensions[get_column_letter(index)].width = width
            if column.endswith("_pct"):
                for row in range(2, sheet.max_row + 1):
                    sheet.cell(row, index).number_format = '0.0"%"'


def _original_data_export(
    records: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[str]]:
    source_columns: list[tuple[tuple[Any, ...], str, str | None]] = []
    labels: dict[tuple[Any, ...], str] = {}
    values_by_record: list[dict[tuple[Any, ...], Any]] = []
    statuses_by_source_field: dict[str, list[str]] = {
        field: [] for field in FIELD_HEADERS
    }
    for rule in RULE_LABELS:
        source_fields = RULE_SOURCE_FIELDS[rule].split(" / ")
        status_anchor = "region" if rule == "province_region" else source_fields[-1]
        statuses_by_source_field[status_anchor].append(f"{rule}_status")

    for record in records:
        headers = record.get("_original_headers")
        original_values = record.get("_original_values")
        field_indexes = record.get("_original_field_indexes")
        if headers is None or original_values is None or field_indexes is None:
            headers = tuple(FIELD_HEADERS)
            original_values = tuple(record.get(field) for field in FIELD_HEADERS)
            field_indexes = {
                field: index for index, field in enumerate(FIELD_HEADERS)
            }

        fields_by_index = {
            index: field for field, index in field_indexes.items()
        }
        source_values: dict[tuple[Any, ...], Any] = {}
        extra_occurrences: Counter[str] = Counter()
        for index, header in enumerate(headers):
            field = fields_by_index.get(index)
            if field is not None:
                column_key = ("field", field)
                label = _cell_text(header).strip() or FIELD_HEADERS[field][0]
            else:
                label = _cell_text(header).strip() or f"Column {index + 1}"
                normalized = _normalized_header(label) or f"blank{index + 1}"
                occurrence = extra_occurrences[normalized]
                extra_occurrences[normalized] += 1
                column_key = ("extra", normalized, occurrence)
            if column_key not in labels:
                labels[column_key] = label
                source_columns.append((column_key, label, field))
            source_values[column_key] = (
                original_values[index] if index < len(original_values) else None
            )
        values_by_record.append(source_values)

    used_columns: set[str] = set()

    def unique_column(label: str) -> str:
        candidate = label
        suffix = 2
        while candidate in used_columns:
            candidate = f"{label} ({suffix})"
            suffix += 1
        used_columns.add(candidate)
        return candidate

    output_columns = [
        unique_column(column)
        for column in ("agency", "source_file", "source_row", "source_sheet")
    ]
    source_output_columns: dict[tuple[Any, ...], str] = {}
    status_output_columns: dict[str, str] = {}
    for column_key, label, field in source_columns:
        source_output_columns[column_key] = unique_column(label)
        output_columns.append(source_output_columns[column_key])
        if field is not None:
            for status_column in statuses_by_source_field[field]:
                status_output_columns[status_column] = unique_column(status_column)
                output_columns.append(status_output_columns[status_column])

    export_rows: list[dict[str, Any]] = []
    for record, source_values in zip(records, values_by_record):
        export_row = {
            "agency": record.get("agency"),
            "source_file": record.get("source_file"),
            "source_row": record.get("source_row"),
            "source_sheet": record.get("source_sheet"),
        }
        for column_key, value in source_values.items():
            export_row[source_output_columns[column_key]] = value
        for status_column, output_column in status_output_columns.items():
            rule = status_column.removesuffix("_status")
            export_row[output_column] = record[f"{rule}_status"]
        export_rows.append(export_row)
    return export_rows, output_columns


def build_export(results: list[dict[str, Any]], validity_basis: str = "legacy18") -> bytes:
    if not results:
        raise ValueError("There are no validation results to export.")
    if validity_basis not in VALIDITY_FIELDS_BY_BASIS:
        raise ValueError(f"Unknown validity basis: {validity_basis}")

    fields = VALIDITY_FIELDS_BY_BASIS[validity_basis]
    summaries = [
        {
            key: value
            for key, value in result["summary"].items()
            if key != "validity_by_basis"
        }
        | result["validity_by_basis"][validity_basis]
        | {"validity_basis": validity_basis}
        for result in results
    ]
    field_validity = [
        row
        for result in results
        for row in result["field_validity_by_basis"][validity_basis]
    ]
    rule_counts = [
        row for result in results for row in result["rule_counts"]
    ]
    all_records = [row for result in results for row in result["records"]]
    all_issues = [row for result in results for row in result["issues"]]
    records_for_export = [
        row
        | row["validity_by_basis"][validity_basis]
        | {"validity_basis": validity_basis}
        for row in all_records
    ]
    asset_field_validity = [
        {
            "agency": row["agency"],
            "source_file": row["source_file"],
            "source_row": row.get("source_row"),
            "source_sheet": row.get("source_sheet"),
            "assetname": row["assetname"],
            "valid_field_entries": row["validity_by_basis"][validity_basis]["valid_field_entries"],
            "field_checks": row["validity_by_basis"][validity_basis]["field_checks"],
            "validity_rate_pct": row["validity_by_basis"][validity_basis]["validity_rate_pct"],
            "validity_basis": validity_basis,
            **{
                field_label: row[f"{rule}_status"]
                for _, field_label, rule, _ in fields
            },
        }
        for row in all_records
    ]

    workbook = Workbook()
    workbook.remove(workbook.active)
    _write_sheet(workbook, "summary", summaries)
    _write_sheet(workbook, "field_validity", field_validity)
    _write_sheet(workbook, "asset_field_validity", asset_field_validity)
    _write_sheet(workbook, "rule_counts", rule_counts)
    statuses_by_source_field: dict[str, list[str]] = {
        field: [] for field in FIELD_HEADERS
    }
    for rule in RULE_LABELS:
        source_fields = RULE_SOURCE_FIELDS[rule].split(" / ")
        status_anchor = "region" if rule == "province_region" else source_fields[-1]
        statuses_by_source_field[status_anchor].append(f"{rule}_status")
    record_columns = [
        "agency",
        "source_file",
        "source_row",
        "source_sheet",
        "result",
        "issue_count",
        "valid_field_entries",
        "field_checks",
        "validity_rate_pct",
        "validity_basis",
        "issues",
        "duplicate_name_instances",
    ]
    for field in FIELD_HEADERS:
        record_columns.append(field)
        record_columns.extend(statuses_by_source_field[field])
    _write_sheet(workbook, "record_results", records_for_export, record_columns)
    original_data_rows, original_data_columns = _original_data_export(all_records)
    _write_sheet(
        workbook,
        "original_data",
        original_data_rows,
        original_data_columns,
    )
    _write_sheet(
        workbook,
        "records_with_issues",
        [
            row
            | row["validity_by_basis"][validity_basis]
            | {"validity_basis": validity_basis}
            for row in all_issues
        ],
        record_columns,
    )
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()
