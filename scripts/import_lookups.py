from __future__ import annotations

import argparse
import json
from pathlib import Path

import pyreadstat


LOOKUP_FILES = {
    "asset_type": ("asset_type_library.dta", "assettype_trim"),
    "region": ("region_library.dta", "region_trim"),
    "province": ("province_library.dta", "province_trim"),
    "province_region": ("province_region_library.dta", "province_region"),
    "acquisition_mode": ("acqmode_library.dta", "acqmode_trim"),
    "disposal_mode": ("dispmode_library.dta", "dispmode_trim"),
    "policy_type": ("poltype_library.dta", "poltype_trim"),
    "structural_material": ("structmat_library.dta", "structmat_trim"),
    "flood_defense": ("flooddef_library.dta", "flooddef_trim"),
    "land_classification": ("landclass_library.dta", "landclass_trim"),
    "security_type": ("sectype_library.dta", "sectype_trim"),
}


def build_lookups(source: Path) -> dict[str, list[str]]:
    lookups: dict[str, list[str]] = {}
    for name, (filename, key_column) in LOOKUP_FILES.items():
        path = source / filename
        if not path.is_file():
            raise FileNotFoundError(f"Required lookup file is missing: {path}")
        frame, _ = pyreadstat.read_dta(str(path))
        if key_column not in frame.columns:
            raise ValueError(f"{filename} does not contain {key_column!r}.")
        values = {
            str(value).strip().lower()
            for value in frame[key_column].dropna().tolist()
            if str(value).strip()
        }
        if not values:
            raise ValueError(f"{filename} contains no lookup values.")
        lookups[name] = sorted(values)
    return lookups


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert local Stata lookup tables to the app's shareable codebooks."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=Path.cwd(),
        help="Folder containing the *_library.dta files (default: current folder).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("config/lookups.json"),
        help="Output JSON path (default: config/lookups.json).",
    )
    args = parser.parse_args()

    lookups = build_lookups(args.source)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(lookups, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {len(lookups)} lookup lists to {args.output}.")


if __name__ == "__main__":
    main()
