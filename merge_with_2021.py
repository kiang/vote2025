#!/usr/bin/env python3
"""
Merge 2025 cunli-level results with 2021 (case 17) data by VILLCODE.

Inputs:
- 2025 JSON produced by excel_to_cunli_converter.py (default: docs/referendum_cunli_data.json)
- 2021 data at /home/kiang/public_html/referendums2021/docs/data.json (override via --data2021)

Output:
- docs/referendum_cunli_merged_2025_2021.csv (tabular summary only)
"""

import argparse
import csv
import json
from pathlib import Path
from typing import Dict


def load_2021_lookup(path: Path) -> Dict[str, Dict[str, int]]:
    """Load 2021 referendum (case 17) data and map VILLCODE -> agree/disagree."""
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    lookup: Dict[str, Dict[str, int]] = {}
    for villcode, village_data in data.items():
        if (
            isinstance(village_data, dict)
            and "17_agree" in village_data
            and "17_disagree" in village_data
        ):
            lookup[str(villcode)] = {
                "agree_2021": int(village_data["17_agree"]),
                "disagree_2021": int(village_data["17_disagree"]),
            }
    return lookup


def pct_change(new: int, old: int) -> float:
    if old == 0:
        return 0.0
    return (new - old) / old * 100.0


def merge(input_2025: Path, input_2021: Path, out_csv: Path, unmatched_path: Path) -> None:
    with input_2025.open("r", encoding="utf-8") as f:
        data_2025 = json.load(f)

    villages_2025 = data_2025.get("villages", [])
    lookup_2021 = load_2021_lookup(input_2021)

    merged_villages = []
    unmatched_keys = set()  # key by county|district|village
    unmatched_new_entries = []
    for v in villages_2025:
        villcode = v.get("villcode")
        agree_2025 = int(v.get("total_votes", {}).get("agree", 0))
        disagree_2025 = int(v.get("total_votes", {}).get("disagree", 0))

        agree_2021 = None
        disagree_2021 = None
        agree_diff = None
        disagree_diff = None
        agree_pct = None
        disagree_pct = None

        if villcode and str(villcode) in lookup_2021:
            agree_2021 = int(lookup_2021[str(villcode)]["agree_2021"])
            disagree_2021 = int(lookup_2021[str(villcode)]["disagree_2021"])
            agree_diff = agree_2025 - agree_2021
            disagree_diff = disagree_2025 - disagree_2021
            agree_pct = pct_change(agree_2025, agree_2021)
            disagree_pct = pct_change(disagree_2025, disagree_2021)
        else:
            # Record unmatched for manual mapping update
            key = f"{v.get('county','')}|{v.get('district','')}|{v.get('village','')}"
            if key not in unmatched_keys:
                unmatched_keys.add(key)
                unmatched_new_entries.append(
                    {
                        "county": v.get("county", ""),
                        "district": v.get("district", ""),
                        "village": v.get("village", ""),
                        "villcode": villcode,
                        "suggested_villcode": "",
                        "notes": "Missing in 2021 dataset (case 17)"
                    }
                )

        merged_villages.append(
            {
                "villcode": villcode,
                "county": v.get("county", ""),
                "district": v.get("district", ""),
                "village": v.get("village", ""),
                # 2025 results
                "agree_2025": agree_2025,
                "disagree_2025": disagree_2025,
                # 2021 (case 17) results using requested field names
                "2021_agree": agree_2021,
                "2021_disagree": disagree_2021,
                # Differences and percentage changes (2025 vs 2021; base = 2021)
                "diff": (
                    {"agree": agree_diff, "disagree": disagree_diff}
                    if agree_diff is not None
                    else None
                ),
                "pct_change": (
                    {"agree": agree_pct, "disagree": disagree_pct}
                    if agree_pct is not None
                    else None
                ),
            }
        )

    # Write a concise CSV summary only
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "villcode",
                "county",
                "district",
                "village",
                "agree_2025",
                "disagree_2025",
                "2021_agree",
                "2021_disagree",
                "agree_diff",
                "disagree_diff",
                "agree_pct_change",
                "disagree_pct_change",
            ]
        )
        for r in merged_villages:
            writer.writerow(
                [
                    r.get("villcode"),
                    r.get("county"),
                    r.get("district"),
                    r.get("village"),
                    r.get("agree_2025"),
                    r.get("disagree_2025"),
                    r.get("2021_agree"),
                    r.get("2021_disagree"),
                    None if r.get("diff") is None else r["diff"].get("agree"),
                    None if r.get("diff") is None else r["diff"].get("disagree"),
                    None
                    if r.get("pct_change") is None
                    else round(float(r["pct_change"].get("agree")), 3),
                    None
                    if r.get("pct_change") is None
                    else round(float(r["pct_change"].get("disagree")), 3),
                ]
            )

    # Merge unmatched entries into existing unmatched_for_mapping.json (append-only)
    try:
        if unmatched_path.exists():
            with unmatched_path.open("r", encoding="utf-8") as f:
                existing = json.load(f)
                if not isinstance(existing, list):
                    existing = []
        else:
            existing = []

        # Build index to avoid duplicates by county|district|village
        def key_of(item: Dict) -> str:
            return f"{item.get('county','')}|{item.get('district','')}|{item.get('village','')}"

        seen = {key_of(item) for item in existing}
        to_add = [item for item in unmatched_new_entries if key_of(item) not in seen]
        if to_add:
            combined = existing + to_add
            with unmatched_path.open("w", encoding="utf-8") as f:
                json.dump(combined, f, ensure_ascii=False, indent=2)
    except Exception as e:
        # Do not fail the merge if unmatched file cannot be updated
        print(f"Warning: could not update {unmatched_path}: {e}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Merge 2025 cunli results with 2021 data")
    parser.add_argument(
        "--input",
        default="docs/referendum_cunli_data.json",
        help="Path to 2025 cunli JSON",
    )
    parser.add_argument(
        "--data2021",
        default="/home/kiang/public_html/referendums2021/docs/data.json",
        help="Path to 2021 data.json",
    )
    parser.add_argument(
        "--out-csv",
        default="docs/referendum_cunli_merged_2025_2021.csv",
        help="Output CSV path",
    )
    parser.add_argument(
        "--unmatched",
        default="unmatched_for_mapping.json",
        help="Path to unmatched mapping file to append to",
    )
    args = parser.parse_args()

    input_2025 = Path(args.input)
    input_2021 = Path(args.data2021)
    out_csv = Path(args.out_csv)
    unmatched_path = Path(args.unmatched)

    merge(input_2025, input_2021, out_csv, unmatched_path)
    print(f"Merged CSV:  {out_csv}")


if __name__ == "__main__":
    main()
