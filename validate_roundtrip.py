#!/usr/bin/env python3
"""
validate_roundtrip.py — Part 0 validation.

Deep-compares the original COUNTRY_DATA (extracted from the HTML) with the
`countries` object produced by converter.py after the data has been pushed
through the sheet layout:

    COUNTRY_DATA -> exporter -> sheets -> converter -> data.json
    data.json["countries"]  should deep-equal  COUNTRY_DATA

Numeric comparison uses == semantics (47 == 47.0), matching JavaScript,
since JSON does not distinguish int from float.

Usage:
    python validate_roundtrip.py country_data.json data.json
"""

import json
import sys


def diff(a, b, path="$", out=None):
    if out is None:
        out = []
    if isinstance(a, dict) and isinstance(b, dict):
        for k in a.keys() | b.keys():
            if k not in a:
                out.append(f"{path}.{k}: only in round-trip output")
            elif k not in b:
                out.append(f"{path}.{k}: only in original")
            else:
                diff(a[k], b[k], f"{path}.{k}", out)
    elif isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            out.append(f"{path}: length {len(a)} vs {len(b)}")
        for i, (x, y) in enumerate(zip(a, b)):
            diff(x, y, f"{path}[{i}]", out)
    else:
        if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
            if a == b:
                return out
        if a is None and b is None:
            return out
        if a != b or type(a) is not type(b):
            out.append(f"{path}: {a!r} ({type(a).__name__}) vs {b!r} ({type(b).__name__})")
    return out


def main():
    original_path, output_path = sys.argv[1], sys.argv[2]
    with open(original_path, encoding="utf-8") as f:
        original = json.load(f)
    with open(output_path, encoding="utf-8") as f:
        produced = json.load(f)
    countries = produced.get("countries", produced)

    mismatches = diff(original, countries)
    if mismatches:
        print(f"ROUND-TRIP FAILED: {len(mismatches)} mismatch(es)")
        for m in mismatches[:50]:
            print("  ", m)
        sys.exit(1)
    print("ROUND-TRIP OK: converter output deep-equals the original COUNTRY_DATA")


if __name__ == "__main__":
    main()
