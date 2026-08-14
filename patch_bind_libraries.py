#!/usr/bin/env python3
# Patch: bind LigneeUI and CohortRegistry to tsaraentry, at HEAD.
#
# "At HEAD" means version "0" + developmentMode true — the same pattern
# LigneeUI already uses to bind CohortRegistry, and TSARAENGINE uses too.
# HEAD is required here because lu_submitEvent's new ssId parameter only
# exists at HEAD, not in the pinned v24 that lot files use.
#
# Parsed as JSON rather than text-anchored, because appsscript.json is a
# nested structure and a parse-modify-write round trip cannot silently
# corrupt sibling keys the way a string replace could. Key order is
# preserved (Python dicts and json.load keep insertion order), and the
# two existing libraries (ConsoProvende, AutoCommandes) are left
# byte-for-byte as they were — only two new entries are appended.

import json

PATH = "appsscript.json"

with open(PATH, "r", encoding="utf-8") as f:
    data = json.load(f)

libs = data["dependencies"]["libraries"]
existing_symbols = [l["userSymbol"] for l in libs]

assert len(libs) == 2, "expected exactly 2 existing libraries, found %d" % len(libs)
assert "LigneeUI" not in existing_symbols, "LigneeUI already bound — patch already applied?"
assert "CohortRegistry" not in existing_symbols, "CohortRegistry already bound — patch already applied?"

libs.append({
    "userSymbol": "LigneeUI",
    "libraryId": "1K_h9YyfWHmtkg_ZEdmvzQiS-WQrjtdqb5AQFfSuWy1sV83ybCAJSq0Z5",
    "version": "0",
    "developmentMode": True
})
libs.append({
    "userSymbol": "CohortRegistry",
    "libraryId": "1s2BI_qqCIid4l4RZowX0Lukyr56UYt4uI-Mpzr19ZGfMoFsz5kGBLsoe",
    "version": "0",
    "developmentMode": True
})

assert len(data["dependencies"]["libraries"]) == 4, "expected 4 libraries after patch"

with open(PATH, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2)
    f.write("\n")

print("OK: LigneeUI and CohortRegistry bound to tsaraentry at HEAD (4 libraries total).")
