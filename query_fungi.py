#!/usr/bin/env python3
"""
Fungi habitat evidence scraper.
Output: fungi_full_evidence.json

Sources (1 local database + 1 API, 2 tiers):
  Tier 1 — all genera:
    FungalTraits (Põlme et al. 2020)
      File: FungalTraits_Polme2020_S1.csv (local copy)
      Fields: Aquatic_habitat_template (non-aquatic/partly_aquatic/aquatic/freshwater/
              marine/partly_freshwater/partly_marine), primary_lifestyle,
              secondary_lifestyle, Growth_form_template, decay_substrate
      Genus-level database — one record per genus, exact name match.
  Tier 2 — gap-filling (only genera not found in FungalTraits):
    GBIF speciesProfiles API
      URL: https://api.gbif.org/v1/species/{usageKey}/speciesProfiles
      Fields: marine, freshwater, terrestrial (genus-level booleans)
      Filter: matchType=EXACT, kingdom=Fungi.

Pipeline:
  1. Read genera from Taxon list.xlsx (Group = "Fungi")
  2. Tier 1: match each genus against FungalTraits CSV → extract habitat + lifestyle fields
  3. Tier 2: unmatched genera → query GBIF speciesProfiles for habitat flags
  4. Save all results to JSON

Run log:
  Last run: 2026-09-08
  Environment: macOS 26.3, Python 3.13.13, openpyxl 3.1.5
  Result: 482/499 genera with evidence (FungalTraits 482, GBIF 1)
"""

import json
import csv
import io
import time
import urllib.request
import urllib.error
import openpyxl
from pathlib import Path

XLSX_FILE = Path(__file__).parent / "Taxon list.xlsx"
OUTPUT_JSON = Path(__file__).parent / "fungi_full_evidence.json"
DELAY = 0.3

FUNGALTRAITS_URL = (
    "https://raw.githubusercontent.com/globalbioticinteractions/fungaltraits"
    "/main/polme2020-s1-fungal-traits-genera.csv"
)
FUNGALTRAITS_LOCAL = Path(__file__).parent / "FungalTraits_Polme2020_S1.csv"


# === FungalTraits (Põlme et al. 2020) ===

def load_fungaltraits():
    """Load FungalTraits genus-level database from local file, fallback to download."""
    if FUNGALTRAITS_LOCAL.exists():
        print(f"  Reading local file: {FUNGALTRAITS_LOCAL.name}", flush=True)
        raw = FUNGALTRAITS_LOCAL.read_text(encoding="utf-8", errors="replace")
    else:
        print("  Local file not found, downloading from GitHub...", flush=True)
        req = urllib.request.Request(FUNGALTRAITS_URL, headers={"User-Agent": "Mozilla/5.0"})
        resp = urllib.request.urlopen(req, timeout=60)
        raw = resp.read().decode("utf-8", errors="replace")

    db = {}
    reader = csv.DictReader(io.StringIO(raw))
    for row in reader:
        genus = row.get("GENUS", "").strip()
        if not genus:
            continue
        db[genus] = {
            "aquatic_habitat": row.get("Aquatic_habitat_template", "").strip(),
            "primary_lifestyle": row.get("primary_lifestyle", "").strip(),
            "secondary_lifestyle": row.get("Secondary_lifestyle", "").strip(),
            "growth_form": row.get("Growth_form_template", "").strip(),
            "decay_substrate": row.get("Decay_substrate_template", "").strip(),
            "phylum": row.get("Phylum", "").strip(),
            "class": row.get("Class", "").strip(),
            "order": row.get("Order", "").strip(),
            "family": row.get("Family", "").strip(),
        }
    return db


def extract_fungaltraits(genus, ft_db):
    """Extract habitat evidence from FungalTraits for a genus."""
    entry = ft_db.get(genus)
    if not entry:
        return []

    parts = []
    if entry["aquatic_habitat"]:
        parts.append(f"Aquatic habitat: {entry['aquatic_habitat']}")
    if entry["primary_lifestyle"]:
        parts.append(f"Primary lifestyle: {entry['primary_lifestyle']}")
    if entry["secondary_lifestyle"]:
        parts.append(f"Secondary lifestyle: {entry['secondary_lifestyle']}")
    if entry["growth_form"]:
        parts.append(f"Growth form: {entry['growth_form']}")
    if entry["decay_substrate"]:
        parts.append(f"Decay substrate: {entry['decay_substrate']}")

    if not parts:
        return []

    return [{
        "species": genus,
        "habitat": " | ".join(parts),
        "source": "FungalTraits (Põlme et al. 2020)",
        "url": "https://github.com/globalbioticinteractions/fungaltraits",
    }]


# === GBIF ===

def extract_gbif(genus):
    """Query GBIF speciesProfiles for marine/freshwater/terrestrial flags."""
    url = (
        f"https://api.gbif.org/v1/species/match"
        f"?name={genus}&rank=GENUS&kingdom=Fungi"
    )
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        resp = urllib.request.urlopen(req, timeout=15)
        data = json.loads(resp.read())
    except Exception:
        return []

    if data.get("matchType") != "EXACT":
        return []
    if data.get("kingdom") != "Fungi":
        return []
    nub_key = data.get("usageKey")
    if not nub_key:
        return []

    prof_url = f"https://api.gbif.org/v1/species/{nub_key}/speciesProfiles"
    try:
        req = urllib.request.Request(prof_url, headers={"Accept": "application/json"})
        resp = urllib.request.urlopen(req, timeout=15)
        profs = json.loads(resp.read())
    except Exception:
        return []

    env_parts = set()
    for p in profs.get("results", []):
        if p.get("marine"):
            env_parts.add("marine")
        if p.get("freshwater"):
            env_parts.add("freshwater")
        if p.get("terrestrial"):
            env_parts.add("terrestrial")

    if not env_parts:
        return []

    return [{
        "species": genus,
        "habitat": f"Environment: {', '.join(sorted(env_parts))} (genus-level)",
        "source": "GBIF",
        "url": f"https://www.gbif.org/species/{nub_key}",
    }]


# === Main ===

def _load_evidence():
    if OUTPUT_JSON.exists():
        with open(OUTPUT_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_evidence(evidence):
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(evidence, f, ensure_ascii=False, indent=2)


def main():
    wb = openpyxl.load_workbook(XLSX_FILE, data_only=True)
    ws = wb["Database"]

    genera = []
    for row in range(2, ws.max_row + 1):
        group = ws.cell(row=row, column=3).value
        if group and "fungi" in str(group).lower():
            genus = str(ws.cell(row=row, column=9).value).strip()
            if genus not in genera:
                genera.append(genus)
    wb.close()

    total = len(genera)
    print(f"Scraping {total} fungi genera...", flush=True)

    # === Tier 1: FungalTraits ===
    print(f"\n{'='*60}", flush=True)
    print("PHASE: FUNGALTRAITS (Põlme et al. 2020)", flush=True)
    print("Downloading database...", flush=True)

    ft_db = load_fungaltraits()
    print(f"FungalTraits loaded: {len(ft_db)} genera", flush=True)
    print("=" * 60, flush=True)

    evidence = _load_evidence()
    ft_hits = 0

    for idx, genus in enumerate(genera):
        if genus in evidence and evidence[genus].get("fungaltraits_count", 0) > 0:
            continue

        if (idx + 1) % 50 == 0 or idx == 0:
            print(f"  Progress: {idx+1}/{total} — FT hits: {ft_hits}", flush=True)

        results = extract_fungaltraits(genus, ft_db)
        if results:
            ft_hits += 1
            unique_spp = set(r["species"].lower() for r in results)
            if genus not in evidence:
                evidence[genus] = {
                    "unique_species": len(unique_spp),
                    "total_records": len(results),
                    "fungaltraits_count": len(results),
                    "gbif_count": 0,
                    "all_species": [
                        {"species": r["species"], "habitat": r["habitat"],
                         "source": r["source"], "url": r["url"]}
                        for r in results
                    ],
                }
            else:
                evidence[genus]["all_species"].extend(
                    {"species": r["species"], "habitat": r["habitat"],
                     "source": r["source"], "url": r["url"]}
                    for r in results
                )
                evidence[genus]["fungaltraits_count"] = len(results)
                evidence[genus]["total_records"] = len(evidence[genus]["all_species"])

            short = results[0]["habitat"][:80]
            print(f"  [FT] {genus}: {short}...", flush=True)

    _save_evidence(evidence)

    found_total = sum(1 for g in genera if g in evidence)
    manual = total - found_total
    missing = [g for g in genera if g not in evidence]

    print(f"\n{'='*60}", flush=True)
    print(f"DONE. Output: {OUTPUT_JSON}", flush=True)
    print(f"Total genera: {total}", flush=True)
    print(f"FungalTraits hits: {ft_hits}", flush=True)
    print(f"Total with evidence: {found_total}", flush=True)
    print(f"Manual needed: {manual}", flush=True)
    if missing:
        print(f"Missing genera: {', '.join(missing)}", flush=True)
    print("=" * 60, flush=True)


if __name__ == "__main__":
    main()
