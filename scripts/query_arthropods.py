#!/usr/bin/env python3
"""
Arthropod habitat evidence scraper.
Output: arthropods_full_evidence.json

Sources (3 APIs, 2 tiers):
  Tier 1 — broad pass over all genera (WoRMS + GBIF run together):
    WoRMS REST API
      URL: https://www.marinespecies.org/rest/AphiaRecordsByName/{genus}
      Fields: isMarine, isBrackish, isFreshwater, isTerrestrial (species-level)
      Filter: phylum=Arthropoda. Max 10 species per genus.
    GBIF speciesProfiles API
      URL: https://api.gbif.org/v1/species/{usageKey}/speciesProfiles
      Fields: marine, freshwater, terrestrial (genus-level booleans)
      Filter: matchType=EXACT, phylum=Arthropoda.
  Tier 2 — gap-filling (only genera with <3 unique species after Tier 1):
    NatureServe Explorer API
      URL: https://explorer.natureserve.org/api/data/speciesSearch
      Fields: Habitat Type (Terrestrial/Aquatic enum)
      Filter: Canada only, informalTaxonomy=Animals, phylum=Arthropoda.

Pipeline:
  1. Read genera from Taxon list.xlsx (Group = "Arthropods")
  2. Tier 1: query WoRMS (genus → AphiaID → species flags) and GBIF (genus → speciesProfiles)
     for all genera; merge results, keeping records from both sources
  3. Tier 2: for genera with <3 unique species, query NatureServe Canada for habitat type
  4. Save all results to JSON 

Usage:
  python query_arthropods.py           # run all tiers in order
  python query_arthropods.py worms     # run one phase only

Run log:
  Last run: 2026-09-08
  Environment: macOS 26.3, Python 3.13.13, openpyxl 3.1.5
  Result: 554/981 genera with evidence (WoRMS 334, GBIF 421, NatureServe 120)
"""

import json
import time
import urllib.request
import urllib.error
import openpyxl
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
XLSX_FILE = PROJECT_ROOT / "data" / "Taxon list.xlsx"
OUTPUT_JSON = PROJECT_ROOT / "data" / "evidence" / "arthropods_full_evidence.json"
DELAY = 0.5
MAX_SPECIES = 10

WORMS_BASE = "https://www.marinespecies.org/rest"


# ================================================================
# WoRMS
# ================================================================

def env_string(record):
    parts = []
    if record.get("isMarine"):
        parts.append("marine")
    if record.get("isBrackish"):
        parts.append("brackish")
    if record.get("isFreshwater"):
        parts.append("freshwater")
    if record.get("isTerrestrial"):
        parts.append("terrestrial")
    return ", ".join(parts) if parts else ""


def worms_search_genus(genus):
    url = f"{WORMS_BASE}/AphiaRecordsByName/{genus}?like=false&marine_only=false"
    try:
        req = urllib.request.Request(url, headers={
            "Accept": "application/json", "User-Agent": "Mozilla/5.0"
        })
        resp = urllib.request.urlopen(req, timeout=15)
        data = json.loads(resp.read())
    except Exception:
        return None, ""
    if not data:
        return None, ""
    for r in data:
        if (r.get("rank") == "Genus" and r.get("status") == "accepted"
                and r.get("phylum") == "Arthropoda"):
            return r["AphiaID"], env_string(r)
    for r in data:
        if r.get("phylum") == "Arthropoda":
            return r.get("AphiaID"), env_string(r)
    return None, ""


def worms_get_species(aphia_id, genus):
    results = []
    offset = 1
    while True:
        url = (
            f"{WORMS_BASE}/AphiaChildrenByAphiaID/{aphia_id}"
            f"?marine_only=false&offset={offset}"
        )
        try:
            req = urllib.request.Request(url, headers={
                "Accept": "application/json", "User-Agent": "Mozilla/5.0"
            })
            resp = urllib.request.urlopen(req, timeout=15)
            children = json.loads(resp.read())
        except urllib.error.HTTPError as e:
            if e.code == 204:
                break
            break
        except Exception:
            break
        if not children:
            break
        for c in children:
            if c.get("rank") != "Species":
                continue
            env = env_string(c)
            if not env:
                continue
            sci_name = c.get("scientificname", "")
            if not sci_name.startswith(genus):
                continue
            worms_url = c.get("url", f"https://www.marinespecies.org/aphia.php?p=taxdetails&id={c['AphiaID']}")
            results.append({
                "species": sci_name,
                "habitat": f"Environment: {env}",
                "source": "WoRMS",
                "url": worms_url,
            })
            if len(results) >= MAX_SPECIES:
                return results
        if len(children) < 50:
            break
        offset += 50
        time.sleep(DELAY)
    return results


def extract_worms(genus):
    aphia_id, genus_env = worms_search_genus(genus)
    time.sleep(DELAY)
    if not aphia_id:
        return []
    species = worms_get_species(aphia_id, genus)
    time.sleep(DELAY)
    if genus_env and not species:
        return [{
            "species": genus,
            "habitat": f"Environment: {genus_env} (genus-level)",
            "source": "WoRMS",
            "url": f"https://www.marinespecies.org/aphia.php?p=taxdetails&id={aphia_id}",
        }]
    return species


# ================================================================
# NatureServe
# ================================================================

def extract_natureserve(genus):
    search_url = "https://explorer.natureserve.org/api/data/speciesSearch"
    payload = json.dumps({
        "criteriaType": "species",
        "textCriteria": [{"paramType": "quickSearch", "searchToken": genus}],
    }).encode("utf-8")

    try:
        req = urllib.request.Request(
            search_url,
            data=payload,
            headers={"Content-Type": "application/json", "Accept": "application/json"},
        )
        resp = urllib.request.urlopen(req, timeout=20)
        data = json.loads(resp.read())
    except Exception:
        return []

    search_results = data.get("results", [])
    results = []

    for r in search_results:
        sci_name = r.get("scientificName", "")
        if not (sci_name == genus or sci_name.startswith(genus + " ")):
            continue
        sg = r.get("speciesGlobal", {})
        if sg.get("phylum") != "Arthropoda":
            continue
        uid = r.get("uniqueId")
        if not uid:
            continue

        detail_url = f"https://explorer.natureserve.org/api/data/taxon/{uid}"
        try:
            req = urllib.request.Request(
                detail_url, headers={"Accept": "application/json"}
            )
            resp = urllib.request.urlopen(req, timeout=15)
            detail = json.loads(resp.read())
        except Exception:
            continue

        parts = []

        ac = detail.get("animalCharacteristics", {})
        major_hab = ac.get("majorHabitat")
        if isinstance(major_hab, dict):
            major_hab = major_hab.get("majorHabitatDescEn", "")
        if major_hab:
            parts.append(f"Habitat Type: {major_hab}")

        sc = detail.get("speciesCharacteristics", {})
        for field, label in [
            ("speciesTerrestrialHabitats", "Terrestrial"),
            ("speciesRiverineHabitats", "Riverine"),
            ("speciesPalustrineHabitats", "Palustrine"),
            ("speciesLacustrineHabitats", "Lacustrine"),
            ("speciesEstuarineHabitats", "Estuarine"),
            ("speciesMarineHabitats", "Marine"),
            ("speciesSubterraneanHabitats", "Subterranean"),
        ]:
            hab_list = sc.get(field, [])
            if hab_list:
                names = []
                for h in hab_list:
                    inner = list(h.values())
                    for v in inner:
                        if isinstance(v, dict):
                            desc = [val for key, val in v.items() if "DescEn" in key]
                            if desc:
                                names.append(desc[0])
                if names:
                    parts.append(f"{label}: {', '.join(names)}")

        hab_comments = sc.get("habitatComments", "")
        if hab_comments:
            parts.append(f"Comments: {hab_comments[:200]}")

        if not parts:
            eco = detail.get("ecologicalInformation", {})
            eco_hab = eco.get("majorHabitat")
            eco_comments = eco.get("habitatComments", "")
            if eco_hab:
                parts.append(f"majorHabitat: {eco_hab}")
            if eco_comments:
                parts.append(f"Comments: {eco_comments[:200]}")

        if parts:
            page_url = (
                f"https://explorer.natureserve.org/Taxon/{uid}"
                f"/{sci_name.replace(' ', '_')}"
            )
            results.append({
                "species": sci_name,
                "habitat": " | ".join(parts),
                "source": "NatureServe Explorer",
                "url": page_url,
            })

        time.sleep(0.3)
        if len(results) >= MAX_SPECIES:
            break

    return results


# ================================================================
# GBIF
# ================================================================

def extract_gbif(genus):
    url = (
        f"https://api.gbif.org/v1/species/match"
        f"?name={genus}&rank=GENUS&kingdom=Animalia&phylum=Arthropoda"
    )
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        resp = urllib.request.urlopen(req, timeout=15)
        data = json.loads(resp.read())
    except Exception:
        return []

    if data.get("matchType") != "EXACT":
        return []
    if data.get("phylum") != "Arthropoda":
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


# ================================================================
# Main
# ================================================================

def _load_evidence():
    if OUTPUT_JSON.exists():
        with open(OUTPUT_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_evidence(evidence):
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(evidence, f, ensure_ascii=False, indent=2)


def _add_records(evidence, genus, new_records, source_key):
    if genus in evidence:
        existing = evidence[genus]
        existing["all_species"].extend(
            {"species": r["species"], "habitat": r["habitat"],
             "source": r["source"], "url": r["url"]}
            for r in new_records
        )
        existing[f"{source_key}_count"] = len(new_records)
        existing["total_records"] = len(existing["all_species"])
        unique = set(r["species"].lower() for r in existing["all_species"])
        existing["unique_species"] = len(unique)
    else:
        unique_spp = set(r["species"].lower() for r in new_records)
        evidence[genus] = {
            "unique_species": len(unique_spp),
            "total_records": len(new_records),
            "worms_count": 0,
            "natureserve_count": 0,
            "gbif_count": 0,
            "all_species": [
                {"species": r["species"], "habitat": r["habitat"],
                 "source": r["source"], "url": r["url"]}
                for r in new_records
            ],
        }
        evidence[genus][f"{source_key}_count"] = len(new_records)


def _get_genera():
    wb = openpyxl.load_workbook(XLSX_FILE, data_only=True)
    ws = wb["Database"]
    genera = []
    for row in range(2, ws.max_row + 1):
        group = ws.cell(row=row, column=3).value
        if group and "arthropod" in str(group).lower():
            genus = str(ws.cell(row=row, column=9).value).strip()
            if genus not in genera:
                genera.append(genus)
    wb.close()
    return genera


TIER1_SOURCES = ["worms", "gbif"]
TIER2_SOURCES = ["natureserve"]


def run_phase(phase, only_missing=False):
    genera = _get_genera()
    evidence = _load_evidence()
    total = len(genera)

    source_count_key = f"{phase}_count"
    already_queried = [g for g in genera if
        g in evidence and evidence[g].get(source_count_key, 0) > 0]

    if only_missing:
        candidates = [g for g in genera if g not in evidence]
        candidates = [g for g in candidates if g not in already_queried]
        label = "gap-filling"
    else:
        candidates = [g for g in genera if g not in already_queried]
        label = "broad"

    no_evidence = sum(1 for g in genera if g not in evidence)

    print(f"\n{'='*60}", flush=True)
    print(f"PHASE: {phase.upper()} ({label})", flush=True)
    print(f"Total genera: {total} | No evidence yet: {no_evidence} | To query: {len(candidates)}", flush=True)
    print("=" * 60, flush=True)

    if not candidates:
        print("  Nothing to query — all covered.", flush=True)
        return

    hits = 0
    for idx, genus in enumerate(candidates):
        if (idx + 1) % 20 == 0 or idx == 0:
            print(f"  Progress: {idx+1}/{len(candidates)} — hits: {hits}", flush=True)

        if phase == "worms":
            results = extract_worms(genus)
            source_key = "worms"
        elif phase == "gbif":
            results = extract_gbif(genus)
            source_key = "gbif"
        elif phase == "natureserve":
            results = extract_natureserve(genus)
            source_key = "natureserve"
        else:
            print(f"Unknown phase: {phase}", flush=True)
            return

        if results:
            _add_records(evidence, genus, results, source_key)
            hits += 1
            short = results[0]["habitat"][:60]
            print(f"  [{phase.upper()}] {genus}: {short}...", flush=True)

        if (idx + 1) % 50 == 0:
            _save_evidence(evidence)

    _save_evidence(evidence)

    found_total = sum(1 for g in genera if g in evidence)
    manual = total - found_total
    print(f"\n{'='*60}", flush=True)
    print(f"PHASE {phase.upper()} DONE", flush=True)
    print(f"Hits this phase: {hits}", flush=True)
    print(f"Total with any evidence: {found_total}/{total}", flush=True)
    print(f"No evidence yet: {manual}", flush=True)
    print("=" * 60, flush=True)


def main():
    import sys
    if len(sys.argv) > 1:
        phase = sys.argv[1].lower()
        is_gap = phase in TIER2_SOURCES
        run_phase(phase, only_missing=is_gap)
    else:
        for phase in TIER1_SOURCES:
            run_phase(phase, only_missing=False)
        for phase in TIER2_SOURCES:
            run_phase(phase, only_missing=True)
        genera = _get_genera()
        evidence = _load_evidence()
        found_total = sum(1 for g in genera if g in evidence)
        print(f"\n{'='*60}", flush=True)
        print(f"ALL PHASES DONE", flush=True)
        print(f"Total with evidence: {found_total}/{len(genera)}", flush=True)
        print(f"Manual: {len(genera) - found_total}", flush=True)
        print("=" * 60, flush=True)


if __name__ == "__main__":
    main()
