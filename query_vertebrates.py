#!/usr/bin/env python3
"""
Vertebrate habitat evidence scraper.
Output: vertebrates_full_evidence.json

Sources (2 APIs, 2 tiers):
  Tier 1 — all genera:
    IUCN Red List API v4
      URL: https://api.iucnredlist.org/api/v4/taxa/scientific_name
      Fields: systems (Terrestrial/Freshwater/Marine), habitats (habitat descriptions)
      Genus → species list via GBIF species/search (kingdom=Animalia, phylum=Chordata),
      then per-species IUCN assessment for systems + habitats. Max 5 species per genus.
  Tier 2 — gap-filling (only genera not found via IUCN):
    WoRMS REST API
      URL: https://www.marinespecies.org/rest/AphiaRecordsByName/{genus}
      Fields: isMarine, isBrackish, isFreshwater, isTerrestrial (genus-level)
      Filter: phylum=Chordata. Genus-level flags only (no species breakdown).

Pipeline:
  1. Read genera from Taxon list.xlsx (Group = "Vertebrates")
  2. Tier 1: for each genus, get species list from GBIF → query IUCN per species(need specific species name to query IUCN)
  3. Tier 2: genera with no IUCN results → query WoRMS for genus-level environment
  4. Save all results to JSON 

Run log:
  Last run: 2026-09-08
  Environment: macOS 26.3, Python 3.13.13, openpyxl 3.1.5
  Result: 83/83 genera with evidence (IUCN 77, WoRMS 6)
"""

import json
import time
import urllib.request
import urllib.error
import openpyxl
from pathlib import Path

XLSX_FILE = Path(__file__).parent / "Taxon list.xlsx"
OUTPUT_JSON = Path(__file__).parent / "vertebrates_full_evidence.json"
DELAY = 0.3
MAX_SPECIES = 5

IUCN_TOKEN = "aR94RweFqGF6aQsi2LPa8ZrnVQxpC8WaAHRr"
IUCN_BASE = "https://api.iucnredlist.org/api/v4"
IUCN_HEADERS = {
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)",
    "Authorization": IUCN_TOKEN,
}

WORMS_BASE = "https://www.marinespecies.org/rest"


# === GBIF species list (used for IUCN genus→species lookup) ===

def gbif_species_list(genus):
    """Get species names under a genus from GBIF for IUCN lookup."""
    url = (
        f"https://api.gbif.org/v1/species/search"
        f"?q={genus}&rank=SPECIES&limit=20"
        f"&kingdom=Animalia&phylum=Chordata"
    )
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        resp = urllib.request.urlopen(req, timeout=15)
        data = json.loads(resp.read())
    except Exception:
        return []

    species = []
    for sp in data.get("results", []):
        if sp.get("genus") != genus or sp.get("rank") != "SPECIES":
            continue
        name = sp.get("canonicalName", "")
        parts = name.split()
        if len(parts) == 2:
            species.append((parts[0], parts[1]))
        if len(species) >= 10:
            break
    return species


# === IUCN Red List ===

def extract_iucn(genus):
    """Query IUCN Red List for habitat data. Max 5 species per genus."""
    species_list = gbif_species_list(genus)
    if not species_list:
        return []

    results = []
    for g, s in species_list:
        if len(results) >= MAX_SPECIES:
            break

        url = f"{IUCN_BASE}/taxa/scientific_name?genus_name={g}&species_name={s}"
        try:
            req = urllib.request.Request(url, headers=IUCN_HEADERS)
            resp = urllib.request.urlopen(req, timeout=15)
            data = json.loads(resp.read())
        except Exception:
            continue

        assessments = data.get("assessments", [])
        if not assessments:
            continue

        aid = assessments[0]["assessment_id"]
        try:
            req = urllib.request.Request(
                f"{IUCN_BASE}/assessment/{aid}", headers=IUCN_HEADERS
            )
            resp = urllib.request.urlopen(req, timeout=15)
            full = json.loads(resp.read())
        except Exception:
            continue

        parts = []
        systems = [x["description"]["en"] for x in full.get("systems", [])]
        if systems:
            parts.append(f"System: {', '.join(systems)}")
        habitats = [x["description"]["en"] for x in full.get("habitats", [])]
        if habitats:
            parts.append(f"Habitats: {', '.join(habitats[:5])}")

        if parts:
            results.append({
                "species": f"{g} {s}",
                "habitat": " | ".join(parts),
                "source": "IUCN Red List",
                "url": full.get("url", f"https://www.iucnredlist.org/species/{aid}"),
            })

        time.sleep(0.2)

    return results


# === WoRMS ===

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


def extract_worms(genus):
    url = f"{WORMS_BASE}/AphiaRecordsByName/{genus}?like=false&marine_only=false"
    try:
        req = urllib.request.Request(url, headers={
            "Accept": "application/json", "User-Agent": "Mozilla/5.0"
        })
        resp = urllib.request.urlopen(req, timeout=15)
        data = json.loads(resp.read())
    except Exception:
        return []
    if not data:
        return []

    aphia_id = None
    genus_env = ""
    for r in data:
        if (r.get("rank") == "Genus" and r.get("status") == "accepted"
                and r.get("phylum") == "Chordata"):
            aphia_id = r["AphiaID"]
            genus_env = env_string(r)
            break
    if not aphia_id:
        for r in data:
            if r.get("phylum") == "Chordata":
                aphia_id = r.get("AphiaID")
                genus_env = env_string(r)
                break
    if not aphia_id:
        return []

    if genus_env:
        return [{
            "species": genus,
            "habitat": f"Environment: {genus_env} (genus-level)",
            "source": "WoRMS",
            "url": f"https://www.marinespecies.org/aphia.php?p=taxdetails&id={aphia_id}",
        }]
    return []


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
        if group and str(group).strip() == "Vertebrates":
            genus = str(ws.cell(row=row, column=9).value).strip()
            if genus not in genera:
                genera.append(genus)
    wb.close()

    total = len(genera)
    evidence = _load_evidence()

    # === Tier 1: IUCN ===
    print(f"Scraping {total} vertebrate genera...", flush=True)
    print(f"\n{'='*60}", flush=True)
    print("PHASE: IUCN Red List", flush=True)
    print("=" * 60, flush=True)

    iucn_hits = 0
    for idx, genus in enumerate(genera):
        if genus in evidence and evidence[genus].get("iucn_count", 0) > 0:
            continue

        if (idx + 1) % 10 == 0 or idx == 0:
            print(f"  Progress: {idx+1}/{total} — IUCN hits: {iucn_hits}", flush=True)

        results = extract_iucn(genus)
        if results:
            iucn_hits += 1
            unique_spp = set(r["species"].lower() for r in results)
            evidence[genus] = {
                "unique_species": len(unique_spp),
                "total_records": len(results),
                "iucn_count": len(results),
                "worms_count": 0,
                "all_species": [
                    {"species": r["species"], "habitat": r["habitat"],
                     "source": r["source"], "url": r["url"]}
                    for r in results
                ],
            }
            short = results[0]["habitat"][:80]
            print(f"  [IUCN] {genus} ({len(unique_spp)} spp): {short}...", flush=True)

        if (idx + 1) % 20 == 0:
            _save_evidence(evidence)

    _save_evidence(evidence)

    found_after_iucn = sum(1 for g in genera if g in evidence)
    print(f"\n{'='*60}", flush=True)
    print(f"PHASE IUCN DONE — Hits: {iucn_hits}", flush=True)
    print(f"Total with evidence: {found_after_iucn}/{total}", flush=True)
    print(f"No evidence yet: {total - found_after_iucn}", flush=True)
    print("=" * 60, flush=True)

    # === Tier 2: WoRMS (gap-filling) ===
    missing = [g for g in genera if g not in evidence]

    if missing:
        print(f"\n{'='*60}", flush=True)
        print(f"PHASE: WoRMS (gap-filling)", flush=True)
        print(f"To query: {len(missing)}", flush=True)
        print("=" * 60, flush=True)

        worms_hits = 0
        for idx, genus in enumerate(missing):
            if (idx + 1) % 10 == 0 or idx == 0:
                print(f"  Progress: {idx+1}/{len(missing)} — WoRMS hits: {worms_hits}", flush=True)

            results = extract_worms(genus)
            time.sleep(DELAY)

            if results:
                worms_hits += 1
                unique_spp = set(r["species"].lower() for r in results)
                evidence[genus] = {
                    "unique_species": len(unique_spp),
                    "total_records": len(results),
                    "iucn_count": 0,
                    "worms_count": len(results),
                    "all_species": [
                        {"species": r["species"], "habitat": r["habitat"],
                         "source": r["source"], "url": r["url"]}
                        for r in results
                    ],
                }
                short = results[0]["habitat"][:80]
                print(f"  [WoRMS] {genus}: {short}...", flush=True)

        _save_evidence(evidence)

        found_total = sum(1 for g in genera if g in evidence)
        print(f"\n{'='*60}", flush=True)
        print(f"PHASE WoRMS DONE — Hits: {worms_hits}", flush=True)
        print(f"Total with evidence: {found_total}/{total}", flush=True)
        print(f"No evidence yet: {total - found_total}", flush=True)
        print("=" * 60, flush=True)

    # Final
    found_total = sum(1 for g in genera if g in evidence)
    manual = total - found_total
    missing_list = [g for g in genera if g not in evidence]
    print(f"\n{'='*60}", flush=True)
    print(f"DONE. Output: {OUTPUT_JSON}", flush=True)
    print(f"Total genera: {total}", flush=True)
    print(f"Total with evidence: {found_total}", flush=True)
    print(f"Manual needed: {manual}", flush=True)
    if missing_list:
        print(f"Missing: {', '.join(missing_list)}", flush=True)
    print("=" * 60, flush=True)


if __name__ == "__main__":
    main()
