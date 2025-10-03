#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Create .prj (and .cpg UTF-8) files for Shapefiles using common Norwegian CRS.
- Lists EUREF89 / UTM (32N, 33N, 35N), WGS84 / UTM (32N, 33N, 35N), and NTM zones (5–20).
- User can select by index or type "UTM33", "WGS84/UTM32", "EUREF89/UTM35", or "NTM/10".
- Writes .prj next to each .shp in ./ConvertedDXF2SHP (default) or a folder passed as argv[1].

Notes:
- EUREF89 is the common Scandinavian name for the ETRS89 frame. For CRS/WKT compatibility,
  the GEOGCS/DATUM in WKT remain the EPSG-official "ETRS89" / "ETRS_1989". Labels shown to the user
  and menu keys use "EUREF89".
"""
import os
import sys
import re
from typing import Dict, Tuple, List

# --------------------------
# Configuration / constants
# --------------------------

DEFAULT_FOLDER = os.path.join(".", "ConvertedDXF2SHP")  # same as previous output folder
if not os.path.isdir(DEFAULT_FOLDER):
    os.makedirs(DEFAULT_FOLDER)
    print(f"Created output directory: {DEFAULT_FOLDER}")
CPG_CONTENT = "UTF-8"

# --------------------------
# WKT builders (WKT1)
# --------------------------

def wkt_geogcs_wgs84() -> str:
    return (
        'GEOGCS["WGS 84",'
        'DATUM["WGS_1984",SPHEROID["WGS 84",6378137,298.257223563]],'
        'PRIMEM["Greenwich",0],'
        'UNIT["degree",0.0174532925199433]]'
    )

def wkt_geogcs_etrs89() -> str:
    # EPSG-official naming retained here for maximum compatibility
    # (Even though user-facing labels say EUREF89)
    return (
        'GEOGCS["ETRS89",'
        'DATUM["ETRS_1989",SPHEROID["GRS 1980",6378137,298.257222101]],'
        'PRIMEM["Greenwich",0],'
        'UNIT["degree",0.0174532925199433]]'
    )

def build_wkt_tm(name: str, geogcs: str, lat0: float, lon0: float, k0: float, fe: float, fn: float) -> str:
    # ESRI/OGC WKT1, axis order E,N implied by PROJCS/UNIT
    return (
        f'PROJCS["{name}",'
        f'{geogcs},'
        'PROJECTION["Transverse_Mercator"],'
        f'PARAMETER["latitude_of_origin",{lat0}],'
        f'PARAMETER["central_meridian",{lon0}],'
        f'PARAMETER["scale_factor",{k0}],'
        f'PARAMETER["false_easting",{fe}],'
        f'PARAMETER["false_northing",{fn}],'
        'UNIT["metre",1]]'
    )

def wkt_utm_wgs84(zone: int) -> Tuple[str, str]:
    # UTM north zones: lon0 = 6*zone - 183 ; scale 0.9996 ; FE 500000 ; FN 0
    lon0 = 6 * zone - 183
    name = f'WGS 84 / UTM zone {zone}N'
    return name, build_wkt_tm(name, wkt_geogcs_wgs84(), 0, lon0, 0.9996, 500000.0, 0.0)

def wkt_utm_euref89(zone: int) -> Tuple[str, str]:
    # Same numeric params as ETRS89; user-facing label says EUREF89
    lon0 = 6 * zone - 183
    name = f'EUREF89 / UTM zone {zone}N'
    return name, build_wkt_tm(name, wkt_geogcs_etrs89(), 0, lon0, 0.9996, 500000.0, 0.0)

def wkt_ntm(zone: int) -> Tuple[str, str]:
    """
    NTM: 1° wide zones centered at (zone + 0.5)°E, scale=1, FE=100000, FN=1000000
    Label: EUREF89 / NTM zone {zone} (internally uses ETRS89 GEOGCS)
    """
    name = f"EUREF89 / NTM zone {zone}"
    lon0 = zone + 0.5
    return name, build_wkt_tm(name, wkt_geogcs_etrs89(), 0.0, lon0, 1.0, 100000.0, 1000000.0)

# --------------------------
# Options list
# --------------------------

def build_options() -> List[Tuple[str, str, str]]:
    """
    Returns list of (label, key, wkt) where:
      - label: human-readable label shown in menu
      - key: canonical key like 'EUREF89/UTM33', 'WGS84/UTM32', 'NTM/10'
      - wkt: WKT string
    """
    items: List[Tuple[str, str, str]] = []

    # EUREF89 / UTM commonly used in Norway: zones 32, 33, 35 (EPSG:25832/25833/25835)
    for z in (32, 33, 35):
        name, wkt = wkt_utm_euref89(z)
        key = f"EUREF89/UTM{z}"
        items.append((f"{name} (EPSG:258{z})", key, wkt))

    # WGS84 / UTM (GNSS/raw or general-purpose): 32, 33, 35 (EPSG:32632/32633/32635)
    for z in (32, 33, 35):
        name, wkt = wkt_utm_wgs84(z)
        key = f"WGS84/UTM{z}"
        items.append((f"{name} (EPSG:326{z})", key, wkt))

    # NTM zones (5–20) per EPSG; used for engineering projects (1° wide bands)
    for z in range(5, 21):
        name, wkt = wkt_ntm(z)
        # EPSG codes are 5105..5120 (zone==code-5100)
        epsg = 5100 + z
        items.append((f"{name} (EPSG:{epsg})", f"NTM/{z}", wkt))

    return items

# --------------------------
# Parsing user input
# --------------------------

def parse_choice(s: str, options: List[Tuple[str, str, str]]) -> Tuple[str, str]:
    """
    Returns (label, wkt) for the chosen CRS. Accepts:
      - integer index into options (1-based)
      - tokens like:
          'UTM33'                   -> EUREF89 default
          'EUREF89/UTM35'           -> EUREF89 explicit
          'WGS/UTM32' or 'WGS84/UTM32' -> WGS84 explicit
          'NTM/10'                  -> EUREF89 NTM zone
    """
    s = s.strip()
    # 1) Numeric index
    if s.isdigit():
        idx = int(s)
        if 1 <= idx <= len(options):
            label, _, wkt = options[idx - 1]
            return label, wkt
        raise ValueError("Index out of range.")

    # 2) Normalized token
    t = s.upper().replace(" ", "")

    # Enforce WGS prefix for WGS UTM:
    # - Accept 'WGS/UTMxx' or 'WGS84/UTMxx' for WGS84
    m = re.fullmatch(r"(WGS|WGS84)/UTM(\d{2})N?", t)
    if m:
        zone = int(m.group(2))
        name, wkt = wkt_utm_wgs84(zone)
        return f"{name} (WKT built)", wkt

    # Default UTM -> EUREF89 (and allow explicit EUREF89/UTMxx too)
    m = re.fullmatch(r"((EUREF89|ETRS89)/)?UTM(\d{2})N?", t)
    if m:
        zone = int(m.group(3))
        name, wkt = wkt_utm_euref89(zone)
        return f"{name} (WKT built)", wkt

    # NTM zones (EUREF89)
    m = re.fullmatch(r"NTM/(\d{1,2})", t)
    if m:
        zone = int(m.group(1))
        if not (5 <= zone <= 20):
            raise ValueError("NTM zone must be 5–20 for mainland Norway.")
        name, wkt = wkt_ntm(zone)
        return f"{name} (WKT built)", wkt

    raise ValueError(
        "Unrecognized input. Use index (1..N), 'UTM33' (EUREF), "
        "'WGS/UTM32' or 'WGS84/UTM32' (WGS), 'EUREF89/UTM35', or 'NTM/10'."
    )

    """
    Returns (label, wkt) for the chosen CRS. Accepts:
      - integer index into options (1-based)
      - tokens like 'UTM33', 'WGS84/UTM32', 'EUREF89/UTM35', or 'NTM/10'
      - Backward compatibility: also accepts 'ETRS89/UTM33'
    """
    s = s.strip()
    # 1) Numeric index
    if s.isdigit():
        idx = int(s)
        if 1 <= idx <= len(options):
            label, _, wkt = options[idx - 1]
            return label, wkt
        raise ValueError("Index out of range.")

    # 2) Text keys
    t = s.upper().replace(" ", "")
    # Normalize variants: 'UTM33' -> assume EUREF89 unless WGS84 specified
    m = re.fullmatch(r"((EUREF89|ETRS89)/)?UTM(\d{2})N?", t)
    if m:
        zone = int(m.group(3))
        # If WGS84 explicitly requested elsewhere, that path is handled below
        name, wkt = wkt_utm_euref89(zone)
        return f"{name} (WKT built)", wkt

    m = re.fullmatch(r"WGS84/UTM(\d{2})N?", t)
    if m:
        zone = int(m.group(1))
        name, wkt = wkt_utm_wgs84(zone)
        return f"{name} (WKT built)", wkt

    m = re.fullmatch(r"NTM/(\d{1,2})", t)
    if m:
        zone = int(m.group(1))
        if not (5 <= zone <= 20):
            raise ValueError("NTM zone must be 5–20 for mainland Norway.")
        name, wkt = wkt_ntm(zone)
        return f"{name} (WKT built)", wkt

    raise ValueError("Unrecognized input. Use index (1..N), 'UTM33', 'WGS84/UTM32', 'EUREF89/UTM35', or 'NTM/10'.")

# --------------------------
# Writing .prj / .cpg
# --------------------------

def write_sidecars(shp_path: str, wkt: str) -> None:
    base, _ = os.path.splitext(shp_path)
    prj_path = base + ".prj"
    cpg_path = base + ".cpg"
    with open(prj_path, "w", encoding="utf-8") as f:
        f.write(wkt + "\n")
    # .cpg to mark DBF encoding
    with open(cpg_path, "w", encoding="ascii", newline="") as f:
        f.write(CPG_CONTENT)
    print(f"  wrote: {os.path.basename(prj_path)}, {os.path.basename(cpg_path)}")

def find_shapefiles(folder: str) -> List[str]:
    return [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".shp")]

# --------------------------
# Main
# --------------------------

def main():
    target_folder = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FOLDER
    if not os.path.isdir(target_folder):
        print(f"Folder not found: {target_folder}")
        sys.exit(1)

    options = build_options()

    print("\nSelect Coordinate Reference System for all .shp in:", os.path.abspath(target_folder))
    print("\nCommon CRSs for Norway (EUREF89 UTM default, WGS UTM only when explicitly requested, and NTM zones):\n")
    for i, (label, key, _) in enumerate(options, start=1):
        print(f"{i:2d}. {label:40s}   key: {key}")

    print("\nEnter the index (e.g., 3) or a key like 'UTM33' (EUREF), 'WGS/UTM32' or 'WGS84/UTM32' (WGS), 'EUREF89/UTM35', or 'NTM/10'")
    choice = input("> ").strip()

    try:
        label, wkt = parse_choice(choice, options)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(2)

    print(f"\nSelected: {label}\n")

    shp_files = find_shapefiles(target_folder)
    if not shp_files:
        print("No .shp files found. Exiting.")
        sys.exit(0)

    for shp in shp_files:
        try:
            write_sidecars(shp, wkt)
        except Exception as e:
            print(f"  FAILED for {os.path.basename(shp)}: {e}")

    print("\nDone.")

if __name__ == "__main__":
    main()
