#!/usr/bin/env python3
"""
fill_addresses.py
=================
Fills the "Property Address" column in Assigned_Properties_Cleaned.xlsx by
looking up each account number (Property ID, e.g. R0002510) against Adams
County, Colorado's public ArcGIS REST service.

WHY THIS IS A SEPARATE SCRIPT
-----------------------------
The environment where this workbook was prepared has an egress policy that
blocks the county GIS hosts, so the lookups could not be run there. Run this
script from any machine with ordinary internet access and it will fill the
column automatically.

USAGE
-----
    pip install openpyxl requests
    python fill_addresses.py --inspect        # show layers/fields (recommended first run)
    python fill_addresses.py                  # fill addresses into the workbook
    python fill_addresses.py --limit 5        # try just 5 accounts as a test

It caches every lookup to address_cache.json, so it is safe to stop and
re-run — it resumes where it left off and never re-queries an account it
already resolved.

The script auto-discovers the parcel layer and the account/address field
names. If the county changes its schema, use --inspect to see the current
layers/fields and override with --layer-url / --acct-field / --addr-fields.
"""

import argparse
import json
import os
import sys
import time

try:
    import requests
except ImportError:
    sys.exit("Missing dependency: pip install requests")
try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency: pip install openpyxl")

# ---------------------------------------------------------------------------
# Configuration (sane defaults; override on the command line if the county
# reorganizes its services).
# ---------------------------------------------------------------------------
DEFAULT_MAPSERVER = "https://gisapp.adcogov.org/arcgis/rest/services/AdvancedExt/MapServer"
WORKBOOK = "Assigned_Properties_Cleaned.xlsx"
SHEET = "Property Master"
ID_HEADER = "Property ID"
ADDR_HEADER = "Property Address"
CACHE_FILE = "address_cache.json"

# Field-name heuristics (matched case-insensitively).
ACCT_FIELD_CANDIDATES = [
    "ACCOUNTNO", "ACCOUNT_NO", "ACCOUNTNUM", "ACCOUNT", "ACCT", "ACCTNO",
    "PIN", "SCHEDULE", "SCHEDULENO", "STRAP", "PARCELNB", "PARCEL_NO", "PARCELID",
]
ADDR_FIELD_CANDIDATES = [
    "SITE_ADDR", "SITEADDR", "SITEADDRESS", "SITUS", "SITUS_ADDR", "SITUSADDR",
    "SITUSADDRESS", "PROP_ADDR", "PROPADDRESS", "FULLADDR", "FULL_ADDRESS",
    "ADDRESS", "SITE_ADDRESS", "LOCATION",
]

HEADERS = {"User-Agent": "property-address-lookup/1.0 (batch join for owned dataset)"}
TIMEOUT = 30


def get_json(url, params=None, tries=4):
    """GET with exponential backoff, returning parsed JSON."""
    params = dict(params or {})
    params.setdefault("f", "json")
    last = None
    for attempt in range(tries):
        try:
            r = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except Exception as e:  # noqa: BLE001
            last = e
            wait = 2 ** attempt
            print(f"  ... request failed ({e}); retry in {wait}s", file=sys.stderr)
            time.sleep(wait)
    raise RuntimeError(f"Giving up on {url}: {last}")


def list_layers(mapserver):
    meta = get_json(mapserver)
    layers = meta.get("layers", []) + meta.get("tables", [])
    return [(l["id"], l.get("name", "")) for l in layers]


def layer_fields(mapserver, layer_id):
    meta = get_json(f"{mapserver}/{layer_id}")
    return meta.get("fields", []) or []


def pick(candidates, field_names):
    """Return the first candidate present in field_names (case-insensitive),
    else the first field_name that *contains* a candidate token."""
    upper = {f.upper(): f for f in field_names}
    for c in candidates:
        if c in upper:
            return upper[c]
    for c in candidates:
        for fu, orig in upper.items():
            if c in fu:
                return orig
    return None


def inspect(mapserver):
    print(f"MapServer: {mapserver}\n")
    for lid, name in list_layers(mapserver):
        print(f"[{lid}] {name}")
        try:
            fields = layer_fields(mapserver, lid)
        except Exception as e:  # noqa: BLE001
            print(f"     (could not read fields: {e})")
            continue
        for f in fields:
            print(f"     - {f.get('name')}  ({f.get('type')})  {f.get('alias','')}")
        print()


def discover(mapserver, sample_acct):
    """Find (layer_url, acct_field, addr_field) that actually resolves an account."""
    layers = list_layers(mapserver)
    # Prefer layers whose name hints at parcels/addresses.
    def score(name):
        n = name.lower()
        s = 0
        for kw, w in (("parcel", 3), ("address", 2), ("property", 2), ("account", 1)):
            if kw in n:
                s += w
        return s
    layers.sort(key=lambda x: score(x[1]), reverse=True)

    for lid, name in layers:
        try:
            fields = layer_fields(mapserver, lid)
        except Exception:
            continue
        fnames = [f["name"] for f in fields]
        acct = pick(ACCT_FIELD_CANDIDATES, fnames)
        addr = pick(ADDR_FIELD_CANDIDATES, fnames)
        if not acct:
            continue
        layer_url = f"{mapserver}/{lid}"
        # Test the sample account against this layer/field.
        for variant in acct_variants(sample_acct):
            res = query_account(layer_url, acct, variant)
            if res and res.get("features"):
                attrs = res["features"][0]["attributes"]
                if not addr:
                    addr = pick(ADDR_FIELD_CANDIDATES, list(attrs.keys()))
                print(f"Discovered: layer [{lid}] '{name}'  acct-field={acct}  "
                      f"addr-field={addr}  (matched '{variant}')")
                return layer_url, acct, addr
    raise SystemExit(
        "Could not auto-discover the parcel layer/fields.\n"
        "Run with --inspect, then pass --layer-url, --acct-field and --addr-fields."
    )


def acct_variants(acct):
    """Account numbers can be stored with/without the leading letter or
    zero-padding. Yield sensible variants to try, most-likely first."""
    a = str(acct).strip()
    seen, out = set(), []
    for v in (a, a.upper(), a.lstrip("Rr"), a.lstrip("Rr").lstrip("0")):
        v = v.strip()
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def query_account(layer_url, acct_field, value):
    # Escape single quotes for the SQL where-clause.
    safe = value.replace("'", "''")
    return get_json(
        f"{layer_url}/query",
        {
            "where": f"UPPER({acct_field})='{safe.upper()}'",
            "outFields": "*",
            "returnGeometry": "false",
            "resultRecordCount": 1,
        },
    )


def build_address(attrs, addr_field):
    """Prefer the single situs/site address field; otherwise stitch common
    house-number / street / city / zip components together."""
    if addr_field and attrs.get(addr_field):
        return str(attrs[addr_field]).strip()
    # Fallback: assemble from parts if a single field wasn't found.
    parts_keys = [
        ["HOUSE_NO", "HOUSENO", "STR_NUM", "ADDRNO"],
        ["PREDIR", "PRE_DIR"],
        ["STREET", "ST_NAME", "STREETNAME", "STR_NAME"],
        ["STREETTYPE", "ST_TYPE", "STRTYPE"],
        ["POSTDIR", "POST_DIR"],
        ["UNIT", "UNIT_NO"],
    ]
    upper = {k.upper(): k for k in attrs}
    chunk = []
    for group in parts_keys:
        for cand in group:
            if cand in upper and attrs[upper[cand]] not in (None, "", " "):
                chunk.append(str(attrs[upper[cand]]).strip())
                break
    line1 = " ".join(chunk).strip()
    city = next((attrs[upper[c]] for c in ("CITY", "SITE_CITY", "SITUS_CITY")
                 if c in upper and attrs[upper[c]]), "")
    zc = next((attrs[upper[c]] for c in ("ZIP", "ZIPCODE", "SITE_ZIP", "ZIP_CODE")
               if c in upper and attrs[upper[c]]), "")
    tail = ", ".join(p for p in [str(city).strip(), str(zc).strip()] if p)
    full = ", ".join(p for p in [line1, tail] if p)
    return full.strip()


def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE) as fh:
            return json.load(fh)
    return {}


def save_cache(cache):
    tmp = CACHE_FILE + ".tmp"
    with open(tmp, "w") as fh:
        json.dump(cache, fh, indent=0)
    os.replace(tmp, CACHE_FILE)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", default=WORKBOOK)
    ap.add_argument("--mapserver", default=DEFAULT_MAPSERVER)
    ap.add_argument("--layer-url", help="Skip discovery; query this layer URL directly")
    ap.add_argument("--acct-field", help="Override account field name")
    ap.add_argument("--addr-fields", help="Comma-separated address field name(s) to try")
    ap.add_argument("--inspect", action="store_true", help="Print layers/fields and exit")
    ap.add_argument("--limit", type=int, default=0, help="Only process the first N unresolved accounts")
    ap.add_argument("--sleep", type=float, default=0.15, help="Delay between requests (seconds)")
    args = ap.parse_args()

    if args.inspect:
        inspect(args.mapserver)
        return

    wb = openpyxl.load_workbook(args.workbook)
    ws = wb[SHEET]
    headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    if ID_HEADER not in headers:
        sys.exit(f"Column '{ID_HEADER}' not found in sheet '{SHEET}'.")
    if ADDR_HEADER not in headers:
        sys.exit(f"Column '{ADDR_HEADER}' not found. Add it first.")
    id_col, addr_col = headers[ID_HEADER], headers[ADDR_HEADER]

    accts = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=id_col).value
        if v:
            accts.append((r, str(v).strip()))
    print(f"{len(accts)} accounts in workbook.")

    cache = load_cache()

    # Establish the query target.
    if args.layer_url and args.acct_field:
        layer_url, acct_field = args.layer_url, args.acct_field
        addr_field = (args.addr_fields.split(",")[0] if args.addr_fields else None)
    else:
        sample = accts[0][1]
        layer_url, acct_field, addr_field = discover(args.mapserver, sample)

    processed = 0
    for r, acct in accts:
        if acct in cache and cache[acct]:
            ws.cell(row=r, column=addr_col).value = cache[acct]
            continue
        if args.limit and processed >= args.limit:
            break
        processed += 1
        address = ""
        for variant in acct_variants(acct):
            try:
                res = query_account(layer_url, acct_field, variant)
            except Exception as e:  # noqa: BLE001
                print(f"  {acct}: error {e}", file=sys.stderr)
                res = None
            if res and res.get("features"):
                address = build_address(res["features"][0]["attributes"], addr_field)
                break
            time.sleep(args.sleep)
        cache[acct] = address
        ws.cell(row=r, column=addr_col).value = address
        status = address if address else "(no match found)"
        print(f"  {acct} -> {status}")
        if processed % 25 == 0:
            save_cache(cache)
            wb.save(args.workbook)
        time.sleep(args.sleep)

    save_cache(cache)
    wb.save(args.workbook)
    filled = sum(1 for _, a in accts if cache.get(a))
    print(f"\nDone. {filled}/{len(accts)} addresses filled. Saved to {args.workbook}.")
    if filled < len(accts):
        print("Accounts with no match remain blank — re-run to retry, or check "
              "them individually at https://gisapp.adcogov.org/PropertySearch")


if __name__ == "__main__":
    main()
