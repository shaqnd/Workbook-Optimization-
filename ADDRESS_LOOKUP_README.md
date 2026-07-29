# Filling the "Property Address" column

`Assigned_Properties_Cleaned.xlsx` now has a **`Property Address`** column
(column K on the **Property Master** tab), added to the existing table and
styled to match. It is currently blank — here is how to fill it.

## Why it isn't already filled

The addresses come from Adams County's public GIS system, keyed on each
`Property ID` (the assessor account number, e.g. `R0002510`). The environment
this workbook was prepared in has a network egress policy that **blocks the
Adams County GIS hosts** (`gisapp.adcogov.org`, `data-adcogov.opendata.arcgis.com`,
etc. all refused the connection), so the lookups could not be run there. The
script below performs them from any machine with normal internet access.

## How to fill it

```bash
pip install openpyxl requests

# 1) See the county's current layers and field names (recommended first run)
python fill_addresses.py --inspect

# 2) Fill all 1,043 addresses into the workbook
python fill_addresses.py
```

- The script **auto-discovers** the parcel layer and the account/address field
  names, then looks up every account number and writes the situs address into
  column K.
- It **caches** every result to `address_cache.json`, so you can stop and
  re-run any time — it resumes and never repeats a lookup.
- Test on a few first with `python fill_addresses.py --limit 5`.
- Accounts with no match are left blank; re-run to retry, or check them by hand
  at <https://gisapp.adcogov.org/PropertySearch>.

If the county reorganizes its GIS services, use `--inspect` to see the current
layers/fields, then override with `--layer-url`, `--acct-field`, and
`--addr-fields`. Run `python fill_addresses.py --help` for all options.
