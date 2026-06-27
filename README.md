# Dorchester S-zone parcel screening

I wrote this script for one narrow job: find FY2026 Boston parcels that
intersect both the BPDA Dorchester boundary and a numeric S-zone, then keep
the parcels whose assessor information supports vacant land, single-family,
or two-family use.

The result is a screening list. It is not a legal zoning opinion, and the
script does not finish the client's requested manual check in the Boston
Zoning Viewer. That distinction matters because the workbook keeps the
client's requested column name, `Verified Subdistrict`, even though the value
in that column comes from the BPDA zoning layer and a spatial join. One still
needs to check each final parcel in the viewer before describing it as manually
verified.

## What the script uses

All three automated inputs are official Boston ArcGIS REST layers:

- BPDA Boston Neighborhood Boundaries
- BPDA Zoning Subdistricts
- FY2026 Parcels Detailed

The exact service URLs are stored in `Config` near the top of `boston.py` and
are also written to the workbook's `Methodology` sheet. The Boston Zoning
Viewer is a separate manual verification source; the script does not scrape
or control it.

## How a parcel reaches the workbook

The script first downloads the neighborhood and zoning layers. It selects the
feature named Dorchester and zoning labels that match a numeric S-zone form,
such as `S0`, `S-1`, or `S-.5`. It then intersects the S-zone polygons with the
Dorchester polygon. This clipped geometry is the final geographic test.

For parcel retrieval, the script uses the rectangular search envelope around
the full S-zone polygons that touch Dorchester. That is only a way to obtain a
manageable candidate pool. A parcel is not accepted merely because it falls
inside that rectangle.

The assessor-use test rejects known non-target uses first. Examples include
commercial, apartment, mixed-use, R3, R4, parking, office, retail, industrial,
bank, condominium, garage, and three-family indicators. A remaining parcel
can pass through one of these kinds of evidence:

- `LU` is `R1`, `R2`, or `RL`.
- `LUC` is one of the allowed residential or land codes in `GOOD_LUC`.
- The assessor description says residential land, vacant, one-family,
  single-family, or two-family.
- The record reports one or two residential units and no commercial or
  residential-commercial units.
- The record has land area but no building, living, or gross area.

Finally, GeoPandas performs an `intersects` spatial join between those
use-filtered parcels and the Dorchester-clipped S-zone geometry. This is an
intersection test, not a requirement that the whole parcel lie inside one
polygon. Parcels on a neighborhood or zoning boundary therefore deserve
extra attention during manual review.

If one parcel intersects more than one qualifying zoning polygon, the script
combines the distinct subdistrict labels into one cell. It then sorts the
output by lot size, largest first, and limits the sheet to 40 rows. It does
not invent additional rows when fewer than 30 parcels qualify.

## Workbook

Running `boston.py` writes `Verified_Dorchester_S_Zone_Parcels.xlsx` in the
directory from which the command is run.

The workbook has two sheets. `Parcels` contains the ten fields requested by
the client: parcel ID, address, subdistrict, lot size, current use, owner,
mailing address, year built, last sale year, and the mailing-address flag.
`Methodology` records the source URLs, feature counts, filtering rules, final
row count, and important limitations from that run.

The mailing-address flag compares a normalized property street address with
the mailing street address. It does not prove that two complete postal
addresses belong to different places. If no mailing street is present, the
current code returns `No`, so that field also needs a quick manual check.

## What the automated result does not prove

The current pipeline has three known gaps against the full client request.

First, the last successful strict run produced 14 qualifying parcels, not the
requested 30 to 40. That number can change if Boston updates its data, but the
script will not add parcels that fail the stated geography or use tests just
to reach a quota.

Second, the FY2026 parcel service used here does not expose a usable sale-year
or sale-date field under the fields the script checks. `Last Sale Year` is
therefore blank unless the service later supplies one of those fields. The
script does not infer a year from unrelated information.

Third, the zoning values are GIS-derived. They have not been manually
confirmed merely because they appear under the heading `Verified
Subdistrict`. Manual viewer confirmation is still required before delivery.

The assessor fields are evidence, not perfect descriptions of present-day
conditions. This script also does not determine buildability, ownership
quality, title status, permitting, or legal compliance.

## Manual checks before delivery

For each parcel in the final sheet, one needs to:

1. Find the parcel in the Boston Zoning Viewer.
2. Confirm the numeric S-zone shown by the viewer.
3. Pay special attention to parcels touching a boundary.
4. Confirm the assessor use and the mailing-address flag.
5. Obtain the last sale year from a reliable official source and cite that
   source in the methodology.
6. Correct or remove any row that does not survive those checks.

If only 14 rows remain after that review, the honest deliverable is a 14-row
strict result with the shortfall explained. A 30-row workbook containing
nonconforming parcels would not satisfy the assignment.

## Running the script

This project declares Python 3.13 and its dependencies in `pyproject.toml`.
With `uv` installed:

```powershell
uv sync
uv run python boston.py
```

The Boston GIS services occasionally time out. The request code makes up to
four attempts with increasing waits before it stops with a `CRITICAL` error.
Rerunning after a service timeout is reasonable; changing the selection rules
because an endpoint is temporarily unavailable is not.

## Reading the code

The shortest route through the program starts at `main()`. It calls
`run_pipeline()`, which creates the HTTP session, loads the spatial inputs,
builds the parcel list, records the run counts, and writes the workbook.

`load_inputs()` handles neighborhoods and zoning. `build_parcels()` fetches
the parcel candidates, applies the use rules, performs the strict spatial
join, and builds the client-facing rows. The smaller helper functions deal
with one concern at a time: REST requests, geometry cleanup, use evidence,
address formatting, methodology notes, or workbook formatting.

The point of this README is to make the script auditable. I should be able to
explain why a parcel was included, where each workbook value came from, and
which parts still depend on manual work without treating the code as a black
box.
