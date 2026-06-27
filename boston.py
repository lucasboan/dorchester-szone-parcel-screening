"""Build a strict Dorchester S-zone Stage 1 parcel workbook."""

from __future__ import annotations

import re
import sys
import time
from dataclasses import dataclass
from typing import Any, Iterable

import geopandas as gpd
import pandas as pd
import requests
from openpyxl.styles import Font

try:
    from shapely import make_valid
except ImportError:
    make_valid = None


@dataclass(frozen=True)
class Config:
    """Store runtime constants for the Boston parcel pipeline."""

    crs: str = "EPSG:2249"
    epsg: str = "2249"
    batch_size: int = 500
    request_timeout: int = 60
    max_http_attempts: int = 4
    retry_backoff: float = 5.0
    request_pause: float = 0.15
    max_rows: int = 40
    out_file: str = "Verified_Dorchester_S_Zone_Parcels.xlsx"
    viewer_url: str = "https://maps.bostonplans.org/zoningviewer/"
    neighborhood_url: str = (
        "https://gis.bostonplans.org/hosting/rest/services/Hosted/"
        "Boston_Neighborhood_Boundaries/FeatureServer/5/query"
    )
    zoning_url: str = (
        "https://gis.bostonplans.org/hosting/rest/services/"
        "Zoning_Subdistricts_Data/FeatureServer/0/query"
    )
    parcel_url: str = (
        "https://gis.bostonplans.org/hosting/rest/services/"
        "FY26_Parcels_Detailed/FeatureServer/0/query"
    )


CFG = Config()
SZONE_RE = r"^S[-]?(?:\d+(?:\.\d+)?|\.\d+)$"
DORCHESTER_ZIPS = {"02121", "02122", "02124", "02125"}
GOOD_LU = {"R1", "R2", "RL"}
GOOD_LUC = {"101", "104", "130", "131", "132"}
BAD_CODES = {
    "A",
    "C",
    "CC",
    "CD",
    "CL",
    "CM",
    "CP",
    "E",
    "I",
    "RC",
    "R3",
    "R4",
}
BAD_TEXT = (
    "APARTMENT",
    "BANK",
    "COMMERCIAL",
    "CONDO",
    "GARAGE",
    "INDUSTRIAL",
    "MIXED",
    "OFFICE",
    "PARKING",
    "RETAIL",
    "STORE",
    "THREE",
    "3-FAM",
    "3 FAM",
    "WAREHOUSE",
)
SALE_COLUMNS = (
    "SALE_YEAR",
    "SALE_YR",
    "LAST_SALE_YEAR",
    "LST_SAL_YR",
    "SALE_DATE",
    "LAST_SALE",
    "LST_SAL_DT",
)
CLIENT_COLUMNS = [
    "Parcel ID",
    "Address",
    "Verified Subdistrict",
    "Lot Size (sq ft)",
    "Current Use",
    "Owner Name",
    "Owner Mailing Address",
    "Year Built",
    "Last Sale Year",
    "Owner Mailing Different From Property",
]


class PipelineError(RuntimeError):
    """Represent an expected pipeline failure."""


def metadata_url(query_url: str) -> str:
    """Return the ArcGIS metadata URL for a query endpoint."""
    if query_url.lower().endswith("/query"):
        return query_url[:-6]
    return query_url


def request_json_once(
    session: requests.Session,
    url: str,
    params: dict[str, Any],
    post: bool,
) -> dict[str, Any]:
    """Request one JSON response from an ArcGIS endpoint."""
    if post:
        response = session.post(url, data=params, timeout=CFG.request_timeout)
    else:
        response = session.get(url, params=params, timeout=CFG.request_timeout)
    response.raise_for_status()
    return response.json()


def request_json(
    session: requests.Session,
    url: str,
    params: dict[str, Any],
    post: bool = True,
) -> dict[str, Any]:
    """Request JSON with retry handling for transient network failures."""
    last_error: Exception | None = None
    for attempt in range(1, CFG.max_http_attempts + 1):
        try:
            data = request_json_once(session, url, params, post)
            if isinstance(data, dict) and "error" in data:
                raise PipelineError(f"ArcGIS error at {url}: {data['error']}")
            return data
        except PipelineError:
            raise
        except (requests.RequestException, ValueError) as error:
            last_error = error
            retry_wait(attempt, url, error)
    raise PipelineError(f"HTTP failure at {url}: {last_error}") from last_error


def retry_wait(attempt: int, url: str, error: Exception) -> None:
    """Pause before retrying a failed HTTP request."""
    if attempt >= CFG.max_http_attempts:
        return
    delay = CFG.retry_backoff * attempt
    print(f"HTTP attempt {attempt} failed at {url}: {error}. Retrying...")
    time.sleep(delay)


def layer_metadata(
    session: requests.Session,
    query_url: str,
) -> dict[str, Any]:
    """Fetch ArcGIS layer metadata."""
    return request_json(
        session,
        metadata_url(query_url),
        {"f": "json"},
        post=False,
    )


def object_ids(
    session: requests.Session,
    url: str,
    where: str,
    geometry: dict[str, str] | None = None,
) -> list[int]:
    """Fetch object IDs from an ArcGIS layer."""
    params: dict[str, Any] = {
        "where": where,
        "returnIdsOnly": "true",
        "f": "json",
    }
    if geometry:
        params.update(geometry)
    data = request_json(session, url, params)
    return [int(value) for value in data.get("objectIds") or []]


def chunks(values: list[int], size: int) -> Iterable[list[int]]:
    """Yield fixed-size chunks from a list."""
    for start in range(0, len(values), size):
        yield values[start : start + size]


def geojson_params(batch: list[int]) -> dict[str, str]:
    """Build ArcGIS GeoJSON request parameters for object IDs."""
    return {
        "objectIds": ",".join(str(value) for value in batch),
        "outFields": "*",
        "f": "geojson",
        "returnGeometry": "true",
        "outSR": CFG.epsg,
    }


def fetch_features(
    session: requests.Session,
    url: str,
    ids: list[int],
    label: str,
) -> gpd.GeoDataFrame:
    """Fetch GeoJSON features by ArcGIS object IDs."""
    frames: list[gpd.GeoDataFrame] = []
    total = (len(ids) + CFG.batch_size - 1) // CFG.batch_size
    for index, batch in enumerate(chunks(ids, CFG.batch_size), start=1):
        frame = fetch_feature_batch(session, url, batch)
        if not frame.empty:
            frames.append(frame)
        print(f"Fetched {label} batch {index}/{total}: {len(frame)}")
        time.sleep(CFG.request_pause)
    return combine_frames(frames)


def fetch_feature_batch(
    session: requests.Session,
    url: str,
    batch: list[int],
) -> gpd.GeoDataFrame:
    """Fetch one GeoJSON feature batch."""
    data = request_json(session, url, geojson_params(batch))
    features = [
        item for item in data.get("features", [])
        if item.get("geometry")
    ]
    if not features:
        return gpd.GeoDataFrame(geometry=[], crs=CFG.crs)
    return gpd.GeoDataFrame.from_features(features, crs=CFG.crs)


def combine_frames(frames: list[gpd.GeoDataFrame]) -> gpd.GeoDataFrame:
    """Combine fetched GeoDataFrames into one clean GeoDataFrame."""
    if not frames:
        return gpd.GeoDataFrame(geometry=[], crs=CFG.crs)
    return clean_geometry(pd.concat(frames, ignore_index=True))


def clean_one_geometry(geometry: Any) -> Any:
    """Return a valid geometry or None."""
    if geometry is None or geometry.is_empty:
        return None
    if geometry.is_valid:
        return geometry
    if make_valid is not None:
        return make_valid(geometry)
    return geometry.buffer(0)


def clean_geometry(frame: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Normalize CRS and repair invalid geometries."""
    if frame.empty:
        return gpd.GeoDataFrame(geometry=[], crs=CFG.crs)
    clean = frame.to_crs(CFG.crs).copy()
    clean = clean[clean.geometry.notna()].copy()
    clean["geometry"] = clean.geometry.apply(clean_one_geometry)
    clean = clean[clean.geometry.notna()].copy()
    return clean[~clean.geometry.is_empty].reset_index(drop=True)


def require_column(frame: pd.DataFrame, column: str, label: str) -> None:
    """Require that a DataFrame contains one named column."""
    if column in frame.columns:
        return
    columns = ", ".join(map(str, frame.columns))
    raise PipelineError(f"{label} lacks {column!r}. Columns: {columns}")


def first_column(frame: pd.DataFrame, names: list[str]) -> str:
    """Return the first available column name from candidates."""
    for name in names:
        if name in frame.columns:
            return name
    raise PipelineError(f"Missing all columns: {names}")


def text_series(frame: pd.DataFrame, names: list[str]) -> pd.Series:
    """Return the first available text series or blanks."""
    for name in names:
        if name in frame.columns:
            return frame[name].fillna("").astype(str)
    return pd.Series([""] * len(frame), index=frame.index, dtype="object")


def number_series(frame: pd.DataFrame, names: list[str]) -> pd.Series:
    """Return the first available numeric series or zeros."""
    series = text_series(frame, names).str.replace(",", "", regex=False)
    return pd.to_numeric(series, errors="coerce").fillna(0)


def load_layer(
    session: requests.Session,
    url: str,
    label: str,
) -> gpd.GeoDataFrame:
    """Load a complete ArcGIS feature layer."""
    meta = layer_metadata(session, url)
    object_id_field = meta.get("objectIdField", "OBJECTID")
    print(f"{label}: object ID field is {object_id_field!r}")
    ids = object_ids(session, url, "1=1")
    print(f"{label}: {len(ids)} IDs returned")
    return fetch_features(session, url, ids, label)


def dorchester_polygon(neighborhoods: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Extract the BPDA Dorchester neighborhood polygon."""
    require_column(neighborhoods, "name", "Neighborhood layer")
    names = neighborhoods["name"].astype(str).str.upper().str.strip()
    dorchester = neighborhoods[names.eq("DORCHESTER")].copy()
    if dorchester.empty:
        raise PipelineError("Dorchester missing from neighborhood layer")
    return clean_geometry(dorchester)


def s_zone_polygons(zoning: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Extract numeric S-zone zoning subdistrict polygons."""
    require_column(zoning, "Zoning_Subdistrict", "Zoning layer")
    values = zoning["Zoning_Subdistrict"]
    values = values.fillna("").astype(str).str.upper().str.strip()
    zones = zoning[values.str.contains(SZONE_RE, regex=True, na=False)].copy()
    if zones.empty:
        raise PipelineError("No numeric S-zone values were found")
    zones["Zoning_Subdistrict"] = values.loc[zones.index]
    extracted = sorted(zones["Zoning_Subdistrict"].unique())
    print(f"S-zone values extracted: {extracted}")
    return clean_geometry(zones)


def dorchester_s_zone_mask(
    zones: gpd.GeoDataFrame,
    dorchester: gpd.GeoDataFrame,
) -> gpd.GeoDataFrame:
    """Clip S-zone polygons to the BPDA Dorchester neighborhood."""
    mask = gpd.overlay(
        zones,
        dorchester,
        how="intersection",
        keep_geom_type=True,
    )
    mask = clean_geometry(mask)
    if mask.empty:
        raise PipelineError("No S-zone area intersects Dorchester")
    print(f"Dorchester clipped S-zone polygons: {len(mask)}")
    return mask


def touching_s_zones(
    zones: gpd.GeoDataFrame,
    dorchester: gpd.GeoDataFrame,
) -> gpd.GeoDataFrame:
    """Select full S-zone polygons touching Dorchester."""
    joined = gpd.sjoin(zones, dorchester, how="inner", predicate="intersects")
    joined = joined.drop(columns=["index_right"], errors="ignore")
    subset = ["OBJECTID"] if "OBJECTID" in joined.columns else None
    touching = joined.drop_duplicates(subset=subset).copy()
    print(f"Full S-zone polygons touching Dorchester: {len(touching)}")
    return clean_geometry(touching)


def envelope(frame: gpd.GeoDataFrame) -> dict[str, str]:
    """Build ArcGIS envelope query parameters for a GeoDataFrame."""
    xmin, ymin, xmax, ymax = frame.total_bounds
    return {
        "geometry": f"{xmin},{ymin},{xmax},{ymax}",
        "geometryType": "esriGeometryEnvelope",
        "inSR": CFG.epsg,
        "spatialRel": "esriSpatialRelIntersects",
    }


def row_text(row: pd.Series) -> str:
    """Join available assessor use text for one row."""
    columns = [
        "LU",
        "LUC",
        "LU_DESC",
        "LU_Actual",
        "BLDG_TYPE",
        "STRUCTURE_CLASS",
    ]
    values = [str(row.get(col, "")).upper() for col in columns]
    return " ".join(value for value in values if value and value != "NAN")


def has_bad_use(row: pd.Series) -> bool:
    """Return True when assessor evidence shows excluded use."""
    code = str(row.get("LU", "")).upper().strip()
    luc = str(row.get("LUC", "")).upper().strip()
    if code in BAD_CODES or luc in BAD_CODES:
        return True
    return any(token in row_text(row) for token in BAD_TEXT)


def code_use_basis(row: pd.Series) -> str:
    """Return inclusion basis from LU or LUC codes."""
    code = str(row.get("LU", "")).upper().strip()
    luc = str(row.get("LUC", "")).upper().strip()
    if code in GOOD_LU:
        return f"LU code {code}"
    if luc in GOOD_LUC:
        return f"LUC code {luc}"
    return ""


def text_use_basis(row: pd.Series) -> str:
    """Return inclusion basis from assessor use descriptions."""
    text = row_text(row)
    patterns = {
        r"\bRESIDENTIAL LAND\b": "description residential land",
        r"\bVACANT\b": "description vacant land",
        r"\bSINGLE[-\s]?FAMILY\b": "description single family",
        r"\bONE[-\s]?FAMILY\b": "description one family",
        r"\bTWO[-\s]?FAMILY\b": "description two family",
        r"\b2[-\s]?FAMILY\b": "description two family",
    }
    for pattern, label in patterns.items():
        if re.search(pattern, text):
            return label
    return ""


def unit_use_basis(row: pd.Series) -> str:
    """Return inclusion basis from residential unit counts."""
    res = pd.to_numeric(row.get("RES_UNITS", 0), errors="coerce")
    com = pd.to_numeric(row.get("COM_UNITS", 0), errors="coerce")
    rc_units = pd.to_numeric(row.get("RC_UNITS", 0), errors="coerce")
    if res in (1, 2) and (pd.isna(com) or com == 0):
        if pd.isna(rc_units) or rc_units == 0:
            return f"RES_UNITS {int(res)}"
    return ""


def vacant_use_basis(row: pd.Series) -> str:
    """Return inclusion basis from vacant physical indicators."""
    buildings = pd.to_numeric(row.get("NUM_BLDGS", 0), errors="coerce")
    living = pd.to_numeric(row.get("LIVING_AREA", 0), errors="coerce")
    gross = pd.to_numeric(row.get("GROSS_AREA", 0), errors="coerce")
    land = pd.to_numeric(row.get("LAND_SF", 0), errors="coerce")
    if (pd.isna(buildings) or buildings == 0) and land > 0:
        if living == 0 and gross == 0:
            return "vacant physical indicators"
    return ""


def use_basis(row: pd.Series) -> str:
    """Return the target-use inclusion basis for one parcel."""
    if has_bad_use(row):
        return ""
    return (
        code_use_basis(row)
        or text_use_basis(row)
        or unit_use_basis(row)
        or vacant_use_basis(row)
    )


def fetch_candidate_parcels(
    session: requests.Session,
    search_mask: gpd.GeoDataFrame,
) -> gpd.GeoDataFrame:
    """Fetch and use-filter parcels near Dorchester-touching S-zones."""
    ids = object_ids(session, CFG.parcel_url, "1=1", envelope(search_mask))
    print(f"Parcel IDs inside search envelope: {len(ids)}")
    parcels = fetch_features(session, CFG.parcel_url, ids, "parcels")
    parcel_id = first_column(parcels, ["MAP_PAR_ID", "PID"])
    parcels = parcels.drop_duplicates(subset=[parcel_id]).copy()
    parcels["Use_Basis"] = parcels.apply(use_basis, axis=1)
    keep = parcels[parcels["Use_Basis"].ne("")].copy()
    print(f"Parcels after target-use rules: {len(keep)}")
    return keep


def zoning_column(frame: pd.DataFrame) -> str:
    """Find the zoning subdistrict column after a spatial join."""
    choices = ["Zoning_Subdistrict", "Zoning_Subdistrict_right"]
    for column in choices:
        if column in frame.columns:
            return column
    matches = [
        col for col in frame.columns
        if col.startswith("Zoning_Subdistrict")
    ]
    if matches:
        return matches[0]
    raise PipelineError("No zoning subdistrict column after spatial join")


def strict_join(
    parcels: gpd.GeoDataFrame,
    mask: gpd.GeoDataFrame,
) -> gpd.GeoDataFrame:
    """Spatially join parcels to Dorchester-clipped S-zone geometry."""
    joined = gpd.sjoin(parcels, mask, how="inner", predicate="intersects")
    joined = joined.drop(columns=["index_right"], errors="ignore")
    if joined.empty:
        raise PipelineError("No parcels survived strict S-zone join")
    return joined


def aggregate_by_parcel(joined: gpd.GeoDataFrame) -> pd.DataFrame:
    """Aggregate duplicate zoning hits per parcel."""
    parcel_id = first_column(joined, ["MAP_PAR_ID", "PID"])
    zone_col = zoning_column(joined)
    values = joined.drop(columns=["geometry"], errors="ignore").copy()
    functions = {
        column: "first"
        for column in values.columns
        if column != parcel_id
    }
    functions[zone_col] = join_unique_text
    return values.groupby(parcel_id, as_index=False).agg(functions)


def join_unique_text(values: pd.Series) -> str:
    """Join unique nonblank text values in sorted order."""
    unique = set(values.dropna().astype(str))
    return ", ".join(sorted(unique))


def normalize_address(value: Any) -> str:
    """Normalize address text for comparison."""
    if pd.isna(value):
        return ""
    text = " ".join(re.sub(r"[^\w\s]", " ", str(value).upper()).split())
    text = re.sub(r"\bSTREET\b", "ST", text)
    text = re.sub(r"\bAVENUE\b", "AVE", text)
    text = re.sub(r"\bROAD\b", "RD", text)
    return text


def site_address(frame: pd.DataFrame) -> pd.Series:
    """Return the parcel site address."""
    if "FULL_ADDRESS" in frame.columns:
        return frame["FULL_ADDRESS"].fillna("").astype(str)
    street_number = text_series(frame, ["ST_NUM"])
    street_name = text_series(frame, ["ST_NAME"])
    street = street_number + " " + street_name
    return street.str.strip()


def mailing_address(frame: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    """Return full and street-only mailing addresses."""
    street = text_series(frame, ["MAIL_STREET_ADDRESS", "MAIL_ADDRESS"])
    city = text_series(frame, ["MAIL_CITY"])
    state = text_series(frame, ["MAIL_STATE"])
    zip_code = text_series(frame, ["MAIL_ZIP_CODE", "MAIL_ZIP"])
    full = (street + ", " + city + " " + state + " " + zip_code).str.strip()
    full = full.str.replace(r"^,\s*|,\s*$", "", regex=True)
    return full.str.replace(r"\s+", " ", regex=True), street


def different_mail_flags(frame: pd.DataFrame) -> list[str]:
    """Return Yes/No flags for mailing address difference."""
    site = site_address(frame).str.split(",").str[0].apply(normalize_address)
    _, mail_street = mailing_address(frame)
    mail = mail_street.apply(normalize_address)
    return [
        "Yes" if mail_value and site_value != mail_value else "No"
        for site_value, mail_value in zip(site, mail)
    ]


def current_use(frame: pd.DataFrame) -> pd.Series:
    """Return readable current-use text."""
    code = text_series(frame, ["LU", "LUC"])
    description = text_series(frame, ["LU_DESC", "LU_Actual"])
    values = []
    for left, right in zip(code, description):
        values.append(f"{left.strip()} - {right.strip()}".strip(" -"))
    return pd.Series(values, index=frame.index, dtype="object")


def clean_year(value: Any) -> str:
    """Return a clean four-digit year string when available."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return text[:4] if re.fullmatch(r"\d{4}\.0", text) else text


def sale_year(frame: pd.DataFrame) -> pd.Series:
    """Return exposed sale year values when source fields exist."""
    raw = text_series(frame, list(SALE_COLUMNS))
    if raw.eq("").all():
        return raw
    return raw.str.extract(r"(19\d{2}|20\d{2})", expand=False).fillna(raw)


def output_data(frame: pd.DataFrame) -> dict[str, Any]:
    """Build client-facing output columns."""
    parcel_id = first_column(frame, ["MAP_PAR_ID", "PID"])
    zone_col = zoning_column(frame)
    mail_full, _ = mailing_address(frame)
    return {
        "Parcel ID": frame[parcel_id].astype(str),
        "Address": site_address(frame),
        "Verified Subdistrict": frame[zone_col].fillna("").astype(str),
        "Lot Size (sq ft)": number_series(frame, ["LAND_SF", "LOT_SIZE"]),
        "Current Use": current_use(frame),
        "Owner Name": text_series(frame, ["OWNER", "Owner_Name"]),
        "Owner Mailing Address": mail_full,
        "Year Built": year_built(frame),
        "Last Sale Year": sale_year(frame).apply(clean_year),
        "Owner Mailing Different From Property": different_mail_flags(frame),
    }


def parcel_sheet(grouped: pd.DataFrame) -> pd.DataFrame:
    """Build the final parcel list sheet."""
    output = pd.DataFrame(output_data(grouped), columns=CLIENT_COLUMNS)
    output = output.sort_values(
        ["Lot Size (sq ft)", "Parcel ID"],
        ascending=[False, True],
    )
    return output.head(CFG.max_rows).reset_index(drop=True)


def year_built(frame: pd.DataFrame) -> pd.Series:
    """Return cleaned year-built values."""
    return text_series(frame, ["YR_BUILT", "YEAR_BUILT"]).apply(clean_year)


def methodology_source_rows() -> list[tuple[str, Any]]:
    """Return source rows for methodology notes."""
    return [
        (
            "Objective",
            "Verify Dorchester parcels in official Boston S-zones.",
        ),
        ("Neighborhood source", CFG.neighborhood_url),
        ("Zoning source", CFG.zoning_url),
        ("Parcel source", CFG.parcel_url),
        ("Zoning viewer", CFG.viewer_url),
    ]


def methodology_count_rows(counts: dict[str, Any]) -> list[tuple[str, Any]]:
    """Return count rows for methodology notes."""
    return [
        ("Neighborhood features loaded", counts["neighborhoods"]),
        ("Zoning geometries loaded", counts["zoning"]),
        ("Numeric S-zone values", counts["szone_values"]),
        ("Dorchester-clipped S-zone polygons", counts["clipped_zones"]),
        ("Use-filtered parcel candidates", counts["use_filtered"]),
        ("Strict final candidate rows", counts["final_rows"]),
        ("Row-count note", counts["row_count_note"]),
    ]


def methodology_rule_rows() -> list[tuple[str, Any]]:
    """Return rule rows for methodology notes."""
    return [
        (
            "Use rule",
            "Kept residential land, single-family, and two-family only.",
        ),
        (
            "Exclusion rule",
            "Excluded commercial, apartment, R3/R4, mixed-use, parking, etc.",
        ),
        (
            "Last sale year note",
            "FY2026 parcel service exposes no sale date/year field.",
        ),
    ]


def methodology_sheet(counts: dict[str, Any]) -> pd.DataFrame:
    """Build the required methodology notes sheet."""
    rows = (
        methodology_source_rows()
        + methodology_count_rows(counts)
        + methodology_rule_rows()
    )
    return pd.DataFrame(rows, columns=["Step", "Detail"])


def row_count_note(row_count: int) -> str:
    """Return the methodology note for short candidate counts."""
    if 30 <= row_count <= 40:
        return "Final count falls inside the requested 30-40 candidate range."
    return (
        f"Only {row_count} parcels met all stated GIS and use criteria; "
        "no non-target parcels were added to inflate the count."
    )


def style_sheet(sheet: Any) -> None:
    """Apply simple standard spreadsheet formatting."""
    sheet.freeze_panes = "A2"
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for column in sheet.columns:
        width = max(len(str(cell.value or "")) for cell in column)
        letter = column[0].column_letter
        sheet.column_dimensions[letter].width = min(width + 2, 48)
    sheet.auto_filter.ref = sheet.dimensions


def write_workbook(parcels: pd.DataFrame, methodology: pd.DataFrame) -> None:
    """Write the two-sheet client deliverable workbook."""
    with pd.ExcelWriter(CFG.out_file, engine="openpyxl") as writer:
        parcels.to_excel(writer, sheet_name="Parcels", index=False)
        methodology.to_excel(writer, sheet_name="Methodology", index=False)
        for sheet in writer.book.worksheets:
            style_sheet(sheet)


def create_session() -> requests.Session:
    """Create the HTTP session for Boston GIS requests."""
    session = requests.Session()
    session.headers.update({"User-Agent": "Boston-SZone-Stage1/3.0"})
    return session


def pipeline_counts(
    neighborhoods: gpd.GeoDataFrame,
    zoning: gpd.GeoDataFrame,
    zones: gpd.GeoDataFrame,
    mask: gpd.GeoDataFrame,
    candidates: gpd.GeoDataFrame,
    final_rows: int,
) -> dict[str, Any]:
    """Build methodology count values."""
    zone_values = zones["Zoning_Subdistrict"].astype(str).unique()
    values = ", ".join(sorted(zone_values))
    return {
        "neighborhoods": len(neighborhoods),
        "zoning": len(zoning),
        "szone_values": values,
        "clipped_zones": len(mask),
        "use_filtered": len(candidates),
        "final_rows": final_rows,
        "row_count_note": row_count_note(final_rows),
    }


def load_inputs(
    session: requests.Session,
) -> tuple[
    gpd.GeoDataFrame,
    gpd.GeoDataFrame,
    gpd.GeoDataFrame,
    gpd.GeoDataFrame,
    gpd.GeoDataFrame,
]:
    """Load neighborhood and zoning inputs."""
    print("Fetching BPDA neighborhoods")
    neighborhoods = load_layer(session, CFG.neighborhood_url, "neighborhoods")
    dorchester = dorchester_polygon(neighborhoods)
    print("Fetching BPDA zoning")
    zoning = load_layer(session, CFG.zoning_url, "zoning")
    zones = s_zone_polygons(zoning)
    mask = dorchester_s_zone_mask(zones, dorchester)
    return neighborhoods, dorchester, zoning, zones, mask


def build_parcels(
    session: requests.Session,
    zones: gpd.GeoDataFrame,
    dorchester: gpd.GeoDataFrame,
    mask: gpd.GeoDataFrame,
) -> tuple[pd.DataFrame, gpd.GeoDataFrame]:
    """Build final parcel rows from zoning and parcel inputs."""
    nearby = touching_s_zones(zones, dorchester)
    print("Fetching FY2026 parcel candidates")
    candidates = fetch_candidate_parcels(session, nearby)
    joined = strict_join(candidates, mask)
    parcels = parcel_sheet(aggregate_by_parcel(joined))
    return parcels, candidates


def run_pipeline() -> None:
    """Run the Boston Dorchester S-zone verification workflow."""
    session = create_session()
    neighborhoods, dorchester, zoning, zones, mask = load_inputs(session)
    parcels, candidates = build_parcels(session, zones, dorchester, mask)
    counts = pipeline_counts(
        neighborhoods,
        zoning,
        zones,
        mask,
        candidates,
        len(parcels),
    )
    write_workbook(parcels, methodology_sheet(counts))
    print(f"Done. Wrote {len(parcels)} rows to {CFG.out_file}")


def main() -> None:
    """Handle command-line execution and expected failures."""
    try:
        run_pipeline()
    except PipelineError as error:
        print(f"CRITICAL: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
