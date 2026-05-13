#!/usr/bin/env python3
import csv
import hashlib
import html
import json
import math
import re
import statistics
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
SOURCE_XLSX = ROOT / "data" / "Service-List-2025-Australia_300126.xlsx"
STAR_RATINGS_XLSX = ROOT / "data" / "star-ratings-quarterly-data-extract-february-2026.xlsx"
CMS_NURSING_HOME_CSV = ROOT / "data" / "NH_ProviderInfo_Apr2026.csv"
CMS_NURSING_HOME_METADATA = ROOT / "data" / "cms_provider_information_metadata.json"
CA_RCFE_CSV = ROOT / "data" / "CA_RCFE_Community_Care_Licensing_Facilities_20250525.csv"
CA_RCFE_DICTIONARY_CSV = ROOT / "data" / "ca_rcfe_data_dictionary.csv"
CA_RCFE_GEOCODES_CSV = ROOT / "data" / "ca_rcfe_geocodes_census_20260513.csv"
OUTPUT = ROOT / "output"

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
HEADERS_ROW = 3
SF_BAY_AREA_COUNTIES = {
    "Alameda",
    "Contra Costa",
    "Marin",
    "Napa",
    "San Francisco",
    "San Mateo",
    "Santa Clara",
    "Solano",
    "Sonoma",
}
CALIFORNIA_BOUNDS = {
    "min_latitude": 32.0,
    "max_latitude": 42.5,
    "min_longitude": -125.0,
    "max_longitude": -114.0,
}
CA_RCFE_ACTIVE_STATUSES = {"LICENSED", "ON PROBATION"}
CA_RCFE_SOURCE_URL = (
    "https://catalog.data.gov/dataset/community-care-licensing-facilities"
)
CA_RCFE_DOWNLOAD_URL = (
    "https://data.chhs.ca.gov/dataset/46ffcbdf-4874-4cc1-92c2-fb715e3ad014/"
    "resource/744d1583-f9eb-45b6-b0f8-b9a9dab936a6/download/tmpacjmwy9v.csv"
)
CENSUS_GEOCODER_URL = "https://geocoding.geo.census.gov/geocoder/locations/addressbatch"


def column_index(cell_ref):
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def read_xlsx(path):
    with ZipFile(path) as zf:
        strings = []
        shared = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        for item in shared.findall("a:si", NS):
            strings.append("".join(t.text or "" for t in item.findall(".//a:t", NS)))

        headers = None
        for _event, row in ET.iterparse(zf.open("xl/worksheets/sheet1.xml"), events=("end",)):
            if not row.tag.endswith("}row"):
                continue

            row_num = int(row.attrib.get("r", "0"))
            values = {}
            for cell in row.findall("a:c", NS):
                ref = cell.attrib.get("r", "")
                value_node = cell.find("a:v", NS)
                value = "" if value_node is None else value_node.text or ""
                if cell.attrib.get("t") == "s" and value:
                    value = strings[int(value)]
                values[column_index(ref)] = value

            if row_num == HEADERS_ROW:
                headers = [values.get(i, "") for i in range(max(values) + 1)]
            elif row_num > HEADERS_ROW and headers and values:
                yield {headers[i]: clean(values.get(i, "")) for i in range(len(headers))}

            row.clear()


def read_xlsx_sheet(path, sheet_name, headers_row=1):
    with ZipFile(path) as zf:
        strings = []
        shared = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        for item in shared.findall("a:si", NS):
            strings.append("".join(t.text or "" for t in item.findall(".//a:t", NS)))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        target = None
        rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        for sheet in workbook.findall("a:sheets/a:sheet", NS):
            if sheet.attrib.get("name") == sheet_name:
                target = relmap[sheet.attrib[rel_ns]]
                break
        if not target:
            raise ValueError(f"Sheet not found: {sheet_name}")

        headers = None
        for _event, row in ET.iterparse(zf.open(f"xl/{target}"), events=("end",)):
            if not row.tag.endswith("}row"):
                continue

            row_num = int(row.attrib.get("r", "0"))
            values = {}
            for cell in row.findall("a:c", NS):
                ref = cell.attrib.get("r", "")
                value_node = cell.find("a:v", NS)
                value = "" if value_node is None else value_node.text or ""
                if cell.attrib.get("t") == "s" and value:
                    value = strings[int(value)]
                values[column_index(ref)] = value

            if row_num == headers_row:
                headers = [values.get(i, "") for i in range(max(values) + 1)]
            elif row_num > headers_row and headers and values:
                yield {headers[i]: clean(values.get(i, "")) for i in range(len(headers))}

            row.clear()


def clean(value):
    return str(value).strip() if value is not None else ""


def to_float(value):
    value = clean(value)
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def to_int(value):
    num = to_float(value)
    return int(num) if num is not None else 0


def file_sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def slug(value):
    value = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-").lower()
    return value or "unknown"


def sort_key(value):
    return (str(value).casefold(), str(value))


def normalize_key(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).casefold()).strip()


def normalize_county(value):
    return " ".join(part.capitalize() for part in clean(value).split())


def parse_rcfe_file_date(value):
    value = clean(value)
    if len(value) != 8:
        return value
    return f"{value[0:2]}/{value[2:4]}/{value[4:8]}"


def care_category(home):
    if home.get("country") == "Australia":
        return "Australian residential aged care"
    if home.get("source_type") == "ca_rcfe_ccrc":
        return "California RCFE-CCRC"
    if home.get("source_type") == "ca_rcfe":
        return "California RCFE"
    if home.get("source_type") == "ca_cms_nursing_home":
        return "California nursing home"
    return home.get("care_type") or "Other"


def california_fit(home):
    if home.get("country") != "United States" or home.get("state") != "CA":
        return ""
    if home.get("source_type") == "ca_rcfe":
        return "High-fit elder residential"
    if home.get("source_type") in {"ca_rcfe_ccrc", "ca_cms_nursing_home"}:
        return "Hybrid facility"
    return "Unclassified California facility"


def provider_color(provider, providers):
    idx = providers[provider]
    total = max(len(providers), 1)
    hue = (idx * 137.508) % 360
    saturation = 68
    lightness = 45 + (idx % 3) * 7
    return hsl_to_hex(hue / 360, saturation / 100, lightness / 100)


def hsl_to_hex(h, s, l):
    def hue_to_rgb(p, q, t):
        if t < 0:
            t += 1
        if t > 1:
            t -= 1
        if t < 1 / 6:
            return p + (q - p) * 6 * t
        if t < 1 / 2:
            return q
        if t < 2 / 3:
            return p + (q - p) * (2 / 3 - t) * 6
        return p

    if s == 0:
        r = g = b = l
    else:
        q = l * (1 + s) if l < 0.5 else l + s - l * s
        p = 2 * l - q
        r = hue_to_rgb(p, q, h + 1 / 3)
        g = hue_to_rgb(p, q, h)
        b = hue_to_rgb(p, q, h - 1 / 3)
    return "#{:02x}{:02x}{:02x}".format(round(r * 255), round(g * 255), round(b * 255))


def kml_color(hex_color):
    hex_color = hex_color.lstrip("#")
    rr, gg, bb = hex_color[0:2], hex_color[2:4], hex_color[4:6]
    return f"ff{bb}{gg}{rr}"


def source_region(home):
    if home.get("country") == "United States":
        return "San Francisco Bay Area" if home.get("metro_area") == "San Francisco Bay Area" else "California"
    return "Australia"


def load_australian_homes():
    homes = []
    for row in read_xlsx(SOURCE_XLSX):
        if row.get("Care Type") != "Residential":
            continue
        places = to_int(row.get("Residential Places"))
        lat = to_float(row.get("Latitude"))
        lon = to_float(row.get("Longitude"))
        if places <= 0 or lat is None or lon is None:
            continue
        provider = row.get("Provider Name") or "Unknown provider"
        home = {
            "source_region": "Australia",
            "country": "Australia",
            "service_name": row.get("Service Name"),
            "provider_name": provider,
            "care_type": row.get("Care Type"),
            "residential_places": places,
            "address": ", ".join(
                part
                for part in [
                    row.get("Physical Address"),
                    row.get("Physical Suburb"),
                    row.get("Physical State"),
                    row.get("Physical Post Code"),
                ]
                if part
            ),
            "state": row.get("Physical State"),
            "suburb": row.get("Physical Suburb"),
            "postcode": row.get("Physical Post Code"),
            "organisation_type": row.get("Organisation Type"),
            "remoteness": row.get("ABS Remoteness"),
            "acpr": row.get("2018 Aged Care Planning Region (ACPR)"),
            "lga": row.get("2023 LGA Name"),
            "county": "",
            "metro_area": "",
            "source_dataset": "GEN Aged Care Service List: 30 June 2025",
            "source_status": "Official service list current as at 30 June 2025",
            "source_type": "au_residential",
            "source_identifier": "",
            "license_status": "",
            "geocode_status": "source coordinates",
            "geocode_match_type": "",
            "cms_certification_number": "",
            "legal_business_name": "",
            "chain_name": "",
            "overall_rating": "",
            "average_residents_per_day": "",
            "latitude": lat,
            "longitude": lon,
            "funding_2024_25": to_float(row.get("2024-25 Australian Government Funding")),
        }
        homes.append(home)
    return homes


def load_california_nursing_homes():
    if not CMS_NURSING_HOME_CSV.exists():
        return []

    homes = []
    with CMS_NURSING_HOME_CSV.open(newline="", encoding="utf-8-sig") as handle:
        for row in csv.DictReader(handle):
            if row.get("State") != "CA":
                continue
            beds = to_int(row.get("Number of Certified Beds"))
            lat = to_float(row.get("Latitude"))
            lon = to_float(row.get("Longitude"))
            if lat is None or lon is None:
                continue
            chain = row.get("Chain Name")
            legal = row.get("Legal Business Name")
            if legal == "Legal Business Name Not Available":
                legal = ""
            provider = chain or legal or row.get("Provider Name") or "Unknown provider"
            county = row.get("County/Parish")
            metro = "San Francisco Bay Area" if county in SF_BAY_AREA_COUNTIES else ""
            address = ", ".join(
                part
                for part in [
                    row.get("Provider Address"),
                    row.get("City/Town"),
                    row.get("State"),
                    row.get("ZIP Code"),
                    "USA",
                ]
                if part
            )
            home = {
                "source_region": metro or "California",
                "country": "United States",
                "service_name": row.get("Provider Name"),
                "provider_name": provider,
                "care_type": "Nursing home",
                "residential_places": beds,
                "address": address,
                "state": row.get("State"),
                "suburb": row.get("City/Town"),
                "postcode": row.get("ZIP Code"),
                "organisation_type": row.get("Ownership Type"),
                "remoteness": "Urban" if row.get("Urban") == "Y" else "Non-urban",
                "acpr": "",
                "lga": county,
                "county": county,
                "metro_area": metro,
                "source_dataset": "CMS Provider Information: April 2026",
                "source_status": "CMS describes this dataset as currently active nursing homes",
                "source_type": "ca_cms_nursing_home",
                "source_identifier": row.get("CMS Certification Number (CCN)"),
                "license_status": "currently active in CMS Provider Information",
                "geocode_status": "source coordinates",
                "geocode_match_type": "",
                "cms_certification_number": row.get("CMS Certification Number (CCN)"),
                "legal_business_name": legal,
                "chain_name": chain,
                "overall_rating": row.get("Overall Rating"),
                "average_residents_per_day": row.get("Average Number of Residents per Day"),
                "latitude": lat,
                "longitude": lon,
                "funding_2024_25": "",
            }
            homes.append(home)
    return homes


def load_rcfe_rows():
    if not CA_RCFE_CSV.exists():
        return []
    with CA_RCFE_CSV.open(newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


def load_rcfe_geocodes():
    if not CA_RCFE_GEOCODES_CSV.exists():
        return {}
    geocodes = {}
    with CA_RCFE_GEOCODES_CSV.open(newline="", encoding="utf-8") as handle:
        for row in csv.reader(handle):
            if len(row) < 3:
                continue
            lon = lat = None
            if row[2] == "Match" and len(row) >= 6:
                parts = row[5].split(",")
                if len(parts) == 2:
                    lon = to_float(parts[0])
                    lat = to_float(parts[1])
            geocodes[row[0]] = {
                "input_address": row[1],
                "match_status": row[2],
                "match_type": row[3] if len(row) > 3 else "",
                "matched_address": row[4] if len(row) > 4 else "",
                "longitude": lon,
                "latitude": lat,
            }
    return geocodes


def load_california_rcfe_homes():
    rows = load_rcfe_rows()
    if not rows:
        return []
    geocodes = load_rcfe_geocodes()
    homes = []
    for row in rows:
        if row.get("facility_status") not in CA_RCFE_ACTIVE_STATUSES:
            continue
        capacity = to_int(row.get("facility_capacity"))
        if capacity <= 0:
            continue
        geocode = geocodes.get(row.get("facility_number"), {})
        if geocode.get("match_status") != "Match":
            continue
        lat = geocode.get("latitude")
        lon = geocode.get("longitude")
        if lat is None or lon is None:
            continue
        if not (
            CALIFORNIA_BOUNDS["min_latitude"] <= lat <= CALIFORNIA_BOUNDS["max_latitude"]
            and CALIFORNIA_BOUNDS["min_longitude"] <= lon <= CALIFORNIA_BOUNDS["max_longitude"]
        ):
            continue
        county = normalize_county(row.get("county_name"))
        metro = "San Francisco Bay Area" if county in SF_BAY_AREA_COUNTIES else ""
        facility_type = row.get("facility_type")
        is_ccrc = facility_type == "RCFE-CONTINUING CARE RETIREMENT COMMUNITY"
        provider = row.get("licensee") or row.get("facility_name") or "Unknown provider"
        city = row.get("facility_city")
        state = row.get("facility_state")
        zip_code = row.get("facility_zip")
        address = ", ".join(part for part in [row.get("facility_address"), city, state, zip_code, "USA"] if part)
        homes.append(
            {
                "source_region": metro or "California",
                "country": "United States",
                "service_name": row.get("facility_name"),
                "provider_name": provider,
                "care_type": "Continuing care retirement community" if is_ccrc else "Residential care facility for the elderly",
                "care_category": "California RCFE-CCRC" if is_ccrc else "California RCFE",
                "residential_places": capacity,
                "address": address,
                "state": state,
                "suburb": city,
                "postcode": zip_code,
                "organisation_type": facility_type,
                "remoteness": "",
                "acpr": "",
                "lga": county,
                "county": county,
                "metro_area": metro,
                "source_dataset": "California DSS Community Care Licensing: Residential Care Facilities for the Elderly",
                "source_status": (
                    "Official CCLD RCFE file dated "
                    f"{parse_rcfe_file_date(row.get('file_date'))}; status {row.get('facility_status')}"
                ),
                "source_type": "ca_rcfe_ccrc" if is_ccrc else "ca_rcfe",
                "source_identifier": row.get("facility_number"),
                "license_status": row.get("facility_status"),
                "geocode_status": "matched by U.S. Census Geocoder",
                "geocode_match_type": geocode.get("match_type", ""),
                "cms_certification_number": "",
                "legal_business_name": row.get("licensee"),
                "chain_name": "",
                "overall_rating": "",
                "average_residents_per_day": "",
                "latitude": lat,
                "longitude": lon,
                "funding_2024_25": "",
            }
        )
    return homes


def load_cms_provider_rows():
    if not CMS_NURSING_HOME_CSV.exists():
        return []
    with CMS_NURSING_HOME_CSV.open(newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


def load_cms_metadata():
    if not CMS_NURSING_HOME_METADATA.exists():
        return {}
    return json.loads(CMS_NURSING_HOME_METADATA.read_text(encoding="utf-8"))


def load_homes():
    return load_australian_homes() + load_california_nursing_homes() + load_california_rcfe_homes()


def load_star_ratings():
    if not STAR_RATINGS_XLSX.exists():
        return []
    return list(read_xlsx_sheet(STAR_RATINGS_XLSX, "Star Ratings"))


def verification_keys_for_home(home):
    exact = (
        normalize_key(home["service_name"]),
        normalize_key(home["provider_name"]),
        normalize_key(home["suburb"]),
        normalize_key(home["state"]),
    )
    service_location = (
        normalize_key(home["service_name"]),
        normalize_key(home["suburb"]),
        normalize_key(home["state"]),
    )
    return exact, service_location


def verification_keys_for_rating(row):
    exact = (
        normalize_key(row.get("Service Name")),
        normalize_key(row.get("Provider Name")),
        normalize_key(row.get("Service Suburb")),
        normalize_key(row.get("State/Territory")),
    )
    service_location = (
        normalize_key(row.get("Service Name")),
        normalize_key(row.get("Service Suburb")),
        normalize_key(row.get("State/Territory")),
    )
    return exact, service_location


def verification_status_for_home(home, exact_rating_keys, service_location_rating_keys):
    exact_key, service_location_key = verification_keys_for_home(home)
    if exact_key in exact_rating_keys:
        return "confirmed_in_feb_2026_star_ratings", "Exact service, provider, suburb and state match."
    if service_location_key in service_location_rating_keys:
        return "service_location_match_provider_changed", "Service name, suburb and state match; provider name differs in Star Ratings."
    return (
        "not_matched_in_feb_2026_star_ratings",
        "Included in official 30 June 2025 service list, but not matched in February 2026 Star Ratings by service/suburb/state.",
    )


def verification_counts(homes):
    homes = [home for home in homes if home.get("country") == "Australia"]
    ratings = load_star_ratings()
    exact_rating_keys = {verification_keys_for_rating(row)[0] for row in ratings}
    service_location_rating_keys = {verification_keys_for_rating(row)[1] for row in ratings}
    counts = Counter()
    for home in homes:
        status, _notes = verification_status_for_home(home, exact_rating_keys, service_location_rating_keys)
        counts[status] += 1
    return ratings, counts, exact_rating_keys, service_location_rating_keys


def write_verification_report(homes):
    path = OUTPUT / "verification_report.csv"
    summary_path = OUTPUT / "verification_summary.json"
    homes = [home for home in homes if home.get("country") == "Australia"]
    ratings, counts, exact_rating_keys, service_location_rating_keys = verification_counts(homes)
    fields = [
        "country",
        "service_name",
        "provider_name",
        "state",
        "suburb",
        "address",
        "residential_places",
        "verification_status",
        "official_service_list_current_at",
        "star_ratings_extract",
        "notes",
    ]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for home in homes:
            status, notes = verification_status_for_home(home, exact_rating_keys, service_location_rating_keys)
            writer.writerow(
                {
                    "service_name": home["service_name"],
                    "country": home["country"],
                    "provider_name": home["provider_name"],
                    "state": home["state"],
                    "suburb": home["suburb"],
                    "address": home["address"],
                    "residential_places": home["residential_places"],
                    "verification_status": status,
                    "official_service_list_current_at": "30 June 2025",
                    "star_ratings_extract": "February 2026",
                    "notes": notes,
                }
            )

    summary = {
        "source": "GEN Aged Care Service List: 30 June 2025",
        "source_verified_against_official_download": True,
        "included_care_type": "Residential",
        "mapped_residential_homes": len(homes),
        "star_ratings_source": "Star Ratings quarterly data extract – February 2026",
        "star_ratings_rows": len(ratings),
        "verification_counts": dict(counts),
        "interpretation": (
            "All mapped homes are official Australian Government subsidised residential aged care services "
            "current as at 30 June 2025. Star Ratings matching is an additional February 2026 activity signal; "
            "absence from Star Ratings is flagged for review rather than treated as closure."
        ),
    }
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    return path, summary_path


def california_coordinate_in_bounds(row):
    lat = to_float(row.get("Latitude"))
    lon = to_float(row.get("Longitude"))
    if lat is None or lon is None:
        return False
    return (
        CALIFORNIA_BOUNDS["min_latitude"] <= lat <= CALIFORNIA_BOUNDS["max_latitude"]
        and CALIFORNIA_BOUNDS["min_longitude"] <= lon <= CALIFORNIA_BOUNDS["max_longitude"]
    )


def write_source_validation_report(homes):
    report_path = OUTPUT / "source_validation_report.csv"
    summary_path = OUTPUT / "source_validation_summary.json"

    australian_homes = [home for home in homes if home.get("country") == "Australia"]
    california_homes = [home for home in homes if home.get("country") == "United States" and home.get("state") == "CA"]
    ratings, au_counts, exact_rating_keys, service_location_rating_keys = verification_counts(australian_homes)
    cms_rows = load_cms_provider_rows()
    cms_metadata = load_cms_metadata()
    cms_ca_rows = [row for row in cms_rows if row.get("State") == "CA"]
    ca_cms_homes = [home for home in california_homes if home.get("source_type") == "ca_cms_nursing_home"]
    ca_rcfe_homes = [home for home in california_homes if home.get("source_type") in {"ca_rcfe", "ca_rcfe_ccrc"}]
    cms_ca_ccns = {row.get("CMS Certification Number (CCN)") for row in cms_ca_rows}
    mapped_ca_ccns = {home.get("cms_certification_number") for home in ca_cms_homes}
    sf_city_rows = [row for row in cms_ca_rows if row.get("City/Town", "").casefold() == "san francisco"]
    sf_county_rows = [row for row in cms_ca_rows if row.get("County/Parish") == "San Francisco"]
    bay_area_rows = [row for row in cms_ca_rows if row.get("County/Parish") in SF_BAY_AREA_COUNTIES]
    rcfe_rows = load_rcfe_rows()
    rcfe_geocodes = load_rcfe_geocodes()
    rcfe_active_rows = [row for row in rcfe_rows if row.get("facility_status") in CA_RCFE_ACTIVE_STATUSES]
    rcfe_mapped_numbers = {home.get("source_identifier") for home in ca_rcfe_homes}
    rcfe_unmapped_active = [
        row
        for row in rcfe_active_rows
        if row.get("facility_number") not in rcfe_mapped_numbers
    ]
    rcfe_sf_active_rows = [
        row for row in rcfe_active_rows if normalize_county(row.get("county_name")) == "San Francisco"
    ]
    rcfe_bay_area_active_rows = [
        row for row in rcfe_active_rows if normalize_county(row.get("county_name")) in SF_BAY_AREA_COUNTIES
    ]

    fields = [
        "country",
        "region",
        "service_name",
        "provider_name",
        "state",
        "city_or_suburb",
        "county_or_lga",
        "source_identifier",
        "validation_status",
        "official_source",
        "source_date_or_release",
        "coordinates_status",
        "notes",
    ]
    validation_counts = Counter()
    with report_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for home in homes:
            if home.get("country") == "Australia":
                status, notes = verification_status_for_home(home, exact_rating_keys, service_location_rating_keys)
                official_source = "GEN Aged Care Service List"
                source_date = "Current as at 30 June 2025"
                identifier = ""
            elif home.get("country") == "United States" and home.get("state") == "CA":
                if home.get("source_type") == "ca_cms_nursing_home":
                    status = "confirmed_currently_active_in_cms_provider_information_apr_2026"
                    notes = "CCN is present in CMS Provider Information dataset 4pq5-n9py."
                    if not home.get("residential_places"):
                        notes += " CMS did not supply Number of Certified Beds for this row."
                    official_source = "CMS Provider Information"
                    source_date = "Released 2026-04-29; next update 2026-05-27"
                    identifier = home.get("cms_certification_number")
                elif home.get("source_type") in {"ca_rcfe", "ca_rcfe_ccrc"}:
                    status = "confirmed_active_in_cdss_ccld_rcfe_file_2025_05_25"
                    notes = (
                        f"Facility number is present in CDSS CCLD RCFE file with status {home.get('license_status')}; "
                        f"coordinates are {home.get('geocode_match_type', '').lower()} matches from the U.S. Census Geocoder."
                    )
                    official_source = "California DSS Community Care Licensing Facilities - Residential Care Facilities for the Elderly"
                    source_date = "File date 05/25/2025"
                    identifier = home.get("source_identifier")
                else:
                    status = "unsupported_california_source"
                    notes = "California row does not identify a configured source type."
                    official_source = home.get("source_dataset")
                    source_date = ""
                    identifier = home.get("source_identifier")
            else:
                status = "unsupported_region"
                notes = "Region is outside the configured validation sources."
                official_source = home.get("source_dataset")
                source_date = ""
                identifier = ""

            validation_counts[status] += 1
            writer.writerow(
                {
                    "country": home.get("country"),
                    "region": source_region(home),
                    "service_name": home.get("service_name"),
                    "provider_name": home.get("provider_name"),
                    "state": home.get("state"),
                    "city_or_suburb": home.get("suburb"),
                    "county_or_lga": home.get("county") or home.get("lga"),
                    "source_identifier": identifier,
                    "validation_status": status,
                    "official_source": official_source,
                    "source_date_or_release": source_date,
                    "coordinates_status": "valid latitude/longitude",
                    "notes": notes,
                }
            )

    ca_duplicate_ccns = [
        ccn for ccn, count in Counter(row.get("CMS Certification Number (CCN)") for row in cms_ca_rows).items() if count > 1
    ]
    california_summary = {
        "mapped_california_rows": len(california_homes),
        "prospecting_fit_counts": dict(
            sorted(Counter(california_fit(home) for home in california_homes if california_fit(home)).items())
        ),
        "prospecting_fit_interpretation": (
            "High-fit elder residential is a practical filter for standard CDSS Residential Care Facility for the "
            "Elderly rows. Hybrid facility covers CDSS RCFE-CCRC and CMS nursing-home/skilled-nursing rows."
        ),
        "bay_area_counties": sorted(SF_BAY_AREA_COUNTIES),
        "cms_nursing_homes": {
            "official_source": "CMS Provider Information",
            "dataset_id": cms_metadata.get("identifier", "4pq5-n9py"),
            "dataset_title": cms_metadata.get("title", "Provider Information"),
            "dataset_description": cms_metadata.get("description", ""),
            "landing_page": cms_metadata.get("landingPage", "https://data.cms.gov/provider-data/dataset/4pq5-n9py"),
            "download_url": (cms_metadata.get("distribution") or [{}])[0].get("downloadURL", ""),
            "released": cms_metadata.get("released"),
            "modified": cms_metadata.get("modified"),
            "next_update_date": cms_metadata.get("nextUpdateDate"),
            "source_file": CMS_NURSING_HOME_CSV.name,
            "source_file_sha256": file_sha256(CMS_NURSING_HOME_CSV) if CMS_NURSING_HOME_CSV.exists() else "",
            "all_california_source_rows_mapped": cms_ca_ccns == mapped_ca_ccns,
            "cms_total_rows": len(cms_rows),
            "cms_california_rows": len(cms_ca_rows),
            "mapped_california_rows": len(ca_cms_homes),
            "unique_california_ccns": len(cms_ca_ccns),
            "duplicate_california_ccns": ca_duplicate_ccns,
            "california_rows_missing_coordinates": sum(
                1 for row in cms_ca_rows if not row.get("Latitude") or not row.get("Longitude")
            ),
            "california_rows_outside_coordinate_bounds": sum(
                1
                for row in cms_ca_rows
                if row.get("Latitude") and row.get("Longitude") and not california_coordinate_in_bounds(row)
            ),
            "california_rows_missing_certified_beds": sum(1 for row in cms_ca_rows if not row.get("Number of Certified Beds")),
            "san_francisco_city_rows": len(sf_city_rows),
            "san_francisco_county_rows": len(sf_county_rows),
            "san_francisco_bay_area_rows": len(bay_area_rows),
        },
        "rcfe": {
            "official_source": "California DSS Community Care Licensing Facilities - Residential Care Facilities for the Elderly",
            "landing_page": CA_RCFE_SOURCE_URL,
            "download_url": CA_RCFE_DOWNLOAD_URL,
            "source_file": CA_RCFE_CSV.name,
            "source_file_sha256": file_sha256(CA_RCFE_CSV) if CA_RCFE_CSV.exists() else "",
            "data_dictionary_file": CA_RCFE_DICTIONARY_CSV.name,
            "geocoder": "U.S. Census Geocoder address batch",
            "geocoder_url": CENSUS_GEOCODER_URL,
            "geocode_file": CA_RCFE_GEOCODES_CSV.name,
            "geocode_file_sha256": file_sha256(CA_RCFE_GEOCODES_CSV) if CA_RCFE_GEOCODES_CSV.exists() else "",
            "file_date_values": sorted({parse_rcfe_file_date(row.get("file_date")) for row in rcfe_rows if row.get("file_date")}),
            "total_rows": len(rcfe_rows),
            "status_counts": dict(Counter(row.get("facility_status") for row in rcfe_rows)),
            "included_statuses": sorted(CA_RCFE_ACTIVE_STATUSES),
            "active_rows": len(rcfe_active_rows),
            "mapped_active_rows": len(ca_rcfe_homes),
            "unmapped_active_rows": len(rcfe_unmapped_active),
            "geocode_status_counts": dict(Counter(geocode.get("match_status", "missing") for geocode in rcfe_geocodes.values())),
            "geocode_match_type_counts": dict(
                Counter(geocode.get("match_type", "") for geocode in rcfe_geocodes.values() if geocode.get("match_status") == "Match")
            ),
            "mapped_rcfe_rows": sum(1 for home in ca_rcfe_homes if home.get("source_type") == "ca_rcfe"),
            "mapped_rcfe_ccrc_rows": sum(1 for home in ca_rcfe_homes if home.get("source_type") == "ca_rcfe_ccrc"),
            "san_francisco_active_rows": len(rcfe_sf_active_rows),
            "san_francisco_mapped_rows": sum(1 for home in ca_rcfe_homes if home.get("county") == "San Francisco"),
            "san_francisco_active_capacity": sum(to_int(row.get("facility_capacity")) for row in rcfe_sf_active_rows),
            "san_francisco_mapped_capacity": sum(
                home.get("residential_places") or 0 for home in ca_rcfe_homes if home.get("county") == "San Francisco"
            ),
            "san_francisco_bay_area_active_rows": len(rcfe_bay_area_active_rows),
            "san_francisco_bay_area_mapped_rows": sum(
                1 for home in ca_rcfe_homes if home.get("metro_area") == "San Francisco Bay Area"
            ),
        },
    }
    australia_summary = {
        "official_source": "GEN Aged Care Service List",
        "source_file": SOURCE_XLSX.name,
        "source_file_sha256": file_sha256(SOURCE_XLSX) if SOURCE_XLSX.exists() else "",
        "mapped_australian_rows": len(australian_homes),
        "star_ratings_source_file": STAR_RATINGS_XLSX.name,
        "star_ratings_rows": len(ratings),
        "star_ratings_validation_counts": dict(au_counts),
    }
    summary = {
        "official_sources_only": True,
        "generated_home_count": len(homes),
        "validation_counts": dict(validation_counts),
        "australia": australia_summary,
        "california": california_summary,
        "interpretation": (
            "California nursing-home rows are validated against the official CMS Provider Information file, which CMS "
            "describes as currently active nursing homes. California RCFE rows are active LICENSED or ON PROBATION "
            "facilities from the official CDSS Community Care Licensing RCFE file and are plotted only when the U.S. "
            "Census Geocoder returned a usable address match. Australian rows are validated against the official GEN "
            "service list and cross-checked against the February 2026 Star Ratings extract."
        ),
    }
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    return report_path, summary_path


def write_csv(homes):
    path = OUTPUT / "aged_care_homes_by_provider.csv"
    fields = [
        "source_region",
        "country",
        "service_name",
        "provider_name",
        "provider_color",
        "care_type",
        "residential_places",
        "address",
        "state",
        "suburb",
        "postcode",
        "organisation_type",
        "remoteness",
        "acpr",
        "lga",
        "county",
        "metro_area",
        "source_dataset",
        "source_status",
        "source_type",
        "source_identifier",
        "license_status",
        "geocode_status",
        "geocode_match_type",
        "cms_certification_number",
        "legal_business_name",
        "chain_name",
        "overall_rating",
        "average_residents_per_day",
        "latitude",
        "longitude",
        "funding_2024_25",
    ]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for home in homes:
            writer.writerow({key: home.get(key, "") for key in fields})
    return path


def write_geojson(homes):
    path = OUTPUT / "aged_care_homes_by_provider.geojson"
    features = []
    for home in homes:
        props = {k: v for k, v in home.items() if k not in ("latitude", "longitude")}
        props["care_category"] = care_category(home)
        props["california_fit"] = california_fit(home)
        features.append(
            {
                "type": "Feature",
                "geometry": {"type": "Point", "coordinates": [home["longitude"], home["latitude"]]},
                "properties": props,
            }
        )
    data = {"type": "FeatureCollection", "features": features}
    path.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    return path


def write_kml(homes, providers):
    path = OUTPUT / "aged_care_homes_by_provider.kml"
    chunks = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2">',
        "<Document>",
        "<name>Residential care homes by provider</name>",
    ]
    for provider in providers:
        color = kml_color(provider_color(provider, providers))
        sid = f"provider-{slug(provider)}"
        chunks.extend(
            [
                f'<Style id="{sid}">',
                "<IconStyle>",
                f"<color>{color}</color>",
                "<scale>0.8</scale>",
                '<Icon><href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href></Icon>',
                "</IconStyle>",
                "</Style>",
            ]
        )

    for provider, provider_homes in sorted(group_by_provider(homes).items(), key=lambda item: sort_key(item[0])):
        chunks.append("<Folder>")
        chunks.append(f"<name>{xml(provider)} ({len(provider_homes)})</name>")
        for home in provider_homes:
            desc = (
                f"<b>Provider:</b> {html.escape(home['provider_name'])}<br>"
                f"<b>Country:</b> {html.escape(home['country'])}<br>"
                f"<b>Region:</b> {html.escape(source_region(home))}<br>"
                f"<b>Care category:</b> {html.escape(care_category(home))}<br>"
                f"<b>Places/capacity/beds:</b> {home['residential_places']}<br>"
                f"<b>Address:</b> {html.escape(home['address'])}<br>"
                f"<b>Organisation type:</b> {html.escape(home['organisation_type'])}<br>"
                f"<b>Source:</b> {html.escape(home['source_dataset'])}"
            )
            chunks.extend(
                [
                    "<Placemark>",
                    f"<name>{xml(home['service_name'])}</name>",
                    f"<styleUrl>#provider-{slug(provider)}</styleUrl>",
                    f"<description><![CDATA[{desc}]]></description>",
                    f"<Point><coordinates>{home['longitude']},{home['latitude']},0</coordinates></Point>",
                    "</Placemark>",
                ]
            )
        chunks.append("</Folder>")

    chunks.extend(["</Document>", "</kml>"])
    path.write_text("\n".join(chunks), encoding="utf-8")
    return path


def xml(value):
    return html.escape(str(value), quote=True)


def group_by_provider(homes):
    grouped = {}
    for home in homes:
        grouped.setdefault(home["provider_name"], []).append(home)
    return grouped


def write_html(homes, providers):
    path = OUTPUT / "aged_care_homes_by_provider.html"
    docs_path = ROOT / "docs" / "index.html"
    states = sorted({home["state"] for home in homes if home["state"]})
    regions = ["Australia", "California", "San Francisco Bay Area"]
    care_categories = sorted({care_category(home) for home in homes}, key=sort_key)
    california_fits = ["High-fit elder residential", "Hybrid facility"]
    provider_counts = Counter(home["provider_name"] for home in homes)
    top_providers = provider_counts.most_common(20)
    _ratings, verification, _exact_rating_keys, _service_location_rating_keys = verification_counts(homes)
    country_counts = Counter(home["country"] for home in homes)
    california_count = sum(1 for home in homes if home.get("country") == "United States" and home.get("state") == "CA")
    california_rcfe_count = sum(1 for home in homes if home.get("source_type") in {"ca_rcfe", "ca_rcfe_ccrc"})
    california_high_fit_count = sum(1 for home in homes if california_fit(home) == "High-fit elder residential")
    california_hybrid_count = sum(1 for home in homes if california_fit(home) == "Hybrid facility")
    california_nursing_home_count = sum(1 for home in homes if home.get("source_type") == "ca_cms_nursing_home")
    bay_area_count = sum(1 for home in homes if home.get("metro_area") == "San Francisco Bay Area")
    sf_rcfe_count = sum(
        1
        for home in homes
        if home.get("source_type") in {"ca_rcfe", "ca_rcfe_ccrc"} and home.get("county") == "San Francisco"
    )
    lats = [home["latitude"] for home in homes]
    lons = [home["longitude"] for home in homes]
    data = [
        {
            **home,
            "provider_color": home["provider_color"],
            "funding_2024_25": round(home["funding_2024_25"] or 0, 2),
            "care_category": care_category(home),
            "california_fit": california_fit(home),
        }
        for home in homes
    ]
    html_text = HTML_TEMPLATE.replace("__DATA__", json.dumps(data, ensure_ascii=False))
    html_text = html_text.replace("__PROVIDERS__", json.dumps(sorted(providers, key=sort_key), ensure_ascii=False))
    html_text = html_text.replace("__STATES__", json.dumps(states, ensure_ascii=False))
    html_text = html_text.replace("__REGIONS__", json.dumps(regions, ensure_ascii=False))
    html_text = html_text.replace("__CARE_CATEGORIES__", json.dumps(care_categories, ensure_ascii=False))
    html_text = html_text.replace("__CALIFORNIA_FITS__", json.dumps(california_fits, ensure_ascii=False))
    html_text = html_text.replace("__COUNT__", str(len(homes)))
    html_text = html_text.replace("__PROVIDER_COUNT__", str(len(providers)))
    html_text = html_text.replace("__AUSTRALIA_COUNT__", str(country_counts["Australia"]))
    html_text = html_text.replace("__CALIFORNIA_COUNT__", str(california_count))
    html_text = html_text.replace("__CALIFORNIA_RCFE_COUNT__", str(california_rcfe_count))
    html_text = html_text.replace("__CALIFORNIA_HIGH_FIT_COUNT__", str(california_high_fit_count))
    html_text = html_text.replace("__CALIFORNIA_HYBRID_COUNT__", str(california_hybrid_count))
    html_text = html_text.replace("__CALIFORNIA_NURSING_HOME_COUNT__", str(california_nursing_home_count))
    html_text = html_text.replace("__BAY_AREA_COUNT__", str(bay_area_count))
    html_text = html_text.replace("__SF_RCFE_COUNT__", str(sf_rcfe_count))
    html_text = html_text.replace("__VERIFIED_EXACT__", str(verification["confirmed_in_feb_2026_star_ratings"]))
    html_text = html_text.replace("__VERIFIED_PROVIDER_CHANGED__", str(verification["service_location_match_provider_changed"]))
    html_text = html_text.replace("__VERIFIED_UNMATCHED__", str(verification["not_matched_in_feb_2026_star_ratings"]))
    html_text = html_text.replace("__TOP_PROVIDERS__", json.dumps(top_providers, ensure_ascii=False))
    html_text = html_text.replace("__BOUNDS__", json.dumps([[min(lats), min(lons)], [max(lats), max(lons)]]))
    path.write_text(html_text, encoding="utf-8")
    docs_path.parent.mkdir(exist_ok=True)
    docs_path.write_text(html_text, encoding="utf-8")
    (docs_path.parent / ".nojekyll").write_text("", encoding="utf-8")
    return path, docs_path


def write_summary(homes, providers):
    states = Counter(home["state"] for home in homes)
    care_types = Counter(home["care_type"] for home in homes)
    care_categories = Counter(care_category(home) for home in homes)
    california_fits = Counter(california_fit(home) for home in homes if california_fit(home))
    countries = Counter(home["country"] for home in homes)
    regions = Counter(source_region(home) for home in homes)
    places = [home["residential_places"] for home in homes]
    path = OUTPUT / "summary.json"
    summary = {
        "sources": [
            "GEN Aged Care Data / Department of Health, Disability and Ageing, Aged care service list: 30 June 2025",
            "CMS Provider Information, Nursing homes including rehab services: April 2026",
            "California DSS Community Care Licensing Facilities, Residential Care Facilities for the Elderly: file date 05/25/2025",
        ],
        "source_files": [SOURCE_XLSX.name, CMS_NURSING_HOME_CSV.name, CA_RCFE_CSV.name],
        "included_care_type": (
            "Australia Residential; California CMS currently active nursing homes; "
            "California CDSS licensed/on-probation Residential Care Facilities for the Elderly"
        ),
        "generated_home_count": len(homes),
        "provider_count": len(providers),
        "total_residential_places": sum(places),
        "median_residential_places": statistics.median(places),
        "country_counts": dict(sorted(countries.items())),
        "region_counts": dict(sorted(regions.items())),
        "care_category_counts": dict(sorted(care_categories.items())),
        "california_fit_counts": dict(sorted(california_fits.items())),
        "care_type_counts": dict(sorted(care_types.items())),
        "state_counts": dict(sorted(states.items())),
        "top_20_providers": Counter(home["provider_name"] for home in homes).most_common(20),
    }
    path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def write_readme(paths):
    path = ROOT / "README.md"
    path.write_text(
        f"""# Residential care homes by provider

This folder contains a provider-coloured map of Australian residential aged-care homes and California residential care facilities using official government source data.

Australian source: Department of Health, Disability and Ageing / AIHW GEN, "Aged care service list: 30 June 2025". The downloaded source file is `data/{SOURCE_XLSX.name}` and matches the current official Australia XLSX download.

California nursing-home source: Centers for Medicare & Medicaid Services (CMS), "Provider Information", Nursing homes including rehab services. The downloaded source file is `data/{CMS_NURSING_HOME_CSV.name}`.

California residential-care source: California Department of Social Services, "Community Care Licensing Facilities", Residential Care Facilities for the Elderly. The downloaded source file is `data/{CA_RCFE_CSV.name}` and the source data file date is 05/25/2025. Active rows with `facility_status` of `LICENSED` or `ON PROBATION` are included when they receive a usable U.S. Census Geocoder address match. `CLOSED` and `PENDING` rows are excluded.

Australian residential-active verification: Australian rows are restricted to `Care Type == Residential`. The generated verification report compares mapped Australian homes with the Department's `data/{STAR_RATINGS_XLSX.name}` service-level Star Ratings extract for February 2026. CMS describes the California Provider Information table as currently active nursing homes. California RCFE rows are validated against the official CDSS CCLD file and geocoded with the U.S. Census Geocoder.

California prospecting fit: the map adds a practical, non-official filter that labels standard CDSS `RESIDENTIAL CARE ELDERLY` rows as `High-fit elder residential`, and labels CDSS `RCFE-CONTINUING CARE RETIREMENT COMMUNITY` plus CMS nursing homes as `Hybrid facility`.

Generated outputs:

- `output/aged_care_homes_by_provider.html`: interactive browser map with provider and state filters.
- `docs/index.html`: the same interactive map, ready to publish with GitHub Pages.
- `output/aged_care_homes_by_provider.kml`: Google Earth / Google My Maps import file, with styles and provider folders.
- `output/aged_care_homes_by_provider.csv`: Google My Maps-friendly table with a `provider_name` and `provider_color` column.
- `output/aged_care_homes_by_provider.geojson`: GIS/web map point data.
- `output/summary.json`: counts by state, care type, and provider.
- `output/verification_report.csv`: residential-home verification status against the February 2026 Star Ratings extract.
- `output/verification_summary.json`: verification counts and interpretation.
- `output/source_validation_report.csv`: row-level validation evidence for every mapped home.
- `output/source_validation_summary.json`: source metadata, row counts, checksums, and California/San Francisco validation totals.

Inclusion rule: Australian rows with `Care Type == Residential`, `Residential Places > 0`, and valid latitude/longitude; California CMS rows with `State == CA` and valid latitude/longitude; California RCFE rows with `facility_status` in `LICENSED` or `ON PROBATION`, `facility_capacity > 0`, and a usable Census Geocoder latitude/longitude.

Use with Google My Maps:

1. Create a new map at https://www.google.com/mymaps.
2. Import `output/aged_care_homes_by_provider.csv`.
3. Choose `latitude` and `longitude` as the location columns.
4. Choose `service_name` as the title column.
5. Use "Style by data column" and select `provider_name`. The file also includes `provider_color` if you want to manually align colours with the HTML/KML outputs.

Regenerate with:

```bash
python3 scripts/build_map.py
```

Professional HTTPS hosting:

- CloudFront URL: https://dgv72coqns6yt.cloudfront.net/
- Origin bucket: `andromeda-aged-care-provider-map-prod`
- CloudFront distribution: `E2Y6OESD3IEA02`
- The S3 origin is private, public bucket access is blocked, CloudFront redirects HTTP to HTTPS, and the AWS managed security headers policy is attached.

Deploy updates with:

```bash
python3 scripts/build_map.py
scripts/deploy_aws_static.sh
```

Publish online with GitHub Pages:

1. Create a GitHub repository and push this folder to it.
2. In the repository, open Settings -> Pages.
3. Set "Build and deployment" to "Deploy from a branch".
4. Choose your main branch and the `/docs` folder, then save.
5. After GitHub finishes publishing, share the Pages URL. Re-run `python3 scripts/build_map.py`, commit, and push whenever you improve the map.

For a no-GitHub option, upload `output/aged_care_homes_by_provider.html` to Netlify Drop or another static host.
""",
        encoding="utf-8",
    )
    return path


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Residential Care Homes by Provider</title>
  <meta name="description" content="Interactive map of Australian residential aged-care homes and California residential care facilities by provider using official GEN, CMS and CDSS data.">
  <meta property="og:title" content="Residential care homes by provider">
  <meta property="og:description" content="Search, filter and share an interactive provider-coloured map of official Australian and California residential care data.">
  <meta property="og:type" content="website">
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
  <style>
    html, body { height: 100%; margin: 0; font-family: Arial, Helvetica, sans-serif; color: #17212b; }
    #map { height: 100%; width: 100%; background: #eef2f4; }
    .panel { position: absolute; z-index: 1000; top: 16px; left: 16px; width: min(430px, calc(100vw - 32px)); background: rgba(255, 255, 255, .97); border: 1px solid #cbd5df; box-shadow: 0 10px 32px rgba(15, 23, 42, .18); border-radius: 8px; }
    .panel header { padding: 14px 16px 10px; border-bottom: 1px solid #e1e7ee; }
    h1 { font-size: 18px; line-height: 1.25; margin: 0 0 8px; letter-spacing: 0; }
    .meta { display: flex; flex-wrap: wrap; gap: 8px; font-size: 12px; color: #52606d; }
    .controls { display: grid; gap: 10px; padding: 12px 16px; }
    label { display: grid; gap: 4px; font-size: 12px; color: #334e68; }
    select, input, button { font: inherit; }
    select, input { width: 100%; box-sizing: border-box; border: 1px solid #bcccdc; border-radius: 6px; padding: 8px 10px; font-size: 14px; background: #fff; color: #17212b; }
    select:focus, input:focus, button:focus-visible { outline: 3px solid rgba(45, 125, 210, .24); outline-offset: 1px; }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .advanced summary { display: none; }
    .advanced-content { display: grid; gap: 10px; }
    .actions { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    button { border: 1px solid #9fb3c8; border-radius: 6px; background: #f8fafc; color: #17212b; cursor: pointer; font-size: 14px; min-height: 36px; padding: 7px 10px; }
    button:hover { background: #eef4fa; }
    .stats { padding: 10px 16px 14px; border-top: 1px solid #e1e7ee; font-size: 13px; color: #334e68; }
    .legend { max-height: 150px; overflow: auto; display: grid; gap: 5px; margin-top: 8px; }
    .legend-item { display: grid; grid-template-columns: 14px 1fr auto; align-items: center; gap: 7px; }
    .about { border-top: 1px solid #e1e7ee; font-size: 13px; color: #334e68; }
    .about summary { cursor: pointer; padding: 10px 16px; color: #17212b; font-weight: 700; list-style-position: inside; }
    .about-content { display: grid; gap: 8px; padding: 0 16px 14px; line-height: 1.4; }
    .about-content p { margin: 0; }
    .about-content a { color: #1d5fa7; }
    .swatch { width: 10px; height: 10px; border-radius: 50%; border: 1px solid rgba(0,0,0,.25); }
    .leaflet-container { font-family: Arial, Helvetica, sans-serif; }
    .leaflet-control-attribution { max-width: min(260px, calc(100vw - 72px)); }
    .leaflet-popup-content { min-width: 230px; }
    .popup-title { font-weight: 700; margin-bottom: 6px; }
    .popup-row { margin: 3px 0; }
    .dot { width: 9px; height: 9px; border: 1px solid rgba(0,0,0,.35); border-radius: 999px; display: inline-block; margin-right: 6px; vertical-align: middle; }
    .count-pill { background: #e6f0f8; border: 1px solid #bcccdc; border-radius: 999px; color: #243b53; padding: 2px 7px; }
    @media (max-width: 720px) {
      .panel { top: auto; bottom: 8px; left: 8px; width: calc(100vw - 16px); max-height: min(52dvh, 390px); overflow: auto; }
      .panel header { padding: 10px 12px 8px; }
      h1 { font-size: 15px; margin-bottom: 6px; }
      .meta { gap: 6px; font-size: 11px; }
      .controls { gap: 8px; padding: 10px 12px; }
      select, input { min-height: 38px; font-size: 16px; }
      .row { grid-template-columns: 1fr; }
      .actions { grid-template-columns: 1fr; }
      .advanced summary { display: flex; align-items: center; justify-content: space-between; min-height: 36px; border: 1px solid #bcccdc; border-radius: 6px; background: #f8fafc; color: #17212b; cursor: pointer; padding: 0 10px; font-size: 14px; list-style: none; }
      .advanced summary::-webkit-details-marker { display: none; }
      .advanced summary::after { content: "Show"; color: #52606d; font-size: 12px; }
      .advanced[open] summary { margin-bottom: 8px; }
      .advanced[open] summary::after { content: "Hide"; }
      .stats { padding: 8px 12px 10px; }
      .about summary { padding: 9px 12px; }
      .about-content { padding: 0 12px 12px; }
      .legend { display: none; }
      .leaflet-top.leaflet-right { top: 8px; }
    }
  </style>
</head>
<body>
  <div id="map"></div>
  <section class="panel">
    <header>
	      <h1>Residential care homes</h1>
	      <div class="meta"><span class="count-pill">__COUNT__ homes</span><span>__PROVIDER_COUNT__ providers</span><span>Official GEN + CMS + CDSS data</span></div>
    </header>
    <div class="controls">
      <label>Search<input id="search" type="search" placeholder="Service, provider, suburb, address"></label>
      <details id="filtersPanel" class="advanced" open>
	        <summary>Filters and sharing</summary>
	        <div class="advanced-content">
	          <label>Region<select id="region"></select></label>
	          <label>California fit<select id="californiaFit"></select></label>
	          <label>Care category<select id="careCategory"></select></label>
	          <label>Provider<select id="provider"></select></label>
	          <label>State / territory<select id="state"></select></label>
          <label>Map style<select id="mapStyle"></select></label>
          <div class="actions">
            <button id="shareButton" type="button">Copy share link</button>
            <button id="resetButton" type="button">Reset filters</button>
          </div>
        </div>
      </details>
    </div>
    <div class="stats">
      <div id="visibleCount"></div>
      <div class="legend" id="legend"></div>
    </div>
    <details class="about">
      <summary>About this data</summary>
      <div class="about-content">
	        <p>This map shows Australian residential aged care homes and California residential care facilities from official government sources.</p>
	        <p>Source: <a href="https://www.gen-agedcaredata.gov.au/resources/access-data/2025/october/aged-care-service-list-30-june-2025" target="_blank" rel="noopener">Aged care service list: 30 June 2025</a>. The source page says the files are current as at 30 June 2025 and updated annually.</p>
	        <p>California nursing-home source: <a href="https://data.cms.gov/provider-data/dataset/4pq5-n9py" target="_blank" rel="noopener">CMS Provider Information</a>. CMS describes this April 2026 dataset as general information on currently active nursing homes, one row per nursing home.</p>
	        <p>California residential-care source: <a href="https://catalog.data.gov/dataset/community-care-licensing-facilities" target="_blank" rel="noopener">California DSS Community Care Licensing Facilities</a>, Residential Care Facilities for the Elderly. Included RCFE rows have status LICENSED or ON PROBATION in the official file dated 05/25/2025; CLOSED and PENDING rows are excluded. RCFE coordinates are matched from the facility address using the <a href="https://geocoding.geo.census.gov/geocoder/" target="_blank" rel="noopener">U.S. Census Geocoder</a>.</p>
	        <p>Included Australian rows have residential places above zero and valid latitude/longitude. Included California rows include __CALIFORNIA_RCFE_COUNT__ active RCFEs/RCFE-CCRCS and __CALIFORNIA_NURSING_HOME_COUNT__ currently active CMS nursing homes. This version maps __AUSTRALIA_COUNT__ Australian homes, __CALIFORNIA_COUNT__ California homes, __BAY_AREA_COUNT__ San Francisco Bay Area homes, and __SF_RCFE_COUNT__ San Francisco RCFEs.</p>
	        <p>The California fit filter is a practical prospecting classification, not an official license category. High-fit elder residential means a standard CDSS Residential Care Facility for the Elderly (__CALIFORNIA_HIGH_FIT_COUNT__ mapped facilities). Hybrid facility means an RCFE-CCRC or CMS nursing home/skilled nursing facility (__CALIFORNIA_HYBRID_COUNT__ mapped facilities); these may house older long-term residents, but can also include continuing-care, rehabilitation, skilled-nursing, hospital-based, or clinically complex populations.</p>
	        <p>Cross-check: <a href="https://www.health.gov.au/resources/publications/star-ratings-quarterly-data-extract-february-2026?language=en" target="_blank" rel="noopener">February 2026 Star Ratings service-level extract</a>. __VERIFIED_EXACT__ homes matched exactly; __VERIFIED_PROVIDER_CHANGED__ matched by service/suburb/state with provider-name differences; __VERIFIED_UNMATCHED__ were not matched and should be manually reviewed.</p>
	        <p>Absence from the Star Ratings extract is not treated as proof that a home has closed.</p>
      </div>
    </details>
  </section>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <script>
	    const homes = __DATA__;
	    const providers = __PROVIDERS__;
	    const states = __STATES__;
	    const regions = __REGIONS__;
	    const careCategories = __CARE_CATEGORIES__;
	    const californiaFits = __CALIFORNIA_FITS__;
	    const topProviders = __TOP_PROVIDERS__;
    const bounds = __BOUNDS__;
	    const regionSelect = document.getElementById('region');
	    const californiaFitSelect = document.getElementById('californiaFit');
	    const careCategorySelect = document.getElementById('careCategory');
	    const providerSelect = document.getElementById('provider');
    const stateSelect = document.getElementById('state');
    const mapStyleSelect = document.getElementById('mapStyle');
    const searchInput = document.getElementById('search');
    const filtersPanel = document.getElementById('filtersPanel');
    const shareButton = document.getElementById('shareButton');
    const resetButton = document.getElementById('resetButton');
    const visibleCount = document.getElementById('visibleCount');
    const legend = document.getElementById('legend');

	    const map = L.map('map', {
	      preferCanvas: true,
	      zoomControl: false,
	      zoomAnimation: false,
	      fadeAnimation: false,
	      markerZoomAnimation: false
	    });
    map.attributionControl.setPosition('topleft');
    map.attributionControl.setPrefix(false);
    L.control.zoom({ position: 'topright' }).addTo(map);
    const baseLayers = {
      light: L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
        maxZoom: 20,
        attribution: '&copy; OpenStreetMap contributors &copy; CARTO'
      }),
      detail: L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19,
        attribution: '&copy; OpenStreetMap contributors'
      })
    };
	    let activeBaseLayer = baseLayers.light.addTo(map);
	    let lastRegion = null;
	    map.fitBounds(bounds, { paddingTopLeft: [470, 32], paddingBottomRight: [32, 32], animate: false });

    function option(select, value, text) {
      const item = document.createElement('option');
      item.value = value;
      item.textContent = text;
      select.appendChild(item);
    }

	    option(regionSelect, '', 'All regions');
	    regions.forEach(region => option(regionSelect, region, region));
	    option(californiaFitSelect, '', 'All California fit types');
	    californiaFits.forEach(fit => option(californiaFitSelect, fit, fit));
	    option(careCategorySelect, '', 'All care categories');
	    careCategories.forEach(category => option(careCategorySelect, category, category));
	    option(providerSelect, '', 'All providers');
    providers.forEach(provider => option(providerSelect, provider, provider));
    option(stateSelect, '', 'All states');
    states.forEach(state => option(stateSelect, state, state));
    option(mapStyleSelect, 'light', 'Light map');
    option(mapStyleSelect, 'detail', 'Detailed streets');

    const layer = L.layerGroup().addTo(map);
    const markers = homes.map(home => {
      const marker = L.circleMarker([home.latitude, home.longitude], {
        radius: Math.max(5, Math.min(12, Math.sqrt(home.residential_places) / 1.35)),
        color: '#ffffff',
        weight: 1.4,
        fillColor: home.provider_color,
        fillOpacity: 0.9,
        opacity: 1
      });
	      marker.bindPopup(`
	        <div class="popup-title"><span class="dot" style="background:${home.provider_color}"></span>${escapeHtml(home.service_name)}</div>
	        <div class="popup-row"><b>Provider:</b> ${escapeHtml(home.provider_name)}</div>
	        <div class="popup-row"><b>Region:</b> ${escapeHtml(regionLabel(home))}</div>
	        <div class="popup-row"><b>Care category:</b> ${escapeHtml(home.care_category)}</div>
	        ${home.california_fit ? `<div class="popup-row"><b>California fit:</b> ${escapeHtml(home.california_fit)}</div>` : ''}
	        <div class="popup-row"><b>Places/capacity/beds:</b> ${home.residential_places}</div>
	        <div class="popup-row"><b>Address:</b> ${escapeHtml(home.address)}</div>
	        <div class="popup-row"><b>Organisation:</b> ${escapeHtml(home.organisation_type)}</div>
	        ${home.license_status ? `<div class="popup-row"><b>Status:</b> ${escapeHtml(home.license_status)}</div>` : ''}
	        ${home.overall_rating ? `<div class="popup-row"><b>CMS overall rating:</b> ${escapeHtml(home.overall_rating)} / 5</div>` : ''}
	        <div class="popup-row"><b>Source:</b> ${escapeHtml(home.source_dataset)}</div>
	      `);
	      return { home, marker };
	    });

    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
    }

	    function matches(home, q) {
	      if (!q) return true;
	      return [home.service_name, home.provider_name, home.suburb, home.address, home.lga, home.county, home.chain_name, home.legal_business_name].join(' ').toLowerCase().includes(q);
	    }

	    function matchesRegion(home, region) {
	      if (!region) return true;
	      if (region === 'Australia') return home.country === 'Australia';
	      if (region === 'California') return home.country === 'United States' && home.state === 'CA';
	      if (region === 'San Francisco Bay Area') return home.metro_area === 'San Francisco Bay Area';
	      return true;
	    }

	    function regionLabel(home) {
	      if (home.metro_area) return home.metro_area;
	      if (home.country === 'United States') return 'California';
	      return 'Australia';
	    }

    function selectedFilters() {
	      return {
	        region: regionSelect.value,
	        californiaFit: californiaFitSelect.value,
	        careCategory: careCategorySelect.value,
	        provider: providerSelect.value,
        state: stateSelect.value,
        q: searchInput.value.trim(),
        style: mapStyleSelect.value
      };
    }

    function setBaseLayer(style) {
      const nextLayer = baseLayers[style] || baseLayers.light;
      if (activeBaseLayer === nextLayer) return;
      map.removeLayer(activeBaseLayer);
      activeBaseLayer = nextLayer.addTo(map);
      activeBaseLayer.bringToBack();
    }

    function updateUrl() {
      const filters = selectedFilters();
	      const params = new URLSearchParams();
	      if (filters.region) params.set('region', filters.region);
	      if (filters.californiaFit) params.set('caFit', filters.californiaFit);
	      if (filters.careCategory) params.set('care', filters.careCategory);
	      if (filters.provider) params.set('provider', filters.provider);
      if (filters.state) params.set('state', filters.state);
      if (filters.q) params.set('q', filters.q);
      if (filters.style !== 'light') params.set('style', filters.style);
      const nextUrl = `${location.pathname}${params.toString() ? '?' + params.toString() : ''}${location.hash}`;
      history.replaceState(null, '', nextUrl);
    }

    function applyFilters() {
      const filters = selectedFilters();
      const q = filters.q.toLowerCase();
      setBaseLayer(filters.style);
      layer.clearLayers();
	      let shown = 0;
	      let places = 0;
	      const shownBounds = [];
	      const visibleProviders = {};
	      for (const item of markers) {
	        const home = item.home;
	        if (!matchesRegion(home, filters.region)) continue;
	        if (filters.californiaFit && home.california_fit !== filters.californiaFit) continue;
	        if (filters.careCategory && home.care_category !== filters.careCategory) continue;
	        if (filters.provider && home.provider_name !== filters.provider) continue;
        if (filters.state && home.state !== filters.state) continue;
        if (!matches(home, q)) continue;
	        item.marker.addTo(layer);
	        shownBounds.push([home.latitude, home.longitude]);
	        shown += 1;
	        places += home.residential_places;
	        visibleProviders[home.provider_name] = (visibleProviders[home.provider_name] || 0) + 1;
	      }
	      visibleCount.textContent = `${shown.toLocaleString()} visible homes, ${places.toLocaleString()} places/capacity/beds`;
	      if (filters.region !== lastRegion && shownBounds.length) {
	        map.fitBounds(shownBounds, { paddingTopLeft: [470, 32], paddingBottomRight: [32, 32], animate: false });
	        lastRegion = filters.region;
	      }
	      renderLegend(visibleProviders);
	      updateUrl();
	    }

	    function renderLegend(providerCounts = null) {
	      legend.innerHTML = '';
	      const providersToShow = providerCounts
	        ? Object.entries(providerCounts).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0])).slice(0, 20)
	        : topProviders;
	      providersToShow.forEach(([provider, count]) => {
	        const home = homes.find(item => item.provider_name === provider);
	        if (!home) return;
	        const row = document.createElement('div');
        row.className = 'legend-item';
        row.innerHTML = `<span class="swatch" style="background:${home.provider_color}"></span><span>${escapeHtml(provider)}</span><span>${count}</span>`;
        legend.appendChild(row);
      });
    }

	    function applyUrlParams() {
	      const params = new URLSearchParams(location.search);
	      regionSelect.value = params.get('region') || '';
	      californiaFitSelect.value = params.get('caFit') || '';
	      careCategorySelect.value = params.get('care') || '';
	      providerSelect.value = params.get('provider') || '';
      stateSelect.value = params.get('state') || '';
      searchInput.value = params.get('q') || '';
      mapStyleSelect.value = params.get('style') || 'light';
    }

	    function resetFilters() {
	      regionSelect.value = '';
	      californiaFitSelect.value = '';
	      careCategorySelect.value = '';
	      providerSelect.value = '';
      stateSelect.value = '';
      searchInput.value = '';
      mapStyleSelect.value = 'light';
      applyFilters();
    }

    async function copyShareLink() {
      applyFilters();
      try {
        await navigator.clipboard.writeText(location.href);
        shareButton.textContent = 'Link copied';
      } catch (_error) {
        shareButton.textContent = 'Copy from address bar';
      }
      window.setTimeout(() => { shareButton.textContent = 'Copy share link'; }, 1800);
    }

	    [regionSelect, californiaFitSelect, careCategorySelect, providerSelect, stateSelect, mapStyleSelect, searchInput].forEach(el => el.addEventListener('input', applyFilters));
    shareButton.addEventListener('click', copyShareLink);
    resetButton.addEventListener('click', resetFilters);
    applyUrlParams();
    if (matchMedia('(max-width: 720px)').matches) {
      filtersPanel.removeAttribute('open');
    }
	    applyFilters();
  </script>
</body>
</html>
"""


def main():
    OUTPUT.mkdir(exist_ok=True)
    homes = load_homes()
    provider_names = sorted({home["provider_name"] for home in homes}, key=sort_key)
    providers = {provider: index for index, provider in enumerate(provider_names)}
    for home in homes:
        home["provider_color"] = provider_color(home["provider_name"], providers)

    paths = [
        write_csv(homes),
        write_geojson(homes),
        write_kml(homes, providers),
        *write_html(homes, providers),
        write_summary(homes, providers),
        *write_verification_report(homes),
        *write_source_validation_report(homes),
    ]
    paths.append(write_readme(paths))
    print(f"Generated {len(homes)} homes across {len(providers)} providers")
    for path in paths:
        size = path.stat().st_size
        print(f"{path.relative_to(ROOT)} ({math.ceil(size / 1024)} KB)")


if __name__ == "__main__":
    main()
