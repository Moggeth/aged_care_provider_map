#!/usr/bin/env python3
import csv
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


def slug(value):
    value = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-").lower()
    return value or "unknown"


def sort_key(value):
    return (str(value).casefold(), str(value))


def normalize_key(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).casefold()).strip()


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


def load_california_homes():
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


def load_homes():
    return load_australian_homes() + load_california_homes()


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
                f"<b>Certified places/beds:</b> {home['residential_places']}<br>"
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
    provider_counts = Counter(home["provider_name"] for home in homes)
    top_providers = provider_counts.most_common(20)
    _ratings, verification, _exact_rating_keys, _service_location_rating_keys = verification_counts(homes)
    country_counts = Counter(home["country"] for home in homes)
    california_count = sum(1 for home in homes if home.get("country") == "United States" and home.get("state") == "CA")
    bay_area_count = sum(1 for home in homes if home.get("metro_area") == "San Francisco Bay Area")
    lats = [home["latitude"] for home in homes]
    lons = [home["longitude"] for home in homes]
    data = [
        {
            **home,
            "provider_color": home["provider_color"],
            "funding_2024_25": round(home["funding_2024_25"] or 0, 2),
        }
        for home in homes
    ]
    html_text = HTML_TEMPLATE.replace("__DATA__", json.dumps(data, ensure_ascii=False))
    html_text = html_text.replace("__PROVIDERS__", json.dumps(sorted(providers, key=sort_key), ensure_ascii=False))
    html_text = html_text.replace("__STATES__", json.dumps(states, ensure_ascii=False))
    html_text = html_text.replace("__REGIONS__", json.dumps(regions, ensure_ascii=False))
    html_text = html_text.replace("__COUNT__", str(len(homes)))
    html_text = html_text.replace("__PROVIDER_COUNT__", str(len(providers)))
    html_text = html_text.replace("__AUSTRALIA_COUNT__", str(country_counts["Australia"]))
    html_text = html_text.replace("__CALIFORNIA_COUNT__", str(california_count))
    html_text = html_text.replace("__BAY_AREA_COUNT__", str(bay_area_count))
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
    countries = Counter(home["country"] for home in homes)
    regions = Counter(source_region(home) for home in homes)
    places = [home["residential_places"] for home in homes]
    path = OUTPUT / "summary.json"
    summary = {
        "sources": [
            "GEN Aged Care Data / Department of Health, Disability and Ageing, Aged care service list: 30 June 2025",
            "CMS Provider Information, Nursing homes including rehab services: April 2026",
        ],
        "source_files": [SOURCE_XLSX.name, CMS_NURSING_HOME_CSV.name],
        "included_care_type": "Australia Residential; California CMS currently active nursing homes",
        "generated_home_count": len(homes),
        "provider_count": len(providers),
        "total_residential_places": sum(places),
        "median_residential_places": statistics.median(places),
        "country_counts": dict(sorted(countries.items())),
        "region_counts": dict(sorted(regions.items())),
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

This folder contains a provider-coloured map of Australian residential aged-care homes and California nursing homes using official government source data.

Australian source: Department of Health, Disability and Ageing / AIHW GEN, "Aged care service list: 30 June 2025". The downloaded source file is `data/{SOURCE_XLSX.name}` and matches the current official Australia XLSX download.

California source: Centers for Medicare & Medicaid Services (CMS), "Provider Information", Nursing homes including rehab services. The downloaded source file is `data/{CMS_NURSING_HOME_CSV.name}`.

Australian residential-active verification: Australian rows are restricted to `Care Type == Residential`. The generated verification report compares mapped Australian homes with the Department's `data/{STAR_RATINGS_XLSX.name}` service-level Star Ratings extract for February 2026. CMS describes the California Provider Information table as currently active nursing homes.

Generated outputs:

- `output/aged_care_homes_by_provider.html`: interactive browser map with provider and state filters.
- `docs/index.html`: the same interactive map, ready to publish with GitHub Pages.
- `output/aged_care_homes_by_provider.kml`: Google Earth / Google My Maps import file, with styles and provider folders.
- `output/aged_care_homes_by_provider.csv`: Google My Maps-friendly table with a `provider_name` and `provider_color` column.
- `output/aged_care_homes_by_provider.geojson`: GIS/web map point data.
- `output/summary.json`: counts by state, care type, and provider.
- `output/verification_report.csv`: residential-home verification status against the February 2026 Star Ratings extract.
- `output/verification_summary.json`: verification counts and interpretation.

Inclusion rule: Australian rows with `Care Type == Residential`, `Residential Places > 0`, and valid latitude/longitude; California CMS rows with `State == CA` and valid latitude/longitude.

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
  <meta name="description" content="Interactive map of Australian residential aged-care homes and California nursing homes by provider using official GEN and CMS data.">
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
	      <div class="meta"><span class="count-pill">__COUNT__ homes</span><span>__PROVIDER_COUNT__ providers</span><span>Official GEN + CMS data</span></div>
    </header>
    <div class="controls">
      <label>Search<input id="search" type="search" placeholder="Service, provider, suburb, address"></label>
      <details id="filtersPanel" class="advanced" open>
	        <summary>Filters and sharing</summary>
	        <div class="advanced-content">
	          <label>Region<select id="region"></select></label>
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
	        <p>This map shows Australian residential aged care homes and California nursing homes from official government sources.</p>
	        <p>Source: <a href="https://www.gen-agedcaredata.gov.au/resources/access-data/2025/october/aged-care-service-list-30-june-2025" target="_blank" rel="noopener">Aged care service list: 30 June 2025</a>. The source page says the files are current as at 30 June 2025 and updated annually.</p>
	        <p>California source: <a href="https://data.cms.gov/provider-data/dataset/4pq5-n9py" target="_blank" rel="noopener">CMS Provider Information</a>. CMS describes this April 2026 dataset as general information on currently active nursing homes, one row per nursing home.</p>
	        <p>Included Australian rows have residential places above zero and valid latitude/longitude. Included California rows are CMS currently active nursing homes with valid latitude/longitude; certified beds are shown where CMS supplies them. This version maps __AUSTRALIA_COUNT__ Australian homes, __CALIFORNIA_COUNT__ California homes, and __BAY_AREA_COUNT__ San Francisco Bay Area homes.</p>
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
	    const topProviders = __TOP_PROVIDERS__;
    const bounds = __BOUNDS__;
	    const regionSelect = document.getElementById('region');
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
	        <div class="popup-row"><b>Places/beds:</b> ${home.residential_places}</div>
	        <div class="popup-row"><b>Address:</b> ${escapeHtml(home.address)}</div>
	        <div class="popup-row"><b>Organisation:</b> ${escapeHtml(home.organisation_type)}</div>
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
	        if (filters.provider && home.provider_name !== filters.provider) continue;
        if (filters.state && home.state !== filters.state) continue;
        if (!matches(home, q)) continue;
	        item.marker.addTo(layer);
	        shownBounds.push([home.latitude, home.longitude]);
	        shown += 1;
	        places += home.residential_places;
	        visibleProviders[home.provider_name] = (visibleProviders[home.provider_name] || 0) + 1;
	      }
	      visibleCount.textContent = `${shown.toLocaleString()} visible homes, ${places.toLocaleString()} places/beds`;
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
	      providerSelect.value = params.get('provider') || '';
      stateSelect.value = params.get('state') || '';
      searchInput.value = params.get('q') || '';
      mapStyleSelect.value = params.get('style') || 'light';
    }

	    function resetFilters() {
	      regionSelect.value = '';
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

	    [regionSelect, providerSelect, stateSelect, mapStyleSelect, searchInput].forEach(el => el.addEventListener('input', applyFilters));
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
    ]
    paths.append(write_readme(paths))
    print(f"Generated {len(homes)} homes across {len(providers)} providers")
    for path in paths:
        size = path.stat().st_size
        print(f"{path.relative_to(ROOT)} ({math.ceil(size / 1024)} KB)")


if __name__ == "__main__":
    main()
