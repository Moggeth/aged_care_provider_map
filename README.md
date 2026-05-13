# Residential care homes by provider

This folder contains a provider-coloured map of Australian residential aged-care homes and California residential care facilities using official government source data.

Australian source: Department of Health, Disability and Ageing / AIHW GEN, "Aged care service list: 30 June 2025". The downloaded source file is `data/Service-List-2025-Australia_300126.xlsx` and matches the current official Australia XLSX download.

California nursing-home source: Centers for Medicare & Medicaid Services (CMS), "Provider Information", Nursing homes including rehab services. The downloaded source file is `data/NH_ProviderInfo_Apr2026.csv`.

California residential-care source: California Department of Social Services, "Community Care Licensing Facilities", Residential Care Facilities for the Elderly. The downloaded source file is `data/CA_RCFE_Community_Care_Licensing_Facilities_20250525.csv` and the source data file date is 05/25/2025. Active rows with `facility_status` of `LICENSED` or `ON PROBATION` are included when they receive a usable U.S. Census Geocoder address match. `CLOSED` and `PENDING` rows are excluded.

Australian residential-active verification: Australian rows are restricted to `Care Type == Residential`. The generated verification report compares mapped Australian homes with the Department's `data/star-ratings-quarterly-data-extract-february-2026.xlsx` service-level Star Ratings extract for February 2026. CMS describes the California Provider Information table as currently active nursing homes. California RCFE rows are validated against the official CDSS CCLD file and geocoded with the U.S. Census Geocoder.

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
