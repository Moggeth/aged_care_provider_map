# Australia residential aged care homes by provider

This folder contains a provider-coloured map of Australian residential aged-care homes using the official GEN Aged Care Data service list current as at 30 June 2025.

Source: Department of Health, Disability and Ageing / AIHW GEN, "Aged care service list: 30 June 2025". The downloaded source file is `data/Service-List-2025-Australia_300126.xlsx` and matches the current official Australia XLSX download.

Residential-active verification: rows are restricted to `Care Type == Residential`. The generated verification report compares mapped homes with the Department's `data/star-ratings-quarterly-data-extract-february-2026.xlsx` service-level Star Ratings extract for February 2026.

Generated outputs:

- `output/aged_care_homes_by_provider.html`: interactive browser map with provider and state filters.
- `docs/index.html`: the same interactive map, ready to publish with GitHub Pages.
- `output/aged_care_homes_by_provider.kml`: Google Earth / Google My Maps import file, with styles and provider folders.
- `output/aged_care_homes_by_provider.csv`: Google My Maps-friendly table with a `provider_name` and `provider_color` column.
- `output/aged_care_homes_by_provider.geojson`: GIS/web map point data.
- `output/summary.json`: counts by state, care type, and provider.
- `output/verification_report.csv`: residential-home verification status against the February 2026 Star Ratings extract.
- `output/verification_summary.json`: verification counts and interpretation.

Inclusion rule: rows with `Care Type == Residential`, `Residential Places > 0`, and valid latitude/longitude.

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
