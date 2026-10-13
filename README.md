# fetch-data-googlemaps

Small experiment for collecting nearby places from Google Maps and saving them to an Excel sheet.
It contains a working Flask endpoint that appends Places API results to `data.xlsx`, and an
unfinished CLI script for fetching those results.

## What's here

| File | Status | Description |
|------|--------|-------------|
| `app.py` | works | Flask app with `POST /`. Accepts a Places API (New) style JSON body (`{"places": [...]}`), extracts display name, types, formatted address and website for each place, and appends the rows to `data.xlsx`. |
| `main.py` | incomplete | Prompts for an address and radius, geocodes the address with the Geocoding API, then calls the legacy Nearby Search endpoint. See limitations below. |
| `data.json` | sample | Example Places API (New) response with 2 restaurants, usable as a request body for `app.py`. |
| `data.xlsx` | sample | Output produced from `data.json`. |

## Getting started

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app.py                  # Flask dev server on http://127.0.0.1:5000
```

Append the sample places to `data.xlsx`:

```bash
curl -X POST http://127.0.0.1:5000/ \
  -H 'Content-Type: application/json' \
  -d @data.json
# -> Data appended and saved successfully
```

Each place in the request must have `types`, `formattedAddress`, `websiteUri` and
`displayName.text`, otherwise the endpoint returns an `Error: ...` message.

For `main.py`, provide a Google Maps Platform key through the environment (never hardcode it):

```bash
export GOOGLE_MAPS_API_KEY=...
python main.py
```

## Known limitations (`main.py`)

- The nearby-search URL interpolates Python's builtin `type` (`&type=<class 'type'>`) instead of a
  place type.
- The legacy Nearby Search API returns `results`, but `save_to_excel()` expects the Places API (New)
  `places` format, and `save_to_excel()` is never called.
- `googlemaps` is imported but not used.

Porting the request to the Places API (New) `places:searchNearby` endpoint would make `main.py` and
`app.py` share one data format.

## License

Apache License 2.0 - see [LICENSE](LICENSE).
