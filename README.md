# Attendance Backend

Flask service that fills an Excel template (`Site_Split_Template.xlsx`) with attendance metrics and returns a generated workbook (`Daily_Attendance_Auto.xlsx`). CORS is enabled.

## Quick Start (Windows, bash)

```bash
# 1) Create and activate a virtual environment
python -m venv .venv
source .venv/Scripts/activate

# 2) Install dependencies
python -m pip install --upgrade pip
pip install -r requirements.txt

# 3) Ensure the Excel template exists in project root
#    Required file: Site_Split_Template.xlsx

# 4) Run the app
export FLASK_DEBUG=1
python app.py
```

- App runs on `http://localhost:10000` by default.
- You can change the port by setting `PORT`, e.g. `export PORT=8080`.

## Endpoints

- `GET /` — health check; returns `Backend running!`.
- `POST /api/generate-dashboard` — generates an Excel workbook and returns it as a download. Provide a JSON payload with metric data.
- `GET /api/generate-dashboard/test` — convenience endpoint using built-in dummy data (manual testing).

### Example payload (partial)
```json
{
  "RegularHC": {
    "inbound_amzn": 10,
    "inbound_temp": 5,
    "da_amzn": 8,
    "da_temp": 3,
    "icqa_amzn": 4,
    "icqa_temp": 1,
    "crets_amzn": 9,
    "crets_temp": 6
  }
}
```

### Try it (curl)

- GET (dummy data):
```bash
curl -L "http://localhost:10000/api/generate-dashboard" -o Daily_Attendance_Auto.xlsx
```

- POST (your data):
```bash
curl -X POST "http://localhost:10000/api/generate-dashboard" \
  -H "Content-Type: application/json" \
  -d '{
    "RegularHC": {
      "inbound_amzn": 10,
      "inbound_temp": 5,
      "da_amzn": 8,
      "da_temp": 3,
      "icqa_amzn": 4,
      "icqa_temp": 1,
      "crets_amzn": 9,
      "crets_temp": 6
    }
  }' \
  -o Daily_Attendance_Auto.xlsx
```

## Excel Logic

- Metrics written to specific rows (e.g., `RegularHC` → row 6).
- Department columns: B,C,D,E,G,H,I,J.
- Totals:
  - `F` (SDC total) = inbound + DA
  - `K` (IXD total) = CRETs
  - `L` (Grand total) = SDC + ICQA + IXD
- `MET*` rows are forced to 0 by design.

## Deployment (Render)

1) Push your repo to GitHub with `app.py`, `requirements.txt`, and `Site_Split_Template.xlsx` in the root.
2) In Render, create a new Web Service:
   - Environment: `Python 3`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app --timeout 120 --workers 2`
3) Render provides `PORT`; the app reads it automatically.
4) Redeploy after changes by pushing to GitHub.

Note: A `Procfile` is included for other platforms; Render primarily uses the Start Command above.

## Troubleshooting

- Missing template error: Place `Site_Split_Template.xlsx` in the project root.
- Invalid JSON on POST: Endpoint returns 400 with a helpful error message.
- CORS: Currently open; restrict origins if needed for production.
- 404 Not Found: Verify the path `/api/generate-dashboard` and that your Render service URL is correct.
- 405 Method Not Allowed: Use POST for `/api/generate-dashboard`; GET is only for `/api/generate-dashboard/test`.
