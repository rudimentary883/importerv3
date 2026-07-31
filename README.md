# Tabbycat API Importer

Web service that imports tournament data into Tabbycat via REST API or generates CSVs.

## Render Free Tier
- 512MB RAM safe
- Single worker, 2 threads
- 120s timeout

## Two Modes
1. Download CSVs - for manual import
2. Direct API Import - connects to Tabbycat automatically

## API Requirements
- Tabbycat URL
- API Token (from Change Password page)
- Tournament Slug

## Deploy to Render
1. Push this repo to GitHub
2. Create Web Service on Render
3. Build: pip install -r requirements.txt
4. Start: gunicorn app:app --workers 1 --threads 2 --timeout 120
