# SQARP: Asset downloader worker

An always-on Python program that does the downloading and keeps the finished ZIP files on its own disk.

## What it needs
- A machine that stays online with **at least 50 GB of free disk** for the finished files.
- A fixed web address, so the website can hand users a download link.

## What the worker does with each row
`asset_downloader_processor.py` is the script that handles the spreadsheet: an
image marked `isolated` is cut out and cropped to a transparent PNG, an image
marked `environmental` keeps its background, and a document is saved as a PDF.
To change that, edit this file — or point `run_processor()` in `worker.py` at
another script. Extra packages go in `requirements.txt`.

The background-removal model (`u2netp`) is fetched while the Docker image is
built when the model host is reachable. If that download fails, the model is
downloaded on first use instead.

Batch settings support `originals_only`, `remove_background`, `ai_fallback`,
`crop`, `documents_to_pdf`, `include_images`, `include_documents` and `dpi`.
Originals are saved without conversion. With background removal disabled,
images keep their background and dimensions. With AI fallback disabled,
existing transparency and reviewed masks can still be used; other images
keep their background because the white-product safeguard requires AI.
PDF conversion supports existing PDFs and Pillow-readable images, not Office
documents. Duplicate output names receive a row suffix so files are not lost.

## The Excel file must look like this
One sheet named **File-list**, with a header row containing:

| column | what it is |
| --- | --- |
| `asset_file_name` | file name used in the ZIP |
| `cdn_filename` | fallback name |
| `asset_type` | `image` or `document` |
| `asset_url` | the link to download |
| `image type` | `isolated` or `environmental` (needed for images) |

Rows without a link, or with another `asset_type`, are skipped. The website
must implement the same validation; the worker also validates the workbook.

## Run locally
Run these commands from the `sqarp-worker` folder. PowerShell:
```powershell
pip install -r requirements.txt
$env:APP_URL = "https://your-app.lovable.app"
$env:WORKER_SECRET = "your-secret"
$env:WORKER_URL = "http://localhost:8080"
$env:WORK_DIR = "./data"
python worker.py
```

Linux/macOS:
```sh
pip install -r requirements.txt
APP_URL=https://your-app.lovable.app WORKER_SECRET=... \
WORKER_URL=http://localhost:8080 WORK_DIR=./data python worker.py
```

## Deploy on Railway
1. Push the contents of `sqarp-worker/` to a GitHub repo, or push the parent folder and set Root Directory to `sqarp-worker`.
2. Railway → New Project → Deploy from GitHub. It detects the Dockerfile.
3. Settings → **Volumes** → add a volume of 50 GB or more, mounted at `/data`.
4. Variables:
   - `APP_URL` — your published site URL
   - `WORKER_SECRET` — the same value saved in Lovable Cloud
   - `WORKER_URL` — this service's own address, the `https://` form (Railway → Settings → Networking)
   - `WORK_DIR` — `/data`
   - optional `PART_SIZE_GB` (default 2) and `HEADROOM_GB` (default 2, free space kept spare)
5. Deploy. Railway sets `PORT` on its own; the worker serves downloads on that port.

## How finished files are kept
- ZIP parts are written to `WORK_DIR/batches/<user>/<job>/part-N.zip`. The website controls expiry (configure 7 days there); this worker does not calculate retention itself.
- The website shows "Available until 16 Oct 2026 · 6 days and 3 hours left", and the files are deleted automatically once that runs out.
- Starting a new batch marks the previous one's files for removal at the next cleanup.
- Free space is checked before saving assets and packing ZIP parts. Images use an estimate based on pixel dimensions; unusually large files can still exceed available space. Finished parts remain on disk until expiry, so a batch must fit alongside retained batches. Expired files are checked every five minutes, including during active jobs. If space is insufficient, the worker retries for up to 12 hours.
- Downloads go straight from this server to the user's browser through a signed link the website issues. The link expires within an hour and only works for the part it was issued for.

Run **one** worker. Multiple workers require atomic job claiming in the website
API and routing each download to the worker holding that file; neither is
provided by these files.

## Verify locally
```sh
python -m unittest discover -s tests -v
```
Tests use generated workbooks and images and mock the website API and asset
downloads. They do not require a running website or an AI model download.
