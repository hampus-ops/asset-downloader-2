"""SQARP: Asset downloader — background worker.

Runs forever on a small always-on server. It picks up queued jobs from the
website, runs the asset processor over the uploaded workbook (images are cut
out and cropped, documents are turned into PDFs), packs the result into ZIP
parts on this server's own disk and keeps them there for RETENTION_DAYS days.
It also serves the finished parts back to users through signed links the
website hands out, so nothing ever has to be copied into the website's own
file storage.

Environment:
  APP_URL        e.g. https://your-app.lovable.app
  WORKER_SECRET  same value as the WORKER_SECRET saved in Lovable Cloud
  WORKER_URL     this server's public address, e.g. https://worker.up.railway.app
  WORK_DIR       folder with at least 50 GB free, e.g. /data (default ./data)
  PART_SIZE_GB   optional, max size of one ZIP part (default 2)
  HEADROOM_GB    optional, free space kept spare while packing (default 2)
  RETENTION_DAYS optional, handled by the website
  PORT           set by the host (Railway does this automatically)

The per-row work happens in asset_downloader_processor.py. To change what a
row does, edit that file — or point run_processor() at another script.
"""
import hmac
import os
import re
import shutil
import threading
import time
import traceback
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import requests

import asset_downloader_processor as processor

APP_URL = os.environ["APP_URL"].rstrip("/")
SECRET = os.environ["WORKER_SECRET"]
WORKER_URL = os.environ.get("WORKER_URL", "").rstrip("/")
WORK_DIR = Path(os.environ.get("WORK_DIR", "./data")).resolve()
BATCHES = WORK_DIR / "batches"
PART_SIZE = int(float(os.environ.get("PART_SIZE_GB", "2")) * 1024**3)
HEADROOM = int(float(os.environ.get("HEADROOM_GB", "2")) * 1024**3)
PORT = int(os.environ.get("PORT", "8080"))

SAFE_PATH = re.compile(r"^[\w./-]+$")
PURGE_LOCK = threading.Lock()


def api(action: str, **payload) -> dict:
    r = requests.post(
        f"{APP_URL}/api/public/worker/{action}",
        json=payload,
        headers={"Authorization": f"Bearer {SECRET}"},
        timeout=30,
    )
    r.raise_for_status()
    return r.json()


def fail(job: dict, message: str) -> None:
    """Tell the website this batch could not run, so the user sees why."""
    print(f"  job failed: {message}")
    try:
        api("progress", job_id=job["id"], status="failed", error=message[:1500])
    except Exception:
        pass


# ---------------------------------------------------------------- download logic
def apply_settings(settings: dict | None) -> None:
    """Copy the batch's dashboard settings onto the processor's globals."""
    s = settings or {}
    def flag(key: str, default: bool) -> bool:
        v = s.get(key, default)
        return v if isinstance(v, bool) else default
    processor.ORIGINALS_ONLY = flag("originals_only", False)
    processor.REMOVE_BACKGROUND = flag("remove_background", True)
    processor.AI_BACKGROUND_FALLBACK = flag("ai_fallback", True)
    processor.CROP_ENABLED = flag("crop", True)
    processor.DOCUMENTS_TO_PDF = flag("documents_to_pdf", True)
    dpi = s.get("dpi", 300)
    processor.PPI = dpi if isinstance(dpi, int) and 36 <= dpi <= 1200 else 300


def run_processor(assets: list, src: Path, report, before_save=None) -> None:
    """Run the asset processor for one batch.

    Images land in src/images, documents in src/documents — everything under
    src ends up in the ZIP. `report(done, failed_url)` is called after every
    row so the website can show progress; failed_url is that row's link when
    the row could not be downloaded or processed.
    """
    processor.IMAGES_DIR = src / "images"
    processor.DOCUMENTS_DIR = src / "documents"
    processor.create_output_folders()

    def on_asset(completed: int, total: int, asset, error) -> None:
        if error is not None:
            print(f"  row {asset.row_number} failed: {error}")
        report(completed, asset.asset_url if error is not None else None)

    # Desktop resume/repair variables must never skip rows in server jobs.
    processor.RESUME_AFTER = 0
    processor.REPAIR_EXTREME_ONLY = False
    processor.process_assets(assets, report=on_asset, before_save=before_save)


# ------------------------------------------------------------------- disk & space
def safe_target(rel: str) -> Path | None:
    """Resolve a stored path inside WORK_DIR/batches, refusing anything odd."""
    if not rel or not SAFE_PATH.match(rel) or ".." in rel:
        return None
    root = BATCHES.resolve()
    target = (root / rel).resolve()
    if not str(target).startswith(str(root) + os.sep):
        return None
    return target


def free_bytes() -> int:
    return shutil.disk_usage(WORK_DIR).free


def purge_expired() -> int:
    with PURGE_LOCK:
        return _purge_expired()


def _purge_expired() -> int:
    """Ask the website which finished files are past their retention period,
    delete them here, then tell the website they are gone."""
    try:
        rows = api("expired").get("parts", [])
    except Exception:
        return 0
    if not rows:
        return 0
    done = []
    for row in rows:
        path = safe_target(row["path"])
        if path and path.is_file():
            path.unlink()
            for parent in path.parents:
                if parent == BATCHES.resolve():
                    break
                try:
                    parent.rmdir()
                except OSError:
                    break
        done.append(row["id"])
    api("purged", ids=done)
    print(f"Removed {len(done)} expired file(s)")
    return len(done)


def ensure_space(job_id: str, need: int) -> None:
    """Make room for `need` bytes: clear expired batches first, then wait."""
    delay, waited = 30, 0
    while free_bytes() < need:
        if purge_expired() == 0:
            if waited > 12 * 3600:
                raise RuntimeError(
                    "The download server ran out of space and nothing expired in time."
                )
            try:
                api("progress", job_id=job_id,
                    error="The download server is out of space. Waiting for room…")
            except Exception:
                pass
            print(f"Out of space, waiting {delay}s")
            time.sleep(delay)
            waited += delay
            delay = min(delay * 2, 900)
    api("progress", job_id=job_id, error="")


def register_part(job_id: str, number: int, path: Path) -> None:
    """Tell the website one part is finished and ready to download. If the
    website cannot be reached, keep retrying (up to ~12 h) instead of failing
    the whole batch - the part is already safely on disk."""
    rel = path.relative_to(BATCHES).as_posix()
    delay, waited = 30, 0
    while True:
        try:
            api("part", job_id=job_id, part_number=number, path=rel,
                size_bytes=path.stat().st_size, base_url=WORKER_URL)
            api("progress", job_id=job_id, error="")
            return
        except Exception as e:
            if waited > 12 * 3600:
                raise RuntimeError(f"Could not register part {number} after 12 hours: {e}")
            print(f"Registering part {number} failed ({e}); retrying in {delay}s")
            try:
                api("progress", job_id=job_id,
                    error=f"Cannot reach the website. Retrying part {number} automatically…")
            except Exception:
                pass
            time.sleep(delay)
            waited += delay
            delay = min(delay * 2, 900)


class PartStreamer:
    """Packs finished files into ZIP parts on this server's disk while the
    download is still running. Source files are deleted as they are packed, so
    temporary source files are bounded; finished parts remain until expiry."""

    def __init__(self, job: dict, src: Path):
        self.job_id = job["id"]
        self.user_id = job["user_id"]
        self.src = src
        self.dest = BATCHES / job["user_id"] / job["id"]
        self.dest.mkdir(parents=True, exist_ok=True)
        self.n = 0

    def flush(self, force: bool = False) -> None:
        while True:
            files = sorted(p for p in self.src.rglob("*") if p.is_file())
            if not files:
                return
            if not force and sum(f.stat().st_size for f in files) < PART_SIZE:
                return
            needed = min(PART_SIZE, sum(f.stat().st_size for f in files))
            needed = max(needed, files[0].stat().st_size)
            ensure_space(self.job_id, needed + len(files) * 1024 + HEADROOM)
            self.n += 1
            path = self.dest / f"part-{self.n}.zip"
            size = 0
            packed = []
            with zipfile.ZipFile(path, "w", zipfile.ZIP_STORED, allowZip64=True) as z:
                for f in files:
                    if size > 0 and size + f.stat().st_size > PART_SIZE:
                        break
                    z.write(f, f.relative_to(self.src))
                    size += f.stat().st_size
                    packed.append(f)
            # Retain sources until the ZIP central directory is safely written.
            for f in packed:
                f.unlink()
            register_part(self.job_id, self.n, path)
            print(f"  part {self.n} packed: {size / 1024**3:.2f} GB")


def dir_size(path: Path) -> int:
    return sum(f.stat().st_size for f in path.rglob("*") if f.is_file())


# ----------------------------------------------------------------------- the job
def process(job: dict, xlsx_url: str) -> None:
    job_dir = BATCHES / job["user_id"] / job["id"]
    src = job_dir / "files"
    src.mkdir(parents=True, exist_ok=True)
    streamer = PartStreamer(job, src)
    xlsx = job_dir / "input.xlsx"

    try:
        with requests.get(xlsx_url, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(xlsx, "wb") as fh:
                for chunk in r.iter_content(1 << 16):
                    fh.write(chunk)

        # The workbook has to look the way the processor expects: a "File-list"
        # sheet with asset_file_name, cdn_filename, asset_type and asset_url.
        try:
            assets = processor.read_assets(xlsx)
        except Exception as e:
            fail(job, f"The Excel file could not be read: {e}")
            return
        xlsx.unlink()

        settings = job.get("settings") or {}
        apply_settings(settings)
        if settings.get("include_images") is False:
            assets = [a for a in assets if a.asset_type != "image"]
        if settings.get("include_documents") is False:
            assets = [a for a in assets if a.asset_type == "image"]

        total = len(assets)
        api("progress", job_id=job["id"], total=total)
        print(f"  {total} assets to process")

        state = {"done": 0, "failed": []}

        def report(done: int, failed_url: str | None) -> None:
            state["done"] = done
            if failed_url:
                state["failed"].append(failed_url)
            if done % 10 == 0 or done == total:
                api("progress", job_id=job["id"], done=done, failed_urls=state["failed"])
            if dir_size(src) >= PART_SIZE:
                streamer.flush()

        def before_save(need: int) -> None:
            # Pack pending files first so they do not consume the spare space.
            if free_bytes() < need + HEADROOM and dir_size(src):
                streamer.flush(force=True)
            ensure_space(job["id"], need + HEADROOM)

        run_processor(assets, src, report, before_save=before_save)
        streamer.flush(force=True)
        api("progress", job_id=job["id"], done=total, failed_urls=state["failed"],
            status="done")
    except Exception as e:
        traceback.print_exc()
        fail(job, str(e) or e.__class__.__name__)
    finally:
        shutil.rmtree(src, ignore_errors=True)
        try:
            xlsx.unlink(missing_ok=True)
        except OSError:
            pass


# -------------------------------------------------------------------- file server
class DownloadHandler(BaseHTTPRequestHandler):
    """Serves finished ZIP parts to browsers through signed, expiring links."""

    server_version = "SqarpWorker"

    def log_message(self, *args) -> None:  # keep the log clean
        pass

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"ok")
            return
        if not parsed.path.startswith("/d/"):
            self.send_error(404)
            return

        rel = unquote(parsed.path[3:])
        query = parse_qs(parsed.query)
        exp = (query.get("exp") or [""])[0]
        sig = (query.get("sig") or [""])[0]
        if not exp.isdigit() or int(exp) < time.time():
            self.send_error(403, "Link expired")
            return
        expected = hmac.new(SECRET.encode(), f"{rel}|{exp}".encode(), "sha256").hexdigest()
        if not hmac.compare_digest(expected, sig):
            self.send_error(403, "Invalid link")
            return
        path = safe_target(rel)
        if not path or not path.is_file():
            self.send_error(404, "Files are no longer available")
            return
        # Optional friendly download name (e.g. asset-downloader-20261012.zip).
        # Sanitised hard: only the base name survives, and only safe characters.
        raw_name = (query.get("fn") or [""])[0]
        safe_name = re.sub(r'[^A-Za-z0-9._-]', "", Path(raw_name).name)
        if not safe_name or not safe_name.lower().endswith(".zip"):
            safe_name = path.name
        self.stream(path, safe_name)

    def stream(self, path: Path, download_name: str | None = None) -> None:
        size = path.stat().st_size
        start, end = 0, size - 1
        rng = self.headers.get("Range")
        if rng and rng.startswith("bytes="):
            first, _, last = rng[6:].partition("-")
            start = int(first or 0)
            end = min(int(last) if last else size - 1, size - 1)
            if start > end:
                self.send_error(416)
                return
        length = end - start + 1
        self.send_response(206 if rng else 200)
        self.send_header("Content-Type", "application/zip")
        self.send_header("Content-Length", str(length))
        self.send_header("Accept-Ranges", "bytes")
        self.send_header(
            "Content-Disposition", f'attachment; filename="{download_name or path.name}"'
        )
        if rng:
            self.send_header("Content-Range", f"bytes {start}-{end}/{size}")
        self.end_headers()
        with open(path, "rb") as fh:
            fh.seek(start)
            left = length
            while left > 0:
                chunk = fh.read(min(1 << 20, left))
                if not chunk:
                    break
                try:
                    self.wfile.write(chunk)
                except (BrokenPipeError, ConnectionResetError):
                    return
                left -= len(chunk)


def serve_files() -> None:
    ThreadingHTTPServer(("0.0.0.0", PORT), DownloadHandler).serve_forever()


def cleanup_loop() -> None:
    """Keep retention cleanup running even while a large job is active."""
    while True:
        try:
            purge_expired()
        except Exception:
            traceback.print_exc()
        time.sleep(300)


# ------------------------------------------------------------------------- main
def main() -> None:
    BATCHES.mkdir(parents=True, exist_ok=True)
    threading.Thread(target=serve_files, daemon=True).start()
    threading.Thread(target=cleanup_loop, daemon=True).start()
    print(f"Worker started. Polling {APP_URL}; serving files on port {PORT}.")
    while True:
        try:
            res = api("claim")
            if res.get("job"):
                print("Processing job", res["job"]["id"])
                process(res["job"], res["xlsx_url"])
                continue
        except Exception:
            traceback.print_exc()
        time.sleep(15)


if __name__ == "__main__":
    main()
