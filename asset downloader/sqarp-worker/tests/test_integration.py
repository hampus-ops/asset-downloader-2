import io
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
os.environ.setdefault("APP_URL", "https://example.invalid")
os.environ.setdefault("WORKER_SECRET", "test-secret")

from PIL import Image
from openpyxl import Workbook
import asset_downloader_processor as p
import worker as w


def asset(row=2, name="photo.png", kind="image", image_type="environmental"):
    return p.Asset(row, name, "", kind, f"https://example.invalid/{row}", image_type)


class IntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for name in ("IMAGES_DIR", "DOCUMENTS_DIR", "RESUME_AFTER", "REPAIR_EXTREME_ONLY",
                     "ORIGINALS_ONLY", "REMOVE_BACKGROUND", "AI_BACKGROUND_FALLBACK",
                     "CROP_ENABLED", "DOCUMENTS_TO_PDF", "PPI"):
            patcher = patch.object(p, name, getattr(p, name))
            patcher.start()
            self.addCleanup(patcher.stop)
        w.apply_settings({})
        image = Image.new("RGBA", (24, 24), (0, 0, 0, 0))
        for x in range(9, 15):
            for y in range(9, 15):
                image.putpixel((x, y), (255, 0, 0, 255))
        buf = io.BytesIO()
        image.save(buf, "PNG")
        image.close()
        self.png = buf.getvalue()

    def fetch(self, a):
        folder = p.IMAGES_DIR if a.asset_type == "image" else p.DOCUMENTS_DIR
        return p.Download(folder / p.initial_filename(a), self.png, "image/png")

    def run_assets(self, assets, settings=None):
        w.apply_settings(settings or {})
        events = []
        with patch.object(p, "fetch_asset", side_effect=self.fetch):
            w.run_processor(assets, self.root / "files", lambda *args: events.append(args))
        return events

    def test_callback_success_and_failure(self):
        def fetch(a):
            if a.row_number == 3:
                raise ValueError("download failed")
            return self.fetch(a)
        events = []
        with patch.object(p, "fetch_asset", side_effect=fetch):
            w.run_processor([asset(), asset(3)], self.root / "files", lambda *a: events.append(a))
        self.assertEqual([a[0] for a in events], [1, 2])
        self.assertEqual(sum(a[1] is not None for a in events), 1)

    def test_originals_and_document_conversion_disabled(self):
        self.run_assets([asset(), asset(3, "doc.tif", "document")], {"originals_only": True})
        self.assertEqual((p.IMAGES_DIR / "photo.png").read_bytes(), self.png)
        self.assertEqual((p.DOCUMENTS_DIR / "doc.tif").read_bytes(), self.png)
        self.run_assets([asset(4, "another.tif", "document")], {"documents_to_pdf": False})
        self.assertEqual((p.DOCUMENTS_DIR / "another.tif").read_bytes(), self.png)

    def test_crop_and_remove_background(self):
        self.run_assets([asset(image_type="isolated")])
        with Image.open(p.IMAGES_DIR / "photo.png") as image:
            self.assertLess(image.width, 24)
        self.run_assets([asset(image_type="isolated")], {"crop": False, "dpi": 150})
        with Image.open(p.IMAGES_DIR / "photo.png") as image:
            self.assertEqual(image.size, (24, 24))
            self.assertAlmostEqual(image.info["dpi"][0], 150, delta=1)
        self.run_assets([asset(image_type="isolated")], {"remove_background": False})
        with Image.open(p.IMAGES_DIR / "photo.png") as image:
            self.assertEqual(image.mode, "RGB")
            self.assertEqual(image.size, (24, 24))

    def test_ai_disabled_does_not_load_model(self):
        buf = io.BytesIO()
        Image.new("RGB", (24, 24), "white").save(buf, "PNG")
        self.png = buf.getvalue()
        with patch.object(p, "load_background_session", side_effect=AssertionError("AI loaded")):
            events = self.run_assets([asset(image_type="isolated")], {"ai_fallback": False})
        self.assertEqual(events, [(1, None)])

    def test_duplicate_names_survive_zip_flushing(self):
        events = []
        batches = self.root / "batches"
        src = batches / "u" / "j" / "files"
        with patch.object(w, "BATCHES", batches), patch.object(w, "PART_SIZE", 1), \
             patch.object(w, "ensure_space"), patch.object(w, "register_part"), \
             patch.object(p, "fetch_asset", side_effect=self.fetch):
            streamer = w.PartStreamer({"id": "j", "user_id": "u"}, src)
            def report(*args):
                events.append(args)
                streamer.flush(force=True)
            w.run_processor([asset(name="same.jpg"), asset(3, "same.png")], src, report)
        names = []
        for part in batches.rglob("*.zip"):
            with zipfile.ZipFile(part) as archive:
                self.assertIsNone(archive.testzip())
                names.extend(archive.namelist())
        self.assertEqual(len(names), 2)
        self.assertEqual(len(set(names)), 2)

    def test_duplicate_document_names_are_not_skipped(self):
        self.run_assets([asset(2, "same.jpg", "document"), asset(3, "same.png", "document")])
        self.assertEqual(len(list(p.DOCUMENTS_DIR.glob("*.pdf"))), 2)

    def test_worker_resets_desktop_resume_flags(self):
        p.RESUME_AFTER = 100
        p.REPAIR_EXTREME_ONLY = True
        self.assertEqual(self.run_assets([asset()]), [(1, None)])

    def test_settings_reset_between_jobs(self):
        w.apply_settings({"crop": False, "originals_only": True, "dpi": 150})
        w.apply_settings({})
        self.assertTrue(p.CROP_ENABLED)
        self.assertFalse(p.ORIGINALS_ONLY)
        self.assertEqual(p.PPI, 300)

    def test_workbook_to_registered_zip(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "File-list"
        sheet.append(["asset_file_name", "cdn_filename", "asset_type", "asset_url", "image type"])
        sheet.append(["photo.png", "", "image", "https://example.invalid/photo", "environmental"])
        sheet.append(["document.png", "", "document", "https://example.invalid/doc", ""])
        data = io.BytesIO()
        workbook.save(data)
        workbook.close()
        class Response:
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def raise_for_status(self): pass
            def iter_content(self, size): yield data.getvalue()
        calls = []
        def api(action, **payload):
            calls.append((action, payload))
            return {}
        with patch.object(w, "BATCHES", self.root / "batches"), \
             patch.object(w, "WORK_DIR", self.root), \
             patch.object(w.requests, "get", return_value=Response()), \
             patch.object(w, "api", side_effect=api), patch.object(w, "ensure_space"), \
             patch.object(p, "fetch_asset", side_effect=self.fetch):
            w.process({"id": "j", "user_id": "u", "settings": {}}, "https://example.invalid/input")
        self.assertEqual(calls[-1][1]["status"], "done")
        self.assertEqual(calls[-1][1]["failed_urls"], [])
        parts = list((self.root / "batches").rglob("*.zip"))
        self.assertEqual(len(parts), 1)
        with zipfile.ZipFile(parts[0]) as archive:
            self.assertEqual(set(archive.namelist()), {"images/photo.png", "documents/document.pdf"})
        self.assertTrue(any(action == "part" for action, _ in calls))


if __name__ == "__main__":
    unittest.main()
