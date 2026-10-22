"""End-to-end contract tests for the LibreOffice/PyMuPDF CLI.

Run with ``python3 -m unittest discover -s tests -p cli.py -v``.  These tests
use a real local LibreOffice conversion; they never contact an HTTP server.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import unittest


ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "tools" / "convert-pptx.py"
FIXTURE = ROOT / "tests" / "fixtures" / "complex.pptx"
NOTE = "PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT"


def run_cli(*args: object, timeout: int = 180) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(CLI), *(str(arg) for arg in args)],
        cwd=ROOT, capture_output=True, text=True, timeout=timeout, check=False,
    )


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class ConversionCliTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not CLI.is_file():
            raise AssertionError(f"CLI missing: {CLI}")
        if shutil.which("soffice") is None:
            raise unittest.SkipTest("LibreOffice is required for integration tests")
        try:
            import pymupdf  # noqa: F401
        except ImportError as exc:
            raise unittest.SkipTest(f"PyMuPDF is required: {exc}") from exc

        cls.tmp = tempfile.TemporaryDirectory(prefix="pptx-cli-test-")
        cls.work = Path(cls.tmp.name)
        cls.input = cls.work / "발표 $(touch SHOULD_NOT_EXIST) space.pptx"
        shutil.copyfile(FIXTURE, cls.input)
        cls.source_hash = digest(cls.input)
        cls.output = cls.work / "default output.html"
        cls.default = run_cli(cls.input, "-o", cls.output, "--json")

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "tmp"):
            cls.tmp.cleanup()

    def assert_converted(self, result: subprocess.CompletedProcess[str], output: Path) -> str:
        self.assertEqual(result.returncode, 0, result.stderr + "\n" + result.stdout)
        self.assertTrue(output.is_file(), f"missing output {output}")
        html = output.read_text(encoding="utf-8")
        self.assertIn("data:image/svg+xml", html)
        self.assertNotIn(NOTE, html)
        self.assertEqual(digest(self.input), self.source_hash, "source PPTX was modified")
        return html

    def test_default_conversion_and_json(self) -> None:
        html = self.assert_converted(self.default, self.output)
        self.assertIn("<html", html.lower())
        self.assertFalse((ROOT / "SHOULD_NOT_EXIST").exists(), "filename triggered shell expansion")
        self.assertIsInstance(json.loads(self.default.stdout), dict)

    def test_visibility_and_original_slide_selection(self) -> None:
        # "2" means original presentation position 2 (intro), not slide2.xml.
        hidden = self.work / "include-hidden.html"
        picked = self.work / "slide-two.html"
        title_output = self.work / "hostile-title.html"
        title = '<img src=x onerror="window.__pptxTitleExecuted=1">'
        self.assert_converted(run_cli(self.input, "-o", hidden, "--include-hidden"), hidden)
        self.assert_converted(run_cli(self.input, "-o", picked, "--slides", "2"), picked)
        self.assert_converted(run_cli(self.input, "-o", title_output, "--title", title), title_output)

        browser_script = ROOT / "tests" / "cli-browser.cjs"
        if shutil.which("node") and browser_script.is_file():
            result = subprocess.run(
                ["node", str(browser_script), str(self.output), str(hidden), str(picked), str(title_output)],
                cwd=ROOT, capture_output=True, text=True, timeout=90, check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr + "\n" + result.stdout)

    def test_existing_output_is_preserved_on_refusal_and_failure(self) -> None:
        before = digest(self.output)
        refused = run_cli(self.input, "-o", self.output)
        self.assertNotEqual(refused.returncode, 0)
        self.assertEqual(digest(self.output), before)

        # A failed conversion with --overwrite must still leave the old file whole.
        bad = self.work / "corrupt.pptx"
        bad.write_bytes(b"not a presentation")
        failed = run_cli(bad, "-o", self.output, "--overwrite")
        self.assertNotEqual(failed.returncode, 0, failed.stdout + failed.stderr)
        self.assertEqual(digest(self.output), before)

        # Harmless non-OOXML bytes with a marker-like prefix. Without a fixed
        # Impress infilter, LibreOffice can convert this as Writer plain text.
        marker_input = self.work / "nasca-looking.pptx"
        marker_input.write_bytes(b"NASCA\x00\r\n" + b"ordinary test bytes\n" * 100)
        marker = run_cli(marker_input, "-o", self.output, "--overwrite", "--json")
        self.assertNotEqual(marker.returncode, 0, marker.stdout + marker.stderr)
        self.assertEqual(json.loads(marker.stdout)["error"]["code"], "CONVERSION_FAILED")
        self.assertEqual(digest(self.output), before)

    def test_missing_converter_fails_without_output(self) -> None:
        output = self.work / "missing-soffice.html"
        missing = self.work / "missing-soffice-binary"
        result = run_cli(self.input, "-o", output, "--soffice", missing)
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(output.exists())

    def test_dependency_check_needs_no_presentation(self) -> None:
        result = run_cli("--check", "--json")
        self.assertEqual(result.returncode, 0, result.stderr + "\n" + result.stdout)
        report = json.loads(result.stdout)
        self.assertTrue(report["ok"])
        self.assertIn("libreoffice", report)

    def test_empty_and_oversized_inputs_are_classified_before_conversion(self) -> None:
        for name, size, code in (("empty.pptx", 0, "INPUT_EMPTY"),
                                 ("oversized.pptx", 64 * 1024 * 1024 + 1, "INPUT_TOO_LARGE")):
            with self.subTest(name=name):
                source = self.work / name
                with source.open("wb") as stream:
                    stream.truncate(size)
                output = self.work / (name + ".html")
                result = run_cli(source, "-o", output, "--json")
                self.assertNotEqual(result.returncode, 0)
                self.assertEqual(json.loads(result.stdout)["error"]["code"], code)
                self.assertFalse(output.exists())

    @unittest.skipUnless(os.name == "posix", "fake executable uses POSIX shebang")
    def test_timeout_kills_unresponsive_converter(self) -> None:
        fake = self.work / "sleeping-soffice"
        fake.write_text(f"#!{sys.executable}\nimport time\ntime.sleep(30)\n", encoding="utf-8")
        fake.chmod(0o755)
        output = self.work / "timeout.html"
        started = time.monotonic()
        result = run_cli(self.input, "-o", output, "--soffice", fake, "--timeout", "1", "--json", timeout=12)
        self.assertLess(time.monotonic() - started, 10, "child process was not bounded")
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(json.loads(result.stdout)["error"]["code"], "TIMEOUT")
        self.assertFalse(output.exists())

    @unittest.skipUnless(os.name == "posix", "fake executable uses POSIX shebang")
    def test_pdf_count_mismatch_does_not_silently_renumber(self) -> None:
        fake = self.work / "one-page-soffice"
        fake.write_text(
            f"#!{sys.executable}\n"
            "import pathlib, sys, pymupdf\n"
            "args = sys.argv[1:]\n"
            "folder = pathlib.Path(args[args.index('--outdir') + 1])\n"
            "output = folder / (pathlib.Path(args[-1]).stem + '.pdf')\n"
            "document = pymupdf.open()\n"
            "document.new_page()\n"
            "document.save(output)\n"
            "document.close()\n", encoding="utf-8")
        fake.chmod(0o755)
        output = self.work / "mismatched.html"
        result = run_cli(self.input, "-o", output, "--soffice", fake, "--json")
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(json.loads(result.stdout)["error"]["code"], "SLIDE_COUNT_MISMATCH")
        self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
