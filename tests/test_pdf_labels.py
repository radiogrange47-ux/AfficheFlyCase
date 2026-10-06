import base64
import tempfile
import unittest
from pathlib import Path
import re
import zlib

from PIL import Image

from afficheflycase.pdf_labels import PAGE_HEIGHT, PAGE_WIDTH, generate_pdf
from afficheflycase.workbook import FlyCase, MaterialItem, clean_comment


class PdfLabelTests(unittest.TestCase):
    def make_case(self):
        return FlyCase(
            identifier="LX07",
            case_type="Fly Case",
            color="Bois",
            width="84 cm",
            length="70 cm",
            height="83 cm",
            footprint="0,6 m²",
            volume="0,49 m³",
            comment=clean_comment("Batteries; Tipable / Gerbable; Fragile"),
            source_row=3,
            materials=[
                MaterialItem("Batterie 24V", "CADREJ (1)", "5", "-", 3),
                MaterialItem("Chargeur", "LOGES", "2", "1", 8),
            ],
        )

    def test_generates_one_a4_landscape_page_per_case(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "labels.pdf"
            count = generate_pdf(output, [self.make_case(), self.make_case()],
                                 "Show", "16/10/2026", "Paris")
            content = output.read_bytes()
        self.assertEqual(count, 2)
        self.assertGreater(len(content), 1000)
        self.assertIn(b"/Count 2", content)
        media_box = re.search(rb"/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]", content)
        self.assertIsNotNone(media_box)
        self.assertAlmostEqual(float(media_box.group(1)), PAGE_WIDTH, places=2)
        self.assertAlmostEqual(float(media_box.group(2)), PAGE_HEIGHT, places=2)

    def test_requires_production_fields(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ValueError, "lieu"):
                generate_pdf(Path(directory) / "labels.pdf", [self.make_case()], "Show", "Dates", "")

    def test_label_omits_handling_status_and_cleans_handling_comments(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "labels.pdf"
            generate_pdf(output, [self.make_case()], "Show", "Dates", "Lieu")
            streams = re.findall(rb"stream\r?\n(.*?)endstream", output.read_bytes(), re.S)
            page_text = b"\n".join(
                zlib.decompress(base64.a85decode(stream.strip(), adobe=True))
                for stream in streams
            ).lower()
        self.assertNotIn(b"tipable", page_text)
        self.assertNotIn(b"gerbable", page_text)
        self.assertNotIn(b"basculable", page_text)
        self.assertNotIn(b"do not tip", page_text)
        self.assertIn(b"batteries", page_text)
        self.assertIn(b"fragile", page_text)
        self.assertIn(b"0,6 m\\262", page_text)
        self.assertIn(b"0,49 m\\263", page_text)

    def test_embeds_optional_logo_in_pdf(self):
        with tempfile.TemporaryDirectory() as directory:
            logo = Path(directory) / "logo.png"
            Image.new("RGB", (120, 60), "blue").save(logo)
            output = Path(directory) / "labels.pdf"
            generate_pdf(output, [self.make_case()], "Show", "Dates", "Lieu", logo_path=logo)
            content = output.read_bytes()
        self.assertIn(b"/Subtype /Image", content)

    def test_rejects_missing_logo(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ValueError, "logo est introuvable"):
                generate_pdf(
                    Path(directory) / "labels.pdf",
                    [self.make_case()],
                    "Show",
                    "Dates",
                    "Lieu",
                    logo_path=Path(directory) / "missing.png",
                )


if __name__ == "__main__":
    unittest.main()
