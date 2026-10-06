import base64
import tempfile
import unittest
from pathlib import Path
import re
import zlib

from PIL import Image

from afficheflycase.pdf_labels import PAGE_HEIGHT, PAGE_WIDTH, generate_pdf
from afficheflycase.workbook import FlyCase, MaterialItem, clean_comment


def pdf_text_streams(content: bytes) -> list[bytes]:
    decoded = []
    for match in re.finditer(rb"stream\r?\n(.*?)endstream", content, re.S):
        header = content[max(0, match.start() - 256):match.start()]
        if b"ASCII85Decode" not in header or b"/Subtype /Image" in header:
            continue
        decoded.append(zlib.decompress(base64.a85decode(match.group(1).strip(), adobe=True)))
    return decoded


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
            page_text = b"\n".join(pdf_text_streams(output.read_bytes())).lower()
        self.assertNotIn(b"tipable", page_text)
        self.assertNotIn(b"gerbable", page_text)
        self.assertNotIn(b"basculable", page_text)
        self.assertNotIn(b"do not tip", page_text)
        self.assertIn(b"batteries", page_text)
        self.assertIn(b"fragile", page_text)
        self.assertIn(b"0,6 m\\262", page_text)
        self.assertIn(b"0,49 m\\263", page_text)

    def test_draws_battery_and_chemical_marks_when_materials_match(self):
        case = self.make_case()
        case.materials.extend(
            [
                MaterialItem("Produit chimique de nettoyage", "", "1", "-", 9),
                MaterialItem("Accumulateur", "", "1", "-", 10),
            ]
        )
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "labels.pdf"
            logo = Path(directory) / "logo.png"
            Image.new("RGB", (120, 60), "blue").save(logo)
            generate_pdf(output, [case], "Show", "Dates", "Lieu", logo_path=logo)
            content = output.read_bytes()
            streams = pdf_text_streams(content)
            page_text = b"\n".join(streams).lower()
        self.assertIn(b"batterie", page_text)
        self.assertIn(b"chimique", page_text)
        self.assertIn(b"/Subtype /Image", content)
        self.assertRegex(streams[0], rb"109\.25 0 0 96 ")
        self.assertRegex(streams[0], rb"192 0 0 96 ")

    def test_shows_warning_when_tip_or_stackability_is_non(self):
        cases = [self.make_case(), self.make_case(), self.make_case()]
        cases[0].tip = "NON"
        cases[1].stackable = "NON"
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "labels.pdf"
            generate_pdf(output, cases, "Show", "Dates", "Lieu")
            pages = [stream.lower() for stream in pdf_text_streams(output.read_bytes())]
        self.assertIn(b"ne pas tiper", pages[0])
        self.assertIn(b"ne pas gerber", pages[1])
        self.assertNotIn(b"ne pas tiper", pages[2])
        self.assertNotIn(b"ne pas gerber", pages[2])
        self.assertIn(b"48 tf", pages[0])

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
