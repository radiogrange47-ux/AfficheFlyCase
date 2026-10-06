import tempfile
import unittest
from pathlib import Path
import re

from afficheflycase.pdf_labels import PAGE_HEIGHT, PAGE_WIDTH, generate_pdf
from afficheflycase.workbook import FlyCase, MaterialItem


class PdfLabelTests(unittest.TestCase):
    def make_case(self, *, tip="NON"):
        return FlyCase(
            identifier="LX07",
            case_type="Fly Case",
            color="Bois",
            width="84 cm",
            length="70 cm",
            height="83 cm",
            footprint="0,6 m²",
            volume="0,49 m³",
            tip=tip,
            stackable="NON",
            comment="Batteries",
            source_row=3,
            materials=[
                MaterialItem("Batterie 24V", "CADREJ (1)", "5", "-", 3),
                MaterialItem("Chargeur", "LOGES", "2", "1", 8),
            ],
        )

    def test_generates_one_a4_landscape_page_per_case(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "labels.pdf"
            count = generate_pdf(output, [self.make_case(), self.make_case(tip="OK")],
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


if __name__ == "__main__":
    unittest.main()
