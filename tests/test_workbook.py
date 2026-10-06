import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from afficheflycase.pdf_labels import generate_pdf
from afficheflycase.workbook import WorkbookFormatError, clean_comment, load_flycases


def make_workbook(
    path: Path,
    *,
    sheet_names=None,
    bad_header=False,
    total_rows=False,
    unassigned_rows=False,
) -> None:
    workbook = Workbook()
    material = workbook.active
    material.title = (sheet_names or ["Listing Materiel", "Listing FlyCase"])[0]
    material.append(["", "", "", "", "", ""])
    material.append(["", "WRONG" if bad_header else "ELEMENT", "POSITION", "QTE", "SPARE", "FLYCASE"])
    material.merge_cells("B3:B5")
    material.merge_cells("C3:C5")
    material.merge_cells("D3:D5")
    material.merge_cells("E3:E5")
    material["B3"] = "Lampe"
    material["C3"] = "Scène"
    material["D3"] = 2
    material["E3"] = "-"
    material["F3"] = "LX01 - Fly Case - Noir"
    material["F4"] = "LX02 - Fly Case - Bleu"
    # Row 5 is visually blank but falls inside the merged material cells above.
    if unassigned_rows:
        material["B6"] = "Projecteur non attribué"
        material["C6"] = "Régie"
        material["D6"] = 3
        material["B7"] = "Machine à fumée sans ID valide"
        material["F7"] = "à déterminer"
        material["B8"] = "Machine avec ID inconnu"
        material["F8"] = "LX99 - Fly Case"

    flycases = workbook.create_sheet((sheet_names or ["Listing Materiel", "Listing FlyCase"])[1])
    headers = ["", "ID", "Type2", "Couleur", "Largeur", "Longeur", "Hauteur",
               "Empatement", "Cubage", "Tip", "Gerbable", "Commentaire"]
    flycases.append([""] * len(headers))
    flycases.append(["", "ID", "Type2", "Couleur", "Largeur", "Longeur", "Hauteur",
                     "Empatement", "Cubage", "Tip", "Gerbable", "Commentaire"])
    flycases.append(["", "LX01", "Fly Case", "Noir", "120 cm", "60 cm", "63 cm",
                     "0,7", "0,45", "OK", "OK", "Fragile / Gerbable"])
    flycases.append(["", "LX02", "Fly Case", "Bleu", "50", "40", "30",
                     "-", "0,06", "OK", "NON", ""])
    if total_rows:
        total_row = 9 if unassigned_rows else 6
        material[f"B{total_row}"] = "TOTAL"
        material[f"D{total_row}"] = "=SUM(D3:D5)"
        flycases["B5"] = "TOTAL"
        flycases["E5"] = "=SUM(E3:E4)"
    workbook.save(path)


class WorkbookImportTests(unittest.TestCase):
    def test_loads_cases_and_expands_merged_material_cells(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "listing.xlsx"
            make_workbook(path)
            result = load_flycases(path)
        self.assertEqual([case.identifier for case in result.cases], ["LX01", "LX02"])
        self.assertEqual(len(result.cases[0].materials), 1)
        self.assertEqual(len(result.cases[1].materials), 1)
        self.assertEqual(result.cases[1].materials[0].element, "Lampe")
        self.assertEqual(result.cases[0].dimensions, "120 cm × 60 cm × 63 cm")
        self.assertEqual(result.cases[1].dimensions, "50 cm × 40 cm × 30 cm")
        self.assertEqual(result.cases[0].comment, "Fragile")
        self.assertEqual(result.cases[0].footprint, "0,7")
        self.assertEqual(result.unassigned_materials, [])

    def test_rejects_incorrect_header_with_cell_location(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "wrong.xlsx"
            make_workbook(path, bad_header=True)
            with self.assertRaisesRegex(WorkbookFormatError, r"Listing Materiel.*B2"):
                load_flycases(path)

    def test_rejects_wrong_sheet_order(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "wrong.xlsx"
            make_workbook(path, sheet_names=["Other", "Listing FlyCase"])
            with self.assertRaisesRegex(WorkbookFormatError, "position 1"):
                load_flycases(path)

    def test_removes_handling_terms_from_comments_but_keeps_other_notes(self):
        self.assertEqual(
            clean_comment("Batteries / non gerbable; fragile — tipable"),
            "Batteries — fragile",
        )

    def test_rejects_unmatched_material_case_id(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "wrong.xlsx"
            make_workbook(path)
            from openpyxl import load_workbook

            workbook = load_workbook(path)
            workbook["Listing Materiel"]["F3"] = "LX99 - Fly Case"
            workbook.save(path)
            result = load_flycases(path)
        self.assertEqual(len(result.unassigned_materials), 1)
        self.assertIn("LX99", result.unassigned_materials[0].reason)
        self.assertEqual(result.cases[0].materials, [])
        self.assertEqual(result.cases[1].materials[0].element, "Lampe")

    def test_ignores_total_rows_and_their_formulas(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "listing.xlsx"
            make_workbook(path, total_rows=True)
            result = load_flycases(path)
        self.assertEqual([case.identifier for case in result.cases], ["LX01", "LX02"])
        self.assertEqual(sum(len(case.materials) for case in result.cases), 2)
        self.assertEqual(result.unassigned_materials, [])

    def test_reports_each_unassigned_material_and_keeps_valid_cases(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "listing.xlsx"
            make_workbook(path, unassigned_rows=True)
            result = load_flycases(path)
            output = Path(directory) / "labels.pdf"
            count = generate_pdf(output, result.cases, "Show", "Dates", "Lieu")
            pdf_created = output.is_file()
        self.assertEqual(len(result.unassigned_materials), 3)
        self.assertEqual([item.material.source_row for item in result.unassigned_materials], [6, 7, 8])
        self.assertIn("Projecteur non attribué", result.unassigned_materials[0].description)
        self.assertIn("Aucun fly case", result.unassigned_materials[0].reason)
        self.assertIn("LX99", result.unassigned_materials[2].reason)
        self.assertEqual(sum(len(case.materials) for case in result.cases), 2)
        self.assertEqual(count, 2)
        self.assertTrue(pdf_created)


if __name__ == "__main__":
    unittest.main()
