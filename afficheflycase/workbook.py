"""Import and validate the AfficheFlyCase Excel workbook format."""

from __future__ import annotations

import re
import unicodedata
import zipfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet


SHEET_HEADERS = {
    "Listing Materiel": {
        2: "ELEMENT",
        3: "POSITION",
        4: "QTE",
        5: "SPARE",
        6: "FLYCASE",
    },
    "Listing FlyCase": {
        2: "ID",
        3: "Type2",
        4: "Couleur",
        5: "Largeur",
        6: "Longeur",
        7: "Hauteur",
        8: "Empatement",
        9: "Cubage",
        10: "Tip",
        11: "Gerbable",
        12: "Commentaire",
    },
}

FLYCASE_ID_PATTERN = re.compile(r"(?<![A-Z0-9])([A-Z]{1,8}\d{1,8})(?![A-Z0-9])", re.IGNORECASE)
HANDLING_TERM_PATTERN = re.compile(
    r"\b(?:tip(?:able)?|do\s+not\s+tip|bascul(?:able|er|e)|gerb(?:able|age|er)|empil(?:able|er))\b",
    re.IGNORECASE,
)
COMMENT_PART_SEPARATOR = re.compile(r"[\r\n;,/|]+|\s+[—–-]\s+")
SUMMARY_ROW_PATTERN = re.compile(
    r"(?<![A-Z])(?:TOTAL|SOUS[\s-]*TOTAL|SUB[\s-]*TOTAL|GRAND[\s-]*TOTAL)(?![A-Z])",
    re.IGNORECASE,
)


class WorkbookFormatError(ValueError):
    """Raised when a workbook does not match the documented template."""


@dataclass(frozen=True)
class MaterialItem:
    element: str
    position: str
    quantity: str
    spare: str
    source_row: int
    other_flycases: tuple[str, ...] = ()

    @property
    def other_flycases_text(self) -> str:
        if not self.other_flycases:
            return ""
        return f" (aussi dans {' et '.join(self.other_flycases)})"

    @property
    def quantity_with_spare(self) -> str:
        quantity = self.quantity.strip()
        spare = self.spare.strip()
        if spare and spare != "-":
            return f"{quantity}+{spare}" if quantity else spare
        return quantity


@dataclass(frozen=True)
class UnassignedMaterial:
    material: MaterialItem
    flycase: str
    reason: str

    @property
    def description(self) -> str:
        parts = [self.material.element or "Matériel non renseigné"]
        if self.material.quantity:
            parts.append(f"Qté : {self.material.quantity}")
        if self.material.spare and self.material.spare != "-":
            parts.append(f"Spare : {self.material.spare}")
        if self.material.position:
            parts.append(f"Position : {self.material.position}")
        return " — ".join(parts)


@dataclass
class WorkbookImportResult:
    cases: list[FlyCase]
    unassigned_materials: list[UnassignedMaterial] = field(default_factory=list)


@dataclass
class FlyCase:
    identifier: str
    case_type: str
    color: str
    width: str
    length: str
    height: str
    footprint: str
    volume: str
    comment: str
    source_row: int
    materials: list[MaterialItem] = field(default_factory=list)
    tip: str = ""
    stackable: str = ""

    @property
    def dimensions(self) -> str:
        dimensions = [
            _with_unit(value, "cm", r"(?:mm|cm|m|in|po)")
            for value in (self.width, self.length, self.height)
            if value
        ]
        return " × ".join(dimensions) if dimensions else "À vérifier"


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value).strip()


def _with_unit(value: str, unit: str, accepted_units: str) -> str:
    value = value.strip()
    if not value or value == "-":
        return value
    if re.search(rf"(?:{accepted_units})\s*$", value, flags=re.IGNORECASE):
        return value
    return f"{value} {unit}"


def format_measure(value: str, unit: str) -> str:
    value = value.strip()
    if not value or value == "-":
        return value
    numeric_value = re.sub(r"\s*m(?:\^?[23]|[²³])\s*$", "", value, flags=re.IGNORECASE)
    try:
        rounded = Decimal(numeric_value.replace(",", ".")).quantize(
            Decimal("0.1"), rounding=ROUND_HALF_UP
        )
    except InvalidOperation:
        if numeric_value != value:
            return f"{numeric_value} {unit}"
        return f"{value} {unit}"
    return f"{rounded:.1f}".replace(".", ",") + f" {unit}"


def clean_comment(value: str) -> str:
    """Remove handling-status phrases while keeping unrelated comment fragments."""
    parts = COMMENT_PART_SEPARATOR.split(value)
    kept_parts = []
    for part in parts:
        normalized = unicodedata.normalize("NFKD", part)
        normalized = "".join(character for character in normalized if not unicodedata.combining(character))
        if HANDLING_TERM_PATTERN.search(normalized):
            continue
        cleaned = part.strip(" \t\r\n.,;:/|—–-")
        if cleaned:
            kept_parts.append(cleaned)
    return " — ".join(kept_parts)


def _normalized_id(value: str) -> str:
    return value.strip().upper()


def _merged_value(sheet: Worksheet, row: int, column: int) -> Any:
    cell = sheet.cell(row=row, column=column)
    if not isinstance(cell, MergedCell):
        return cell.value
    for merged_range in sheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return sheet.cell(merged_range.min_row, merged_range.min_col).value
    return None


def _cached_value(
    formulas_sheet: Worksheet, values_sheet: Worksheet, row: int, column: int, sheet_name: str
) -> Any:
    formula = _merged_value(formulas_sheet, row, column)
    value = _merged_value(values_sheet, row, column)
    if isinstance(formula, str) and formula.startswith("=") and value is None:
        address = formulas_sheet.cell(row=row, column=column).coordinate
        raise WorkbookFormatError(
            f"Formule sans résultat enregistré dans « {sheet_name} »!{address}. "
            "Ouvrez le classeur dans Excel, recalculez-le, enregistrez-le puis réessayez."
        )
    return value


def _direct_cached_value(
    formulas_sheet: Worksheet, values_sheet: Worksheet, row: int, column: int, sheet_name: str
) -> Any:
    formula_cell = formulas_sheet.cell(row=row, column=column)
    value_cell = values_sheet.cell(row=row, column=column)
    if isinstance(formula_cell, MergedCell):
        return None
    formula = formula_cell.value
    value = value_cell.value
    if isinstance(formula, str) and formula.startswith("=") and value is None:
        raise WorkbookFormatError(
            f"Formule sans résultat enregistré dans « {sheet_name} »!{formula_cell.coordinate}. "
            "Ouvrez le classeur dans Excel, recalculez-le, enregistrez-le puis réessayez."
        )
    return value


def _is_non_data_row(sheet: Worksheet, row: int, first_column: int, headers: dict[int, str]) -> bool:
    values = [
        _text(sheet.cell(row=row, column=column).value)
        for column in range(first_column, max(headers) + 1)
    ]
    if any(SUMMARY_ROW_PATTERN.search(value) for value in values):
        return True
    return all(
        values[column - first_column].casefold() == header.casefold()
        for column, header in headers.items()
    )


def _validate_sheet_names(workbook) -> None:
    expected = list(SHEET_HEADERS)
    actual = workbook.sheetnames
    for index, expected_name in enumerate(expected):
        if index >= len(actual):
            raise WorkbookFormatError(
                f"Onglet « {expected_name} » manquant. Le classeur doit commencer par "
                "« Listing Materiel », puis « Listing FlyCase »."
            )
        if actual[index].strip().casefold() != expected_name.casefold():
            raise WorkbookFormatError(
                f"Le premier onglet attendu à la position {index + 1} est "
                f"« {expected_name} », mais le fichier contient « {actual[index]} »."
            )


def _validate_headers(sheet: Worksheet, expected: dict[int, str]) -> None:
    for column, expected_header in expected.items():
        actual = _text(sheet.cell(row=2, column=column).value)
        if actual.casefold() != expected_header.casefold():
            address = sheet.cell(row=2, column=column).coordinate
            displayed = actual if actual else "cellule vide"
            raise WorkbookFormatError(
                f"En-tête incorrect dans « {sheet.title} »!{address} : "
                f"attendu « {expected_header} », trouvé « {displayed} »."
            )


def _extract_case_ids(description: str) -> list[str]:
    return [_normalized_id(match) for match in FLYCASE_ID_PATTERN.findall(description.upper())]


def _material_group_key(sheet: Worksheet, row: int) -> tuple[int, int]:
    merged_row_ranges = [
        (merged_range.min_row, merged_range.max_row)
        for merged_range in sheet.merged_cells.ranges
        if merged_range.min_col <= 5
        and merged_range.max_col >= 2
        and merged_range.min_row <= row <= merged_range.max_row
        and merged_range.max_row > merged_range.min_row
    ]
    if not merged_row_ranges:
        return row, row
    return (
        min(first_row for first_row, _ in merged_row_ranges),
        max(last_row for _, last_row in merged_row_ranges),
    )


def load_flycases(path: str | Path) -> WorkbookImportResult:
    """Validate and load all fly cases from the first two workbook sheets."""
    workbook_path = Path(path)
    if workbook_path.suffix.casefold() not in {".xlsx", ".xlsm"}:
        raise WorkbookFormatError("Format non pris en charge. Sélectionnez un fichier .xlsx ou .xlsm.")
    if not workbook_path.is_file():
        raise WorkbookFormatError("Le fichier Excel sélectionné est introuvable.")

    try:
        formulas_book = load_workbook(workbook_path, read_only=False, data_only=False, keep_vba=False)
        values_book = load_workbook(workbook_path, read_only=False, data_only=True, keep_vba=False)
    except PermissionError as error:
        raise WorkbookFormatError(
            "Impossible d’accéder au fichier Excel. Fermez-le dans Excel ou vérifiez les droits "
            "d’accès, puis réessayez."
        ) from error
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as error:
        raise WorkbookFormatError(
            f"Impossible de lire le fichier Excel : {error}. Vérifiez qu’il n’est pas corrompu "
            "et qu’il s’agit bien d’un classeur Excel pris en charge."
        ) from error

    try:
        _validate_sheet_names(formulas_book)
        material_formula_sheet = formulas_book["Listing Materiel"]
        flycase_formula_sheet = formulas_book["Listing FlyCase"]
        material_values_sheet = values_book["Listing Materiel"]
        flycase_values_sheet = values_book["Listing FlyCase"]
        _validate_headers(material_formula_sheet, SHEET_HEADERS["Listing Materiel"])
        _validate_headers(flycase_formula_sheet, SHEET_HEADERS["Listing FlyCase"])

        cases: dict[str, FlyCase] = {}
        for row in range(3, flycase_formula_sheet.max_row + 1):
            if _is_non_data_row(flycase_formula_sheet, row, 2, SHEET_HEADERS["Listing FlyCase"]):
                continue
            values = [
                _cached_value(flycase_formula_sheet, flycase_values_sheet, row, column, "Listing FlyCase")
                for column in range(2, 13)
            ]
            if not any(_text(value) for value in values):
                continue
            identifier = _normalized_id(_text(values[0]))
            if not identifier:
                raise WorkbookFormatError(
                    f"ID manquant dans « Listing FlyCase »!B{row}. Corrigez la ligne ou supprimez-la."
                )
            if identifier in cases:
                raise WorkbookFormatError(
                    f"ID en double « {identifier} » dans « Listing FlyCase »!B{row}."
                )
            cases[identifier] = FlyCase(
                identifier=identifier,
                case_type=_text(values[1]),
                color=_text(values[2]),
                width=_text(values[3]),
                length=_text(values[4]),
                height=_text(values[5]),
                footprint=_text(values[6]),
                volume=_text(values[7]),
                tip=_text(values[8]),
                stackable=_text(values[9]),
                comment=clean_comment(_text(values[10])),
                source_row=row,
            )

        if not cases:
            raise WorkbookFormatError(
                "Aucune fiche avec un ID n’a été trouvée dans « Listing FlyCase »."
            )

        unassigned_materials: list[UnassignedMaterial] = []
        assigned_materials: dict[tuple[int, int], MaterialItem] = {}
        assigned_case_ids: dict[tuple[int, int], list[str]] = {}
        for row in range(3, material_formula_sheet.max_row + 1):
            if _is_non_data_row(material_formula_sheet, row, 2, SHEET_HEADERS["Listing Materiel"]):
                continue
            direct_values = [
                _direct_cached_value(
                    material_formula_sheet, material_values_sheet, row, column, "Listing Materiel"
                )
                for column in range(2, 7)
            ]
            if not any(_text(value) for value in direct_values):
                continue
            if not _text(direct_values[4]):
                if any(_text(value) for value in direct_values[:4]):
                    raw = [
                        _cached_value(
                            material_formula_sheet, material_values_sheet, row, column, "Listing Materiel"
                        )
                        for column in range(2, 7)
                    ]
                    element, position, quantity, spare, _ = map(_text, raw)
                    unassigned_materials.append(
                        UnassignedMaterial(
                            material=MaterialItem(element, position, quantity, spare, row),
                            flycase="",
                            reason=f"Aucun fly case n’est indiqué en F{row}.",
                        )
                    )
                continue
            raw = [
                _cached_value(material_formula_sheet, material_values_sheet, row, column, "Listing Materiel")
                for column in range(2, 7)
            ]
            element, position, quantity, spare, description = map(_text, raw)
            identifiers = _extract_case_ids(description)
            if not identifiers:
                unassigned_materials.append(
                    UnassignedMaterial(
                        material=MaterialItem(element, position, quantity, spare, row),
                        flycase=description,
                        reason=f"Aucun ID de fly case reconnu en F{row} ({description}).",
                    )
                )
                continue
            material = MaterialItem(
                element=element,
                position=position,
                quantity=quantity,
                spare=spare,
                source_row=row,
            )
            unknown_ids = [identifier for identifier in dict.fromkeys(identifiers) if identifier not in cases]
            if unknown_ids:
                unassigned_materials.append(
                    UnassignedMaterial(
                        material=material,
                        flycase=description,
                        reason=(
                            f"ID « {', '.join(unknown_ids)} » de F{row} absent de "
                            "l’onglet « Listing FlyCase »."
                        ),
                    )
                )
                continue
            unique_identifiers = list(dict.fromkeys(identifiers))
            group_key = _material_group_key(material_formula_sheet, row)
            assigned_materials.setdefault(group_key, material)
            grouped_ids = assigned_case_ids.setdefault(group_key, [])
            for identifier in unique_identifiers:
                if identifier not in grouped_ids:
                    grouped_ids.append(identifier)

        for group_key, material in assigned_materials.items():
            identifiers = assigned_case_ids[group_key]
            for identifier in identifiers:
                cases[identifier].materials.append(
                    MaterialItem(
                        element=material.element,
                        position=material.position,
                        quantity=material.quantity,
                        spare=material.spare,
                        source_row=material.source_row,
                        other_flycases=tuple(
                            other_id for other_id in identifiers if other_id != identifier
                        ),
                    )
                )
        return WorkbookImportResult(list(cases.values()), unassigned_materials)
    finally:
        formulas_book.close()
        values_book.close()
