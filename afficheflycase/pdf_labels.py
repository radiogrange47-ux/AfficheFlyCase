"""A4 landscape PDF label generation."""

from __future__ import annotations

import html
import os
import re
import sys
import tempfile
import unicodedata
from pathlib import Path
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

from .workbook import FlyCase


PAGE_SIZE = landscape(A4)
PAGE_WIDTH, PAGE_HEIGHT = PAGE_SIZE
MARGIN = 24
INK = colors.HexColor("#202124")
WARNING_RED = colors.HexColor("#C62828")
APPLICATION_ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
BATTERY_ICON_PATH = APPLICATION_ROOT / "assets" / "battery_warning.png"
LOGO_BOX_WIDTH = 204
LOGO_WIDTH = 192
LOGO_HEIGHT = 96
BATTERY_MARK_WIDTH = LOGO_BOX_WIDTH


class LabelOverflowError(ValueError):
    """Raised when label content cannot fit while remaining readable."""


def _safe(value: str) -> str:
    return html.escape(value, quote=False).replace("\n", "<br/>")


def _paragraph(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    top: float,
    width: float,
    height: float,
    *,
    font_name: str = "Helvetica",
    font_size: int = 12,
    min_font_size: int = 10,
    color=INK,
    leading: float | None = None,
    bold: bool = False,
    align: int = TA_LEFT,
) -> None:
    for size in range(font_size, min_font_size - 1, -1):
        style = ParagraphStyle(
            name=f"label-{size}",
            fontName="Helvetica-Bold" if bold else font_name,
            fontSize=size,
            leading=leading or size * 1.2,
            textColor=color,
            alignment=align,
            spaceAfter=0,
            spaceBefore=0,
            allowWidows=0,
            allowOrphans=0,
        )
        paragraph = Paragraph(text, style)
        _, measured_height = paragraph.wrap(width, height)
        if measured_height <= height:
            paragraph.drawOn(pdf, x, top - measured_height)
            return
    raise LabelOverflowError(
        "Le contenu dépasse l’espace disponible même avec une police de 10 pt. "
        "Réduisez le texte dans le classeur ou vérifiez les informations du show."
    )


def _fit_single_line(text: str, preferred: int, min_size: int, max_width: float) -> int:
    for size in range(preferred, min_size - 1, -1):
        if stringWidth(text, "Helvetica-Bold", size) <= max_width:
            return size
    raise LabelOverflowError(f"Le texte « {text} » est trop long pour tenir sur une ligne.")


def _value_with_unit(value: str, unit: str) -> str:
    value = value.strip()
    if not value or value == "-":
        return value
    value = re.sub(r"\s*m(?:\^?[23]|[²³])\s*$", "", value, flags=re.IGNORECASE)
    return f"{value} {unit}"


def _normalized_terms(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value.casefold())
    return "".join(character for character in normalized if not unicodedata.combining(character))


def _material_marks(case: FlyCase) -> tuple[bool, bool]:
    material_names = " ".join(item.element for item in case.materials)
    normalized = _normalized_terms(material_names)
    battery = re.search(
        r"\b(?:batter(?:ie|ies|y)|accumulateur(?:s)?|pile(?:s)?)\b",
        normalized,
    ) is not None
    chemical = re.search(
        r"\b(?:produit(?:s)?\s+chimique(?:s)?|chimique(?:s)?|chemical(?:s)?|"
        r"solvant(?:s)?|acide(?:s)?|degraissant(?:s)?|detergent(?:s)?|nettoyant(?:s)?|"
        r"peinture(?:s)?|vernis(?:sage)?|diluant(?:s)?|aerosol(?:s)?|resine(?:s)?|"
        r"colle(?:s)?|huile(?:s)?|graisse(?:s)?|javel|alcool(?:s)?|desinfectant(?:s)?|"
        r"liquide(?:s)?|lubrifiant(?:s)?)\b",
        normalized,
    ) is not None
    return battery, chemical


def _draw_battery_mark(pdf: canvas.Canvas, x: float, y: float) -> None:
    if not BATTERY_ICON_PATH.is_file():
        raise FileNotFoundError(f"Le pictogramme batterie est introuvable : {BATTERY_ICON_PATH}")
    pdf.drawImage(
        ImageReader(str(BATTERY_ICON_PATH)),
        x + 6,
        y,
        width=LOGO_WIDTH,
        height=LOGO_HEIGHT,
        preserveAspectRatio=True,
        anchor="c",
        mask="auto",
    )


def _draw_chemical_mark(pdf: canvas.Canvas, x: float, y: float) -> None:
    pdf.setStrokeColor(INK)
    pdf.setLineWidth(1.8)
    path = pdf.beginPath()
    path.moveTo(x + 23, y + 55)
    path.lineTo(x + 35, y + 55)
    path.moveTo(x + 26, y + 55)
    path.lineTo(x + 26, y + 39)
    path.lineTo(x + 15, y + 21)
    path.lineTo(x + 15, y + 18)
    path.lineTo(x + 43, y + 18)
    path.lineTo(x + 43, y + 21)
    path.lineTo(x + 32, y + 39)
    path.lineTo(x + 32, y + 55)
    pdf.drawPath(path, stroke=1, fill=0)
    pdf.line(x + 19, y + 27, x + 39, y + 27)
    pdf.circle(x + 24, y + 31, 1.3, stroke=1, fill=0)
    pdf.circle(x + 33, y + 34, 1.3, stroke=1, fill=0)
    pdf.setFont("Helvetica-Bold", 7)
    pdf.drawCentredString(x + 29, y + 7, "CHIMIQUE")


def _warning_text(case: FlyCase) -> str:
    messages = []
    if case.tip.strip().upper() == "NON":
        messages.append("NE PAS TIPER")
    if case.stackable.strip().upper() == "NON":
        messages.append("NE PAS GERBER")
    return "  /  ".join(messages)


def _draw_label(
    pdf: canvas.Canvas,
    case: FlyCase,
    show: str,
    dates: str,
    venue: str,
    logo_path: Path | None,
) -> None:
    left, right = MARGIN, PAGE_WIDTH - MARGIN
    width = right - left
    top = PAGE_HEIGHT - MARGIN
    pdf.setStrokeColor(INK)
    pdf.setLineWidth(1.4)
    pdf.rect(left, MARGIN, width, PAGE_HEIGHT - 2 * MARGIN, stroke=1, fill=0)

    id_height = 110
    id_bottom = top - id_height
    battery_mark, chemical_mark = _material_marks(case)
    marks = []
    if battery_mark:
        marks.append((_draw_battery_mark, BATTERY_MARK_WIDTH))
    if chemical_mark:
        marks.append((_draw_chemical_mark, 58))
    logo_box_width = LOGO_BOX_WIDTH if logo_path else 0
    marks_width = sum(mark_width for _, mark_width in marks) + max(0, len(marks) - 1) * 6
    id_text_width = width - 28 - logo_box_width - marks_width - (8 if marks else 0)
    id_size = _fit_single_line(case.identifier, 72, 48, id_text_width)
    pdf.setFont("Helvetica-Bold", id_size)
    pdf.setFillColor(INK)
    pdf.drawString(left + 14, id_bottom + (id_height - id_size) / 2 + 3, case.identifier)
    if logo_path:
        pdf.drawImage(
            ImageReader(str(logo_path)),
            right - 14 - 192,
            top - id_height + 7,
            width=LOGO_WIDTH,
            height=LOGO_HEIGHT,
            preserveAspectRatio=True,
            anchor="c",
            mask="auto",
        )
    marks_right = right - 14 - logo_box_width - (8 if logo_box_width else 0)
    for draw_mark, mark_width in marks:
        marks_right -= mark_width
        draw_mark(pdf, marks_right, id_bottom + (7 if draw_mark is _draw_battery_mark else 14))
        marks_right -= 6
    pdf.line(left, id_bottom, right, id_bottom)

    info_top = id_bottom
    info_height = 68
    info_bottom = info_top - info_height
    pdf.line(left, info_bottom, right, info_bottom)
    info_widths = (width * 0.43, width * 0.22, width * 0.35)
    info_lefts = (left, left + info_widths[0], left + info_widths[0] + info_widths[1])
    for index in (1, 2):
        pdf.line(info_lefts[index], info_bottom, info_lefts[index], info_top)
    info = [("PRODUCTION", show), ("DATE", dates), ("LIEU", venue)]
    for index, (label, value) in enumerate(info):
        cell_left = info_lefts[index]
        cell_width = info_widths[index]
        _paragraph(
            pdf,
            f"<b>{label}</b>",
            cell_left + 10,
            info_top - 8,
            cell_width - 20,
            18,
            font_size=18,
            min_font_size=15,
        )
        _paragraph(
            pdf,
            _safe(value),
            cell_left + 10,
            info_top - 35,
            cell_width - 20,
            34,
            font_size=24,
            min_font_size=13,
            bold=True,
        )

    detail_top = info_bottom
    detail_height = 34
    detail_bottom = detail_top - detail_height
    pdf.line(left, detail_bottom, right, detail_bottom)
    detail_values = [
        ("DIMENSIONS", case.dimensions),
        ("", _value_with_unit(case.footprint, "m²")),
        ("", _value_with_unit(case.volume, "m³")),
    ]
    detail_widths = (width * 0.50, width * 0.25, width * 0.25)
    detail_lefts = (left, left + detail_widths[0], left + detail_widths[0] + detail_widths[1])
    for index, (label, value) in enumerate(detail_values):
        if index:
            pdf.line(
                detail_lefts[index],
                detail_bottom,
                detail_lefts[index],
                detail_top,
            )
        if not value:
            continue
        cell_left = detail_lefts[index]
        cell_width = detail_widths[index]
        if label:
            _paragraph(
                pdf,
                f"<b>{label}</b>",
                cell_left + 8,
                detail_top - 4,
                cell_width - 16,
                12,
                font_size=9,
                min_font_size=9,
            )
            _paragraph(
                pdf,
                _safe(value),
                cell_left + 8,
                detail_top - 16,
                cell_width - 16,
                17,
                font_size=11,
                min_font_size=10,
                bold=True,
            )
        else:
            _paragraph(
                pdf,
                _safe(value),
                cell_left + 8,
                detail_top - 8,
                cell_width - 16,
                22,
                font_size=13,
                min_font_size=10,
                bold=True,
            )

    content_top = detail_bottom
    inset = 14
    warning = _warning_text(case)
    warning_height = 76 if warning else 0
    comments_height = min(64, max(40, 27 + case.comment.count("\n") * 12)) if case.comment else 0
    content_bottom = MARGIN + warning_height
    comments_top = content_bottom + comments_height
    pdf.line(left, comments_top, right, comments_top)
    if case.comment:
        _paragraph(
            pdf,
            f"<b>COMMENTAIRES</b><br/>{_safe(case.comment)}",
            left + inset,
            comments_top - 6,
            width - 2 * inset,
            comments_height - 12,
            font_size=11,
            min_font_size=10,
        )
    if warning:
        pdf.setFillColor(WARNING_RED)
        pdf.rect(left, MARGIN, width, warning_height, stroke=0, fill=1)
        warning_size = _fit_single_line(warning, 48, 28, width - 24)
        pdf.setFillColor(colors.white)
        pdf.setFont("Helvetica-Bold", warning_size)
        pdf.drawCentredString(
            PAGE_WIDTH / 2,
            MARGIN + (warning_height - warning_size) / 2 + 2,
            warning,
        )
        pdf.setFillColor(INK)
    pdf.line(left, content_top, right, content_top)
    available_height = content_top - comments_top - 24
    material_lines = []
    for item in case.materials:
        parts = []
        if item.element:
            parts.append(f"<b>{_safe(item.element)}</b>")
        if item.quantity:
            parts.append(f"Qté : {_safe(item.quantity)}")
        if item.spare and item.spare != "-":
            parts.append(f"Spare : {_safe(item.spare)}")
        if item.position:
            parts.append(f"Position : {_safe(item.position)}")
        material_lines.append(" — ".join(parts) if parts else "Élément à vérifier")
    if material_lines:
        _paragraph(
            pdf,
            "<b>CONTENU</b><br/>" + "<br/>".join(material_lines),
            left + inset,
            content_top - 8,
            width - 2 * inset,
            available_height,
            font_size=12,
            min_font_size=10,
        )

    pdf.showPage()


def generate_pdf(
    destination: str | Path,
    cases: Iterable[FlyCase],
    show: str,
    dates: str,
    venue: str,
    logo_path: str | Path | None = None,
) -> int:
    """Write one A4 landscape page per case, atomically replacing the chosen file."""
    selected = list(cases)
    if not selected:
        raise ValueError("Sélectionnez au moins un fly case avant de générer le PDF.")
    for label, value in (("nom du show", show), ("dates", dates), ("lieu", venue)):
        if not value.strip():
            raise ValueError(f"Le champ « {label} » est obligatoire.")
    selected_logo = Path(logo_path) if logo_path else None
    if selected_logo:
        if selected_logo.suffix.casefold() not in {".png", ".jpg", ".jpeg"}:
            raise ValueError("Format de logo non pris en charge. Choisissez une image PNG ou JPEG.")
        if not selected_logo.is_file():
            raise ValueError(f"Le fichier du logo est introuvable : {selected_logo}")
        try:
            ImageReader(str(selected_logo)).getSize()
        except (OSError, ValueError, TypeError) as error:
            raise ValueError(f"Le fichier du logo est illisible : {error}") from error

    output = Path(destination)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{output.stem}_", suffix=".pdf", dir=output.parent, delete=False
        ) as temp:
            temp_name = temp.name
        pdf = canvas.Canvas(temp_name, pagesize=PAGE_SIZE, pageCompression=1)
        pdf.setTitle(f"{show} - étiquettes fly cases")
        pdf.setAuthor("AfficheFlyCase")
        for case in selected:
            _draw_label(pdf, case, show.strip(), dates.strip(), venue.strip(), selected_logo)
        pdf.save()
        os.replace(temp_name, output)
        temp_name = None
        return len(selected)
    except (OSError, PermissionError):
        raise
    finally:
        if temp_name and os.path.exists(temp_name):
            os.unlink(temp_name)
