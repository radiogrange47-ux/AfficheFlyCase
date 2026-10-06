"""A4 landscape PDF label generation."""

from __future__ import annotations

import html
import os
import re
import tempfile
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
    logo_box_width = 204 if logo_path else 0
    id_text_width = width - 28 - logo_box_width
    id_size = _fit_single_line(case.identifier, 72, 48, id_text_width)
    pdf.setFont("Helvetica-Bold", id_size)
    pdf.setFillColor(INK)
    pdf.drawString(left + 14, id_bottom + (id_height - id_size) / 2 + 3, case.identifier)
    if logo_path:
        pdf.drawImage(
            ImageReader(str(logo_path)),
            right - 14 - 192,
            top - id_height + 7,
            width=192,
            height=id_height - 14,
            preserveAspectRatio=True,
            anchor="c",
            mask="auto",
        )
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
    comments_height = min(64, max(40, 27 + case.comment.count("\n") * 12)) if case.comment else 0
    comments_top = MARGIN + comments_height
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
    body = "<br/>".join(material_lines) if material_lines else "Aucun élément associé dans le listing matériel."
    _paragraph(
        pdf,
        "<b>CONTENU</b><br/>" + body,
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
