"""A4 landscape PDF label generation."""

from __future__ import annotations

import html
import os
import tempfile
from pathlib import Path
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

from .workbook import FlyCase


PAGE_SIZE = landscape(A4)
PAGE_WIDTH, PAGE_HEIGHT = PAGE_SIZE
MARGIN = 24
RED = colors.HexColor("#D62828")
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


def _draw_label(pdf: canvas.Canvas, case: FlyCase, show: str, dates: str, venue: str) -> None:
    left, right = MARGIN, PAGE_WIDTH - MARGIN
    width = right - left
    top = PAGE_HEIGHT - MARGIN
    pdf.setStrokeColor(INK)
    pdf.setLineWidth(1.4)
    pdf.rect(left, MARGIN, width, PAGE_HEIGHT - 2 * MARGIN, stroke=1, fill=0)

    header_height = 104
    header_bottom = top - header_height
    pdf.line(left, header_bottom, right, header_bottom)
    show_size = _fit_single_line(show, 25, 14, width - 28)
    pdf.setFont("Helvetica-Bold", show_size)
    pdf.setFillColor(INK)
    pdf.drawString(left + 14, top - 31, show)
    _paragraph(
        pdf,
        f"<b>DATES</b>  {_safe(dates)}     <b>LIEU</b>  {_safe(venue)}",
        left + 14,
        top - 42,
        width - 28,
        28,
        font_size=14,
        min_font_size=10,
    )
    case_title = f"{case.identifier}  |  {case.case_type or 'Fly Case'}"
    title_size = _fit_single_line(case_title, 30, 18, width - 28)
    pdf.setFont("Helvetica-Bold", title_size)
    pdf.drawString(left + 14, header_bottom + 13, case_title)

    detail_top = header_bottom
    detail_height = 82
    detail_bottom = detail_top - detail_height
    pdf.line(left, detail_bottom, right, detail_bottom)
    third = width / 3
    for index in (1, 2):
        pdf.line(left + index * third, detail_bottom, left + index * third, detail_top)

    details = [
        ("DIMENSIONS", case.dimensions),
        ("COULEUR", case.color or "—"),
        ("GERBAGE", case.stackable_label),
    ]
    for index, (label, value) in enumerate(details):
        cell_left = left + index * third
        _paragraph(
            pdf,
            f"<b>{label}</b>",
            cell_left + 10,
            detail_top - 10,
            third - 20,
            18,
            font_size=13,
            min_font_size=11,
        )
        _paragraph(
            pdf,
            _safe(value),
            cell_left + 10,
            detail_top - 34,
            third - 20,
            36,
            font_size=18,
            min_font_size=11,
            bold=True,
        )
    if case.footprint or case.volume:
        extras = "     ".join(
            f"{label} : {_safe(value)}"
            for label, value in (("EMPATTEMENT", case.footprint), ("CUBAGE", case.volume))
            if value
        )
        _paragraph(pdf, extras, left + 10, detail_bottom - 3, width - 20, 20, font_size=11)
        detail_bottom -= 20
        pdf.line(left, detail_bottom, right, detail_bottom)

    content_top = detail_bottom
    warning_height = 49 if case.tip == "NON" else 28
    warning_top = MARGIN + warning_height
    content_bottom = warning_top
    pdf.line(left, content_bottom, right, content_bottom)
    pdf.line(left, content_top, right, content_top)
    inset = 14
    available_height = content_top - content_bottom - 34
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
    if case.comment:
        body += f"<br/><br/><b>COMMENTAIRES</b><br/>{_safe(case.comment)}"
    _paragraph(
        pdf,
        "<b>CONTENU / REMARQUES</b><br/>" + body,
        left + inset,
        content_top - 11,
        width - 2 * inset,
        available_height,
        font_size=12,
        min_font_size=10,
    )

    pdf.setFillColor(RED if case.tip == "NON" else colors.HexColor("#E9ECEF"))
    pdf.rect(left, MARGIN, width, warning_height, stroke=0, fill=1)
    warning = case.tip_label
    if case.tip == "NON":
        warning += "  /  DO NOT TIP"
        font_size = _fit_single_line(warning, 34, 20, width - 24)
        pdf.setFillColor(colors.white)
    else:
        font_size = _fit_single_line(warning, 14, 10, width - 24)
        pdf.setFillColor(INK)
    pdf.setFont("Helvetica-Bold", font_size)
    pdf.drawCentredString(PAGE_WIDTH / 2, MARGIN + (warning_height - font_size) / 2 + 2, warning)
    pdf.setFillColor(INK)
    pdf.showPage()


def generate_pdf(
    destination: str | Path,
    cases: Iterable[FlyCase],
    show: str,
    dates: str,
    venue: str,
) -> int:
    """Write one A4 landscape page per case, atomically replacing the chosen file."""
    selected = list(cases)
    if not selected:
        raise ValueError("Sélectionnez au moins un fly case avant de générer le PDF.")
    for label, value in (("nom du show", show), ("dates", dates), ("lieu", venue)):
        if not value.strip():
            raise ValueError(f"Le champ « {label} » est obligatoire.")

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
            _draw_label(pdf, case, show.strip(), dates.strip(), venue.strip())
        pdf.save()
        os.replace(temp_name, output)
        temp_name = None
        return len(selected)
    except (OSError, PermissionError):
        raise
    finally:
        if temp_name and os.path.exists(temp_name):
            os.unlink(temp_name)
