from __future__ import annotations

import io
import os
from pathlib import Path
from typing import Any

from fpdf import FPDF
from matplotlib import font_manager

BLUE_RGB = (48, 54, 116)
BLACK_RGB = (37, 36, 35)
GREEN_RGB = (0, 192, 96)
ORANGE_RGB = (254, 148, 65)
RED_RGB = (251, 51, 49)
BLUE_HEX = "#303674"
ORANGE_HEX = "#FE9441"
RED_HEX = "#FB3331"

RS_LOGO_PATH = Path(__file__).resolve().parent / "rs-logo.png"
SPLIT_1505_IMAGE_PATH = Path(__file__).resolve().parent / "1505.png"


def format_decimal(value: Any, digits: int = 2) -> str:
    try:
        return f"{float(value):.{digits}f}"
    except (TypeError, ValueError):
        return "-"


def rounded_corner_cell(pdf: FPDF, x: float, y: float, w: float, h: float, text: str) -> None:
    pdf.set_xy(x, y)
    pdf.rect(x, y, w, h, round_corners=True, style="F")
    pdf.cell(w, h, text, 0, 0, "C")


def configure_pdf_font(pdf: FPDF) -> str:
    font_path = font_manager.findfont("DejaVu Sans")
    if font_path and os.path.isfile(font_path):
        pdf.add_font("DejaVuSans", "", font_path)
        return "DejaVuSans"
    return "Helvetica"


def get_quadrant_badge_fill(quadrant: str) -> tuple[int, int, int]:
    if quadrant == "Q1":
        return (14, 108, 79)
    if quadrant == "Q2":
        return ORANGE_RGB
    if quadrant == "Q3":
        return RED_RGB
    return ORANGE_RGB


def get_quadrant_result_fill(quadrant: str) -> tuple[int, int, int]:
    if quadrant == "Q1":
        return GREEN_RGB
    if quadrant == "Q2":
        return (255, 188, 89)
    if quadrant == "Q3":
        return (240, 131, 133)
    return (255, 188, 89)
