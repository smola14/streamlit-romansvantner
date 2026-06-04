from __future__ import annotations

import io
import sys
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
from fpdf import FPDF

MODULE_DIR = Path(__file__).resolve().parent
if str(MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(MODULE_DIR))

from report_common import (
    BLACK_RGB,
    BLUE_HEX,
    BLUE_RGB,
    RS_LOGO_PATH,
    SPLIT_1505_IMAGE_PATH,
    configure_pdf_font,
    format_decimal,
    rounded_corner_cell,
)


def make_split_speed_time_plot(run: dict[str, Any]) -> io.BytesIO:
    motions = run.get("motions") or []
    fig, ax = plt.subplots(figsize=(5.8, 3.8))

    if motions:
        x_points = [0.0]
        y_points = [float(motions[0].get("avgSpeed") or 0)]
        for motion in motions:
            start_time = float(motion.get("startTime") or 0)
            end_time = float(motion.get("endTime") or 0)
            avg_speed = float(motion.get("avgSpeed") or 0)
            x_points.extend([start_time, end_time])
            y_points.extend([avg_speed, avg_speed])

        ax.plot(x_points, y_points, color=BLUE_HEX, linewidth=2.5)

        for motion in motions:
            mid_time = (float(motion.get("startTime") or 0) + float(motion.get("endTime") or 0)) / 2
            phase_name = str(motion.get("phaseName") or "")
            if phase_name:
                ax.text(
                    mid_time,
                    float(motion.get("avgSpeed") or 0) + 0.15,
                    phase_name,
                    fontsize=8,
                    color="#6b7280",
                    ha="center",
                )

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlabel("Čas [s]")
    ax.set_ylabel("Rýchlosť [m/s]")
    ax.set_title("Rýchlosť počas runu", fontsize=10)
    ax.grid(True, linestyle="--", alpha=0.18)

    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


def build_non_normative_split_pdf(
    run: dict[str, Any],
    player_name: str,
    logo_bytes: bytes | None,
    player_photo_bytes: bytes | None,
) -> bytes:
    chart_buf = make_split_speed_time_plot(run)
    pdf = FPDF("L", "mm", "A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    font_family = configure_pdf_font(pdf)

    left_x = 14
    left_w = 126
    right_x = 154
    right_w = 129
    top_y = 18
    subtitle_y = 29
    chart_y = 54
    chart_w = left_w
    scheme_y = 137
    scheme_w = 96
    scheme_x = left_x + (left_w - scheme_w) / 2
    photo_w = 54
    photo_x = right_x + 4
    photo_y = 16
    logo_w = 30
    logo_x = right_x + right_w - logo_w
    logo_y = 16
    metrics_y = 92
    rs_logo_w = 36
    rs_logo_x = pdf.w - rs_logo_w - 10
    rs_logo_y = pdf.h - 15

    pdf.set_text_color(*BLACK_RGB)
    pdf.set_font(font_family, "", 24)
    pdf.set_xy(left_x, top_y)
    pdf.cell(0, 10, player_name, new_x="LMARGIN", new_y="NEXT")

    pdf.set_text_color(128, 128, 128)
    pdf.set_font(font_family, "", 12)
    pdf.set_xy(left_x, subtitle_y)
    pdf.cell(0, 10, "Deceleračný profil 15-0-5", new_x="LMARGIN", new_y="NEXT")

    pdf.image(chart_buf, x=left_x, y=chart_y, w=chart_w)
    if SPLIT_1505_IMAGE_PATH.is_file():
        pdf.image(str(SPLIT_1505_IMAGE_PATH), x=scheme_x, y=scheme_y, w=scheme_w)

    if player_photo_bytes:
        pdf.image(io.BytesIO(player_photo_bytes), x=photo_x, y=photo_y, w=photo_w, h=photo_w, keep_aspect_ratio=True)
    if logo_bytes:
        pdf.image(io.BytesIO(logo_bytes), x=logo_x, y=logo_y, w=logo_w)

    cell_w = 24
    cell_h = 12
    cell_h_sub = 7
    row_gap = 18
    col_gap = 28
    metrics_block_w = cell_w * 3 + 2 * (col_gap - cell_w)
    data_x = right_x + (right_w - metrics_block_w) / 2

    total_time = float(run.get("time") or 0)
    top_speed_ms = float(run.get("topSpeed") or 0)
    max_acceleration = float(run.get("maxAcceleration") or 0)
    max_deceleration = float(run.get("maxDeceleration") or 0)
    deceleration_time = float(run.get("decelerationTime") or 0)

    pdf.set_fill_color(*BLUE_RGB)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_family, "", 12)

    rounded_corner_cell(pdf, data_x, metrics_y, cell_w, cell_h, format_decimal(total_time))
    rounded_corner_cell(pdf, data_x + col_gap, metrics_y, cell_w, cell_h, format_decimal(top_speed_ms))
    rounded_corner_cell(pdf, data_x + 2 * col_gap, metrics_y, cell_w, cell_h, format_decimal(top_speed_ms * 3.6))
    rounded_corner_cell(pdf, data_x, metrics_y + row_gap, cell_w, cell_h, format_decimal(max_acceleration))
    rounded_corner_cell(pdf, data_x + col_gap, metrics_y + row_gap, cell_w, cell_h, format_decimal(max_deceleration))
    rounded_corner_cell(pdf, data_x + 2 * col_gap, metrics_y + row_gap, cell_w, cell_h, format_decimal(deceleration_time))

    pdf.set_font(font_family, "", 8)
    pdf.set_text_color(220, 220, 220)
    rounded_corner_cell(pdf, data_x, metrics_y + 9, cell_w, cell_h_sub, "Celkový čas [s]")
    rounded_corner_cell(pdf, data_x + col_gap, metrics_y + 9, cell_w, cell_h_sub, "Max rýchlosť [m/s]")
    rounded_corner_cell(pdf, data_x + 2 * col_gap, metrics_y + 9, cell_w, cell_h_sub, "Max rýchlosť [km/h]")
    rounded_corner_cell(pdf, data_x, metrics_y + row_gap + 9, cell_w, cell_h_sub, "Max akcelerácia [m/s²]")
    rounded_corner_cell(pdf, data_x + col_gap, metrics_y + row_gap + 9, cell_w, cell_h_sub, "Max decelerácia [m/s²]")
    rounded_corner_cell(pdf, data_x + 2 * col_gap, metrics_y + row_gap + 9, cell_w, cell_h_sub, "Čas decelerácie [s]")

    if RS_LOGO_PATH.is_file():
        pdf.image(str(RS_LOGO_PATH), x=rs_logo_x, y=rs_logo_y, w=rs_logo_w)

    return bytes(pdf.output(dest="S"))
