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
    ORANGE_HEX,
    RED_HEX,
    RS_LOGO_PATH,
    configure_pdf_font,
    format_decimal,
    rounded_corner_cell,
)


def make_deceleration_speed_time_plot(run: dict[str, Any], texts: dict[str, Any], language: str) -> io.BytesIO:
    samples = run.get("plotSamples") or []
    fig, ax = plt.subplots(figsize=(5.8, 3.8))

    if samples:
        times = [float(sample.get("t_rel") or 0) for sample in samples]
        speeds = [float(sample.get("speed_mps") or 0) for sample in samples]
        accelerations = [float(sample.get("acceleration_mps2") or 0) for sample in samples]
        start_index = int(run.get("startIndex") or 0)
        stop_index = int(run.get("stopIndex") or 0)
        mid_index = int(run.get("midIndex") or start_index)
        v_stop = float(run.get("vStop") or 0.2)

        start_rel_index = max(0, min(start_index, len(samples) - 1))
        stop_rel_index = max(0, min(stop_index, len(samples) - 1))
        mid_rel_index = max(0, min(mid_index, len(samples) - 1))
        segment_times = times[start_rel_index:stop_rel_index + 1]
        segment_speeds = speeds[start_rel_index:stop_rel_index + 1]
        segment_accelerations = accelerations[start_rel_index:stop_rel_index + 1]

        ax.plot(segment_times, segment_speeds, color=BLUE_HEX, linewidth=2.5, zorder=3)
        vmax_index = max(range(len(segment_speeds)), key=lambda index: segment_speeds[index])
        decm_index = min(range(len(segment_accelerations)), key=lambda index: segment_accelerations[index])
        vmax_time = segment_times[vmax_index]
        vmax_speed = segment_speeds[vmax_index]
        decm_time = segment_times[decm_index]
        decm_speed = segment_speeds[decm_index]
        stop_time = segment_times[-1]
        stop_speed = segment_speeds[-1]
        mid_time = times[mid_rel_index]

        ax.scatter(vmax_time, vmax_speed, s=48, color=BLUE_HEX, zorder=5)
        ax.scatter(stop_time, stop_speed, s=48, color=ORANGE_HEX, zorder=5)
        ax.scatter(decm_time, decm_speed, s=58, color=RED_HEX, marker="D", zorder=6)
        ax.annotate("VMax", (vmax_time, vmax_speed), xytext=(6, 6), textcoords="offset points", fontsize=9)
        ax.annotate("Stop", (stop_time, stop_speed), xytext=(6, 6), textcoords="offset points", fontsize=9)
        ax.annotate(
            f"DecM\n{format_decimal(run.get('DecM'))}",
            (decm_time, decm_speed),
            xytext=(8, -14),
            textcoords="offset points",
            fontsize=8,
            ha="left",
            va="top",
        )
        ax.axvspan(segment_times[0], mid_time, color=BLUE_HEX, alpha=0.12, zorder=1)
        ax.axvspan(mid_time, stop_time, color=ORANGE_HEX, alpha=0.10, zorder=1)
        ax.axvline(mid_time, linestyle="--", linewidth=1.1, color="#6b7280", alpha=0.85)
        ax.axvline(stop_time, linestyle="--", linewidth=1.0, color="#6b7280", alpha=0.7)
        ax.axhline(v_stop, linestyle=":", linewidth=1.0, color="#6b7280", alpha=0.6)

        zone_y = max(segment_speeds) * 0.5 if segment_speeds else 0
        ax.text((segment_times[0] + mid_time) / 2, zone_y, texts["early_dec"], ha="center", va="top", fontsize=9, weight="bold")
        ax.text((mid_time + stop_time) / 2, zone_y, texts["late_dec"], ha="center", va="top", fontsize=9, weight="bold")
        ax.annotate("", xy=(stop_time, stop_speed), xytext=(segment_times[0], stop_speed), arrowprops=dict(arrowstyle="<->", lw=1.3, alpha=0.9, color="#252423"))
        ax.text(stop_time / 2, stop_speed + (max(segment_speeds) - min(segment_speeds)) * 0.08, f"TTS = {format_decimal(run.get('TTS'))} s", ha="center", va="bottom", fontsize=9, weight="bold")

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlabel("Time [s]" if language == "English" else "Čas [s]")
    ax.set_ylabel("Speed [m/s]" if language == "English" else "Rýchlosť [m/s]")
    ax.set_title(texts["decel_chart_title"], fontsize=10)
    ax.grid(True, linestyle="--", alpha=0.18)
    ax.margins(x=0.02, y=0.20)
    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
    buf.seek(0)
    plt.close(fig)
    return buf


def build_non_normative_deceleration_pdf(
    run: dict[str, Any],
    player_name: str,
    exercise_name: str,
    logo_bytes: bytes | None,
    player_photo_bytes: bytes | None,
    texts: dict[str, Any],
    language: str,
) -> bytes:
    chart_buf = make_deceleration_speed_time_plot(run, texts, language)
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
    exercise_y = 39
    chart_y = 56
    chart_w = left_w
    photo_w = 46
    photo_x = right_x + (right_w - photo_w) / 2
    photo_y = 14
    logo_w = 32
    logo_x = right_x + right_w - logo_w - 4
    logo_y = 14
    metrics_y = 86
    rs_logo_w = 36
    rs_logo_x = pdf.w - rs_logo_w - 10
    rs_logo_y = pdf.h - 15
    if logo_bytes:
        pdf.image(io.BytesIO(logo_bytes), x=logo_x, y=logo_y, w=logo_w)

    pdf.set_text_color(*BLACK_RGB)
    pdf.set_font(font_family, "", 24)
    pdf.set_xy(left_x, top_y)
    pdf.cell(0, 10, player_name, new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(128, 128, 128)
    pdf.set_font(font_family, "", 12)
    pdf.set_xy(left_x, subtitle_y)
    pdf.cell(0, 10, texts["decel_title"], new_x="LMARGIN", new_y="NEXT")
    pdf.set_font(font_family, "", 11)
    pdf.set_xy(left_x, exercise_y)
    pdf.cell(0, 10, exercise_name, new_x="LMARGIN", new_y="NEXT")
    pdf.image(chart_buf, x=left_x, y=chart_y, w=chart_w)
    if player_photo_bytes:
        pdf.image(io.BytesIO(player_photo_bytes), x=photo_x, y=photo_y, w=photo_w, h=photo_w, keep_aspect_ratio=True)

    cell_w = 24
    cell_h = 12
    cell_h_sub = 7
    row_gap = 18
    col_gap = 28
    metrics_block_w = cell_w * 3 + 2 * (col_gap - cell_w)
    data_x = right_x + (right_w - metrics_block_w) / 2
    pdf.set_fill_color(*BLUE_RGB)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_family, "", 12)
    rounded_corner_cell(pdf, data_x, metrics_y, cell_w, cell_h, format_decimal(run.get("averageDeceleration")))
    rounded_corner_cell(pdf, data_x + col_gap, metrics_y, cell_w, cell_h, format_decimal(run.get("DecM")))
    rounded_corner_cell(pdf, data_x + 2 * col_gap, metrics_y, cell_w, cell_h, format_decimal(run.get("VMax")))
    rounded_corner_cell(pdf, data_x, metrics_y + row_gap, cell_w, cell_h, format_decimal(run.get("TTS")))
    rounded_corner_cell(pdf, data_x + col_gap, metrics_y + row_gap, cell_w, cell_h, format_decimal(run.get("DTS")))
    pdf.set_font(font_family, "", 8)
    pdf.set_text_color(220, 220, 220)
    rounded_corner_cell(pdf, data_x, metrics_y + 9, cell_w, cell_h_sub, texts["decel_avg"])
    rounded_corner_cell(pdf, data_x + col_gap, metrics_y + 9, cell_w, cell_h_sub, texts["decel_max"])
    rounded_corner_cell(pdf, data_x + 2 * col_gap, metrics_y + 9, cell_w, cell_h_sub, texts["decel_vmax"])
    rounded_corner_cell(pdf, data_x, metrics_y + row_gap + 9, cell_w, cell_h_sub, texts["decel_tts"])
    rounded_corner_cell(pdf, data_x + col_gap, metrics_y + row_gap + 9, cell_w, cell_h_sub, texts["decel_dts"])
    if RS_LOGO_PATH.is_file():
        pdf.image(str(RS_LOGO_PATH), x=rs_logo_x, y=rs_logo_y, w=rs_logo_w)
    return bytes(pdf.output(dest="S"))
