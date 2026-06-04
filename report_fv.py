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
    configure_pdf_font,
    get_quadrant_badge_fill,
    get_quadrant_result_fill,
    rounded_corner_cell,
)


def make_fv_profile_player_only(report: dict[str, Any]) -> io.BytesIO:
    v0 = float(report["v0"])
    f0 = float(report["f0"])
    x_player = [0, v0]
    y_player = [f0, 0]
    bbox = dict(boxstyle="round", edgecolor="none", facecolor=BLUE_HEX)

    fig, ax = plt.subplots()
    ax.plot(x_player, y_player, label="Player", color=BLUE_HEX, linewidth=2.5)
    ax.annotate(str(round(v0, 2)), (v0, 0), xytext=(v0, -0.6), textcoords="data", ha="center", va="center", color="white", bbox=bbox)
    ax.annotate(str(round(f0, 2)), (0, f0), xytext=(-0.6, f0), textcoords="data", ha="center", va="center", color="white", bbox=bbox)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlabel("V0 [m/s]")
    ax.set_ylabel("F0 [N/kg]")
    ax.legend(loc="upper right", frameon=False)
    ax.margins(x=0.05, y=0.05)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


def make_norm_scatter_plot(report: dict[str, Any], norm_row: dict[str, Any], scatter_entry: dict[str, Any], texts: dict[str, Any]) -> io.BytesIO:
    points = scatter_entry.get("points") or []
    x_values = [float(point["v0"]) for point in points]
    y_values = [float(point["f0"]) for point in points]
    player_v0 = float(report["v0"])
    player_f0 = float(report["f0"])
    v0_median = float(norm_row["v0_median"])
    f0_median = float(norm_row["f0_median"])

    all_x = [*x_values, player_v0]
    all_y = [*y_values, player_f0]
    x_span = max(all_x) - min(all_x) if all_x else 1.0
    y_span = max(all_y) - min(all_y) if all_y else 1.0
    x_pad = max(x_span * 0.08, 0.15)
    y_pad = max(y_span * 0.08, 0.15)

    fig, ax = plt.subplots()
    if x_values and y_values:
        ax.scatter(x_values, y_values, color="#d0d4db", s=32, alpha=0.9, edgecolors="none")

    ax.scatter([player_v0], [player_f0], color="#FB3331", s=80, zorder=3)
    ax.axhline(f0_median, color="#252423", linestyle="--", linewidth=1.2)
    ax.axvline(v0_median, color="#252423", linestyle="--", linewidth=1.2)
    ax.text(v0_median + x_pad * 0.15, max(all_y) + y_pad * 0.1, "Q1", fontsize=10, color="#252423")
    ax.text(min(all_x) - x_pad * 0.1, max(all_y) + y_pad * 0.1, "Q2", fontsize=10, color="#252423")
    ax.text(min(all_x) - x_pad * 0.1, min(all_y) - y_pad * 0.35, "Q3", fontsize=10, color="#252423")
    ax.text(v0_median + x_pad * 0.15, min(all_y) - y_pad * 0.35, "Q4", fontsize=10, color="#252423")
    ax.set_xlim(min(all_x) - x_pad, max(all_x) + x_pad)
    ax.set_ylim(min(all_y) - y_pad, max(all_y) + y_pad)
    ax.set_xlabel("V0 [m/s]")
    ax.set_ylabel("F0 [N/kg]")
    ax.set_title(f"{texts['quadrant_reference']} | {norm_row.get('category')}")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, linestyle="--", alpha=0.18)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


def make_normative_fv_profile(report: dict[str, Any], norm_row: dict[str, Any], texts: dict[str, Any]) -> io.BytesIO:
    player_v0 = float(report["v0"])
    player_f0 = float(report["f0"])
    bbox = dict(boxstyle="round", edgecolor="none", facecolor=BLUE_HEX)
    upper_x = [0, float(norm_row["v0_max"])]
    upper_y = [float(norm_row["f0_max"]), 0]
    lower_x = [0, float(norm_row["v0_min"])]
    lower_y = [float(norm_row["f0_min"]), 0]
    x_span = max(upper_x[1], player_v0) - min(lower_x[1], 0)
    y_span = max(upper_y[0], player_f0) - min(lower_y[0], 0)
    x_pad = max(x_span * 0.08, 0.15)
    y_pad = max(y_span * 0.08, 0.15)
    x_player = [0, player_v0]
    y_player = [player_f0, 0]

    fig, ax_line = plt.subplots(figsize=(5.8, 3.8))
    ax_line.plot(upper_x, upper_y, label=texts["upper_reference"], linestyle="--", color="#00C060", linewidth=2.0)
    ax_line.plot(x_player, y_player, label=texts["player_label"], color=BLUE_HEX, linewidth=2.5)
    ax_line.plot(lower_x, lower_y, label=texts["lower_reference"], linestyle="--", color="#FB3331", linewidth=2.0)
    ax_line.annotate(str(round(player_v0, 2)), (player_v0, 0), xytext=(player_v0, -0.6), textcoords="data", ha="center", va="center", color="white", bbox=bbox)
    ax_line.annotate(str(round(player_f0, 2)), (0, player_f0), xytext=(-0.6, player_f0), textcoords="data", ha="center", va="center", color="white", bbox=bbox)
    ax_line.spines["top"].set_visible(False)
    ax_line.spines["right"].set_visible(False)
    ax_line.set_xlabel("V0 [m/s]")
    ax_line.set_ylabel("F0 [N/kg]")
    ax_line.set_title(texts["player_fv_profile"], fontsize=10)
    ax_line.set_xlim(-x_pad * 0.2, max(upper_x[1], player_v0) + x_pad)
    ax_line.set_ylim(-y_pad * 0.2, max(upper_y[0], player_f0) + y_pad)
    ax_line.legend(loc="upper right", frameon=False)
    ax_line.margins(x=0.05, y=0.05)

    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


def get_norm_quadrant(report: dict[str, Any], norm_row: dict[str, Any]) -> str:
    f0 = float(report["f0"])
    v0 = float(report["v0"])
    f0_mid = float(norm_row["f0_median"])
    v0_mid = float(norm_row["v0_median"])
    if v0 > v0_mid and f0 > f0_mid:
        return "Q1"
    if v0 <= v0_mid and f0 > f0_mid:
        return "Q2"
    if v0 <= v0_mid and f0 <= f0_mid:
        return "Q3"
    return "Q4"


def get_quadrant_result(quadrant: str, texts: dict[str, Any]) -> str:
    return texts[f"{quadrant.lower()}_result"]


def get_quadrant_recommendations(quadrant: str, texts: dict[str, Any]) -> list[str]:
    return texts[f"{quadrant.lower()}_recs"]


def build_non_normative_fv_pdf(report: dict[str, Any], player_name: str, logo_bytes: bytes | None, player_photo_bytes: bytes | None, texts: dict[str, Any]) -> bytes:
    fv_buf = make_fv_profile_player_only(report)
    pdf = FPDF("L", "mm", "A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    font_family = configure_pdf_font(pdf)
    if logo_bytes:
        pdf.image(io.BytesIO(logo_bytes), x=10, y=12, w=25)

    pdf.set_text_color(*BLACK_RGB)
    pdf.set_font(font_family, "", 32)
    pdf.set_xy(45, 18)
    pdf.cell(0, 10, player_name, new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(128, 128, 128)
    pdf.set_font(font_family, "", 14)
    pdf.set_xy(45, 34)
    pdf.cell(0, 10, texts["fv_title"], new_x="LMARGIN", new_y="NEXT")
    pdf.image(fv_buf, x=20, y=55, w=140)
    if player_photo_bytes:
        pdf.image(io.BytesIO(player_photo_bytes), x=194, y=18, w=45, h=45, keep_aspect_ratio=True)

    data_x = 170
    data_y = 70
    y_second_row = 18
    cell_w = 28
    cell_h = 12
    cell_h_sub = 7
    pdf.set_fill_color(*BLUE_RGB)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_family, "", 12)

    f0 = float(report["f0"])
    v0 = float(report["v0"])
    pmax = float(report["pMax"])
    drf = float(report.get("ratioOfForceDecrease") or 0)
    rfmax = float(report.get("ratioOfForceMax") or 0)
    rounded_corner_cell(pdf, data_x, data_y, cell_w, cell_h, str(round(f0, 2)))
    rounded_corner_cell(pdf, data_x + 32, data_y, cell_w, cell_h, str(round(v0, 2)))
    rounded_corner_cell(pdf, data_x + 64, data_y, cell_w, cell_h, str(round(pmax, 2)))
    rounded_corner_cell(pdf, data_x, data_y + y_second_row, cell_w, cell_h, str(round(v0 * 3.6, 2)))
    rounded_corner_cell(pdf, data_x + 32, data_y + y_second_row, cell_w, cell_h, str(round(drf, 2)))
    rounded_corner_cell(pdf, data_x + 64, data_y + y_second_row, cell_w, cell_h, str(round(rfmax, 2)))
    pdf.set_font(font_family, "", 8)
    pdf.set_text_color(220, 220, 220)
    rounded_corner_cell(pdf, data_x, data_y + 9, cell_w, cell_h_sub, "F0 [N/kg]")
    rounded_corner_cell(pdf, data_x + 32, data_y + 9, cell_w, cell_h_sub, "V0 [m/s]")
    rounded_corner_cell(pdf, data_x + 64, data_y + 9, cell_w, cell_h_sub, "PMax [W]")
    rounded_corner_cell(pdf, data_x, data_y + 9 + y_second_row, cell_w, cell_h_sub, "V0 [km/h]")
    rounded_corner_cell(pdf, data_x + 32, data_y + 9 + y_second_row, cell_w, cell_h_sub, "DRF")
    rounded_corner_cell(pdf, data_x + 64, data_y + 9 + y_second_row, cell_w, cell_h_sub, "RFmax")
    return bytes(pdf.output(dest="S"))


def build_normative_fv_pdf(
    report: dict[str, Any],
    player_name: str,
    logo_bytes: bytes | None,
    player_photo_bytes: bytes | None,
    norm_row: dict[str, Any],
    scatter_entry: dict[str, Any] | None,
    texts: dict[str, Any],
) -> bytes:
    fv_buf = make_normative_fv_profile(report, norm_row, texts)
    scatter_buf = make_norm_scatter_plot(report, norm_row, scatter_entry, texts) if scatter_entry else None
    pdf = FPDF("L", "mm", "A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    font_family = configure_pdf_font(pdf)
    left_x = 14
    left_w = 126
    right_x = 154
    right_w = 129
    logo_w = 32
    top_y = 18
    subtitle_y = 29
    badge_y = 39
    header_bottom_y = badge_y + 11
    logo_x = pdf.w - logo_w - 12
    logo_y = top_y + ((header_bottom_y - top_y) - logo_w) / 2
    identity_x = left_x
    fv_chart_y = 56
    fv_chart_w = left_w
    metrics_y = 148
    rec_title_y = 154
    rec_text_y = 161
    photo_w = 46
    photo_x = right_x + (right_w - photo_w) / 2
    photo_y = 14
    scatter_y = 66
    scatter_w = 96
    scatter_x = right_x + (right_w - scatter_w) / 2
    rs_logo_w = 36
    rs_logo_x = pdf.w - rs_logo_w - 10
    rs_logo_y = pdf.h - 15
    if logo_bytes:
        pdf.image(io.BytesIO(logo_bytes), x=logo_x, y=logo_y, w=logo_w)
    pdf.set_text_color(*BLACK_RGB)
    pdf.set_font(font_family, "", 24)
    pdf.set_xy(identity_x, top_y)
    pdf.cell(0, 10, player_name, new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(128, 128, 128)
    pdf.set_font(font_family, "", 12)
    pdf.set_xy(identity_x, subtitle_y)
    pdf.cell(0, 10, texts["fv_title"], new_x="LMARGIN", new_y="NEXT")

    quadrant = get_norm_quadrant(report, norm_row)
    quadrant_result = get_quadrant_result(quadrant, texts)
    recommendations = get_quadrant_recommendations(quadrant, texts)
    badge_x = identity_x
    pdf.set_font(font_family, "", 12)
    pdf.set_fill_color(*get_quadrant_badge_fill(quadrant))
    pdf.set_text_color(255, 255, 255)
    rounded_corner_cell(pdf, badge_x, badge_y, 14, 11, quadrant)
    pdf.set_fill_color(*get_quadrant_result_fill(quadrant))
    rounded_corner_cell(pdf, badge_x + 16, badge_y, min(100, left_w - 16), 11, quadrant_result)

    pdf.image(fv_buf, x=left_x, y=fv_chart_y, w=fv_chart_w)
    if player_photo_bytes:
        pdf.image(io.BytesIO(player_photo_bytes), x=photo_x, y=photo_y, w=photo_w, h=photo_w, keep_aspect_ratio=True)
    if scatter_buf:
        pdf.image(scatter_buf, x=scatter_x, y=scatter_y, w=scatter_w)

    f0 = float(report["f0"])
    v0 = float(report["v0"])
    pmax = float(report["pMax"])
    drf = float(report.get("ratioOfForceDecrease") or 0)
    rfmax = float(report.get("ratioOfForceMax") or 0)
    metrics_block_w = 74
    data_x = left_x + (left_w - metrics_block_w) / 2
    data_y = metrics_y
    y_second_row = 18
    cell_w = 22
    cell_h = 12
    cell_h_sub = 7
    pdf.set_fill_color(*BLUE_RGB)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_family, "", 12)
    rounded_corner_cell(pdf, data_x, data_y, cell_w, cell_h, str(round(f0, 2)))
    rounded_corner_cell(pdf, data_x + 26, data_y, cell_w, cell_h, str(round(v0, 2)))
    rounded_corner_cell(pdf, data_x + 52, data_y, cell_w, cell_h, str(round(pmax, 2)))
    rounded_corner_cell(pdf, data_x, data_y + y_second_row, cell_w, cell_h, str(round(v0 * 3.6, 2)))
    rounded_corner_cell(pdf, data_x + 26, data_y + y_second_row, cell_w, cell_h, str(round(drf, 2)))
    rounded_corner_cell(pdf, data_x + 52, data_y + y_second_row, cell_w, cell_h, str(round(rfmax, 2)))
    pdf.set_font(font_family, "", 8)
    pdf.set_text_color(220, 220, 220)
    rounded_corner_cell(pdf, data_x, data_y + 9, cell_w, cell_h_sub, "F0 [N/kg]")
    rounded_corner_cell(pdf, data_x + 26, data_y + 9, cell_w, cell_h_sub, "V0 [m/s]")
    rounded_corner_cell(pdf, data_x + 52, data_y + 9, cell_w, cell_h_sub, "PMax [W]")
    rounded_corner_cell(pdf, data_x, data_y + 9 + y_second_row, cell_w, cell_h_sub, "V0 [km/h]")
    rounded_corner_cell(pdf, data_x + 26, data_y + 9 + y_second_row, cell_w, cell_h_sub, "DRF")
    rounded_corner_cell(pdf, data_x + 52, data_y + 9 + y_second_row, cell_w, cell_h_sub, "RFmax")
    pdf.set_font(font_family, "", 14)
    pdf.set_text_color(*BLACK_RGB)
    pdf.text(right_x, rec_title_y, texts["recommendation"])
    pdf.set_font(font_family, "", 8)
    y_coordinate = rec_text_y
    for item in recommendations:
        pdf.text(right_x, y_coordinate, "* " + item)
        y_coordinate += 7
    if RS_LOGO_PATH.is_file():
        pdf.image(str(RS_LOGO_PATH), x=rs_logo_x, y=rs_logo_y, w=rs_logo_w)
    return bytes(pdf.output(dest="S"))
