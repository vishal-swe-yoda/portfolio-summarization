"""
PowerPointRenderer: Declarative slide renderer using python-pptx.
Implements global chrome, typography floor (>=10pt), custom XML cell borders,
and layout renderers: three_table_dashboard, sector_analysis, and holdings_dashboard.
"""

from typing import Dict, Any, List, Optional
import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn


def hex_to_rgb(hex_str: str) -> RGBColor:
    """Converts a hex color string like '#1B4F9C' or '1B4F9C' to RGBColor."""
    hex_str = hex_str.lstrip("#")
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


class PowerPointRenderer:
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.report_cfg = config.get("report", {})
        theme_name = self.report_cfg.get("theme", "uti_blue")
        themes = config.get("themes", {})
        self.theme = themes.get(theme_name, themes.get("uti_blue", {}))

        # Color palette
        self.banner_blue = hex_to_rgb(self.theme.get("banner_blue", "#1B4F9C"))
        self.header_blue = hex_to_rgb(self.theme.get("header_blue", "#004C97"))
        self.chart_navy = hex_to_rgb(self.theme.get("chart_navy", "#002060"))
        self.positive_green = hex_to_rgb(self.theme.get("positive_green", "#00B050"))
        self.negative_red = hex_to_rgb(self.theme.get("negative_red", "#FF0000"))
        self.subheader_grey = hex_to_rgb(self.theme.get("subheader_grey", "#D9D9D9"))
        self.zebra_grey = hex_to_rgb(self.theme.get("zebra_grey", "#E7E6E6"))
        self.accent_cyan = hex_to_rgb(self.theme.get("accent_cyan", "#0072BC"))
        self.accent_orange = hex_to_rgb(self.theme.get("accent_orange", "#F58220"))
        self.text_dark = hex_to_rgb(self.theme.get("text_dark", "#000000"))
        self.text_light = hex_to_rgb(self.theme.get("text_light", "#FFFFFF"))
        self.font_family = self.theme.get("font_family", "Calibri")

        self.logo_path = self.report_cfg.get("logo_path", "assets/uti_logo.png")
        self.prs = Presentation()
        # Widescreen 16:9
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.50)
        self.blank_layout = self.prs.slide_layouts[6]

    # -------------------------------------------------------------------------
    # Typography & XML Cell Border Helpers
    # -------------------------------------------------------------------------
    def format_cell(
        self,
        cell,
        text: str,
        font_size: Pt = Pt(10),
        bold: bool = False,
        color: Optional[RGBColor] = None,
        align: PP_ALIGN = PP_ALIGN.LEFT,
        bg_color: Optional[RGBColor] = None,
        border_color: str = "D0D0D0",
        border_width: str = "12700"
    ):
        """Formats cell text ensuring font size >= 10pt and sets background and borders."""
        if bg_color:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_color
        else:
            cell.fill.background()

        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.06)
        cell.margin_right = Inches(0.06)
        cell.margin_top = Inches(0.03)
        cell.margin_bottom = Inches(0.03)

        tf = cell.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = align

        # Ensure typography floor >= 10pt
        eff_size = max(font_size, Pt(10))
        run = p.add_run()
        run.text = str(text)
        run.font.name = self.font_family
        run.font.size = eff_size
        run.font.bold = bold
        run.font.color.rgb = color if color else self.text_dark

        # Set custom XML borders
        self.set_cell_border(cell, color=border_color, width=border_width)

    def set_cell_border(
        self,
        cell,
        color: str = "D0D0D0",
        width: str = "12700",
        top: bool = True,
        bottom: bool = True,
        left: bool = True,
        right: bool = True
    ):
        """Applies crisp drawingml borders to a table cell."""
        tcPr = cell._tc.get_or_add_tcPr()
        borders = {
            'a:lnT': top,
            'a:lnB': bottom,
            'a:lnL': left,
            'a:lnR': right
        }
        for tag, apply in borders.items():
            existing = tcPr.find(qn(tag))
            if existing is not None:
                tcPr.remove(existing)
            if apply:
                xml = f'<{tag} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width}" cmpd="s"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></{tag}>'
                element = parse_xml(xml)
                tcPr.append(element)

    # -------------------------------------------------------------------------
    # Shared Slide Chrome
    # -------------------------------------------------------------------------
    def add_slide_chrome(
        self,
        slide,
        title: str,
        slide_number: int,
        footer_text: str,
        is_page_1: bool = False
    ):
        """Applies top banner, logo, slide number, footers, and bottom accent strip."""
        shapes = slide.shapes

        # 1. Top Banner (Blue rectangle)
        rect = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.25), Inches(0.18), Inches(10.65), Inches(0.48))
        rect.fill.solid()
        rect.fill.fore_color.rgb = self.banner_blue
        rect.line.fill.background()

        tf = rect.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = f"  {title}"
        run.font.name = self.font_family
        run.font.size = Pt(21)
        run.font.bold = True
        run.font.color.rgb = self.text_light

        # Slanted parallelogram cap
        cap = shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(10.75), Inches(0.18), Inches(0.32), Inches(0.48))
        cap.fill.solid()
        cap.fill.fore_color.rgb = self.banner_blue
        cap.line.fill.background()

        # 2. Top-Right Brand Logo
        if os.path.exists(self.logo_path):
            shapes.add_picture(self.logo_path, Inches(11.75), Inches(0.12), width=Inches(1.34))
        else:
            # Fallback text logo badge
            logo_box = shapes.add_textbox(Inches(11.65), Inches(0.16), Inches(1.45), Inches(0.48))
            tf_logo = logo_box.text_frame
            p_logo = tf_logo.paragraphs[0]
            p_logo.alignment = PP_ALIGN.RIGHT
            run_logo = p_logo.add_run()
            run_logo.text = "UTI MUTUAL FUND"
            run_logo.font.name = self.font_family
            run_logo.font.size = Pt(11)
            run_logo.font.bold = True
            run_logo.font.color.rgb = self.banner_blue

        # 3. Slide Number (Bottom Left)
        num_box = shapes.add_textbox(Inches(0.18), Inches(6.95), Inches(0.35), Inches(0.32))
        tf_num = num_box.text_frame
        p_num = tf_num.paragraphs[0]
        run_num = p_num.add_run()
        run_num.text = str(slide_number)
        run_num.font.name = self.font_family
        run_num.font.size = Pt(11)
        run_num.font.bold = True
        run_num.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

        # 4. Footer Textbox
        footer_y = Inches(6.55) if is_page_1 else Inches(6.95)
        footer_h = Inches(0.75) if is_page_1 else Inches(0.35)
        foot_box = shapes.add_textbox(Inches(0.58), footer_y, Inches(12.50), footer_h)
        tf_foot = foot_box.text_frame
        tf_foot.word_wrap = True
        lines = footer_text.split("\n")
        for i, line in enumerate(lines):
            p_f = tf_foot.paragraphs[0] if i == 0 else tf_foot.add_paragraph()
            run_f = p_f.add_run()
            run_f.text = line
            run_f.font.name = self.font_family
            run_f.font.size = Pt(10)
            run_f.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        # 5. Bottom Tri-Color Accent Strip at y=7.36", h=0.14"
        y_strip = Inches(7.36)
        h_strip = Inches(0.14)
        
        # Segment 1: Navy
        s1 = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.00), y_strip, Inches(6.20), h_strip)
        s1.fill.solid()
        s1.fill.fore_color.rgb = self.banner_blue
        s1.line.fill.background()

        # Segment 2: Cyan
        s2 = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.20), y_strip, Inches(3.05), h_strip)
        s2.fill.solid()
        s2.fill.fore_color.rgb = self.accent_cyan
        s2.line.fill.background()

        # Segment 3: Orange
        s3 = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.25), y_strip, Inches(4.083), h_strip)
        s3.fill.solid()
        s3.fill.fore_color.rgb = self.accent_orange
        s3.line.fill.background()

    # -------------------------------------------------------------------------
    # Layout 1: three_table_dashboard (Portfolio Statistics)
    # -------------------------------------------------------------------------
    def render_three_table_dashboard(self, page_cfg: Dict[str, Any], data_bundle: Dict[str, Any], slide_num: int):
        slide = self.prs.slides.add_slide(self.blank_layout)
        title = page_cfg.get("title", "Portfolio - Statistics")
        disp_date = self.report_cfg.get("display_date_page1", "30/Mar/2026")
        bm = self.report_cfg.get("benchmark", "BSE 100 TRI")

        footer_text = (
            f"Source: Bloomberg, BSE, Internal Portfolio Analytics as of {disp_date}. Quants calculated using {bm} as benchmark.\n"
            "Operating Cash Flow (OCF) Tiers: C1 = OCF/EBITDA > 75%, C2 = 50-75%, C3 < 50%. RoCE Tiers: R1 > 18%, R2 = 10-18%, R3 < 10%."
        )
        self.add_slide_chrome(slide, title, slide_num, footer_text, is_page_1=True)

        shapes = slide.shapes

        # 1. Quants & Ratios Table (0.35", 1.02", 4.15", 5.00") - 9 x 2 grid
        quants = data_bundle.get("quants_ratios", [])
        num_rows = max(len(quants) + 1, 9)
        t_quants = shapes.add_table(num_rows, 2, Inches(0.35), Inches(1.02), Inches(4.15), Inches(5.00)).table
        t_quants.columns[0].width = Inches(2.45)
        t_quants.columns[1].width = Inches(1.70)

        # Merge Header Row 0
        t_quants.cell(0, 0).merge(t_quants.cell(0, 1))
        self.format_cell(t_quants.cell(0, 0), "Quants & Ratios", font_size=Pt(14), bold=True, color=self.text_light, align=PP_ALIGN.CENTER, bg_color=self.header_blue)

        for i in range(1, num_rows):
            bg = self.zebra_grey if i % 2 == 1 else self.text_light
            if i - 1 < len(quants):
                q = quants[i - 1]
                lbl = q.get("metric", "")
                val = str(q.get("display_value", ""))
                self.format_cell(t_quants.cell(i, 0), lbl, font_size=Pt(12), bold=True, color=self.text_dark, align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_quants.cell(i, 1), val, font_size=Pt(12), bold=False, color=self.text_dark, align=PP_ALIGN.CENTER, bg_color=bg)
            else:
                self.format_cell(t_quants.cell(i, 0), "", bg_color=bg)
                self.format_cell(t_quants.cell(i, 1), "", bg_color=bg)

        # 2. Market Capitalization Table (4.80", 1.02", 3.45", 2.45") - 5 x 3 grid
        mc_rows = data_bundle.get("market_cap_summary", [])
        t_mc = shapes.add_table(5, 3, Inches(4.80), Inches(1.02), Inches(3.45), Inches(2.45)).table
        t_mc.columns[0].width = Inches(1.35)
        t_mc.columns[1].width = Inches(1.05)
        t_mc.columns[2].width = Inches(1.05)

        # Merge Header Row 0
        t_mc.cell(0, 0).merge(t_mc.cell(0, 2))
        self.format_cell(t_mc.cell(0, 0), "Market Capitalization", font_size=Pt(14), bold=True, color=self.text_light, align=PP_ALIGN.CENTER, bg_color=self.header_blue)

        # Subheader Row 1
        self.format_cell(t_mc.cell(1, 0), "Segment", font_size=Pt(12), bold=True, color=self.text_dark, align=PP_ALIGN.LEFT, bg_color=self.subheader_grey)
        self.format_cell(t_mc.cell(1, 1), "Fund", font_size=Pt(12), bold=True, color=self.text_dark, align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)
        self.format_cell(t_mc.cell(1, 2), "BM", font_size=Pt(12), bold=True, color=self.text_dark, align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)

        for i in range(2, 5):
            bg = self.zebra_grey if i % 2 == 0 else self.text_light
            if i - 2 < len(mc_rows):
                m = mc_rows[i - 2]
                self.format_cell(t_mc.cell(i, 0), m.get("segment", ""), font_size=Pt(11.5), bold=True, align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_mc.cell(i, 1), m.get("fund", ""), font_size=Pt(11.5), align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_mc.cell(i, 2), m.get("bm", ""), font_size=Pt(11.5), align=PP_ALIGN.CENTER, bg_color=bg)

        # 3. Valuation Parameters Table (4.80", 3.68", 3.45", 2.34") - 4 x 3 grid
        val_rows = data_bundle.get("valuation_metrics", [])
        t_val = shapes.add_table(4, 3, Inches(4.80), Inches(3.68), Inches(3.45), Inches(2.34)).table
        t_val.columns[0].width = Inches(1.65)
        t_val.columns[1].width = Inches(0.90)
        t_val.columns[2].width = Inches(0.90)

        self.format_cell(t_val.cell(0, 0), "Parameters", font_size=Pt(13), bold=True, color=self.text_light, align=PP_ALIGN.LEFT, bg_color=self.header_blue)
        self.format_cell(t_val.cell(0, 1), "Fund", font_size=Pt(13), bold=True, color=self.text_light, align=PP_ALIGN.CENTER, bg_color=self.header_blue)
        self.format_cell(t_val.cell(0, 2), "BM", font_size=Pt(13), bold=True, color=self.text_light, align=PP_ALIGN.CENTER, bg_color=self.header_blue)

        for i in range(1, 4):
            bg = self.zebra_grey if i % 2 == 1 else self.text_light
            if i - 1 < len(val_rows):
                v = val_rows[i - 1]
                self.format_cell(t_val.cell(i, 0), v.get("parameter", ""), font_size=Pt(11.5), bold=True, align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_val.cell(i, 1), v.get("fund", ""), font_size=Pt(11.5), align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_val.cell(i, 2), v.get("bm", ""), font_size=Pt(11.5), align=PP_ALIGN.CENTER, bg_color=bg)

        # 4. Portfolio Composition Table (8.55", 1.02", 4.43", 2.50") - 6 x 6 grid
        comp = data_bundle.get("composition_tiers", {})
        ocf_data = comp.get("ocf", [])
        roce_data = comp.get("roce", [])
        t_comp = shapes.add_table(6, 6, Inches(8.55), Inches(1.02), Inches(4.43), Inches(2.50)).table
        col_widths = [0.69, 0.77, 0.75, 0.69, 0.77, 0.76]
        for idx, w in enumerate(col_widths):
            t_comp.columns[idx].width = Inches(w)

        # Row 0: Merged Portfolio Composition (0..5)
        t_comp.cell(0, 0).merge(t_comp.cell(0, 5))
        self.format_cell(t_comp.cell(0, 0), "Portfolio Composition", font_size=Pt(14), bold=True, color=self.text_light, align=PP_ALIGN.CENTER, bg_color=self.header_blue)

        # Row 1: Merged OCF (0..2) and RoCE (3..5)
        t_comp.cell(1, 0).merge(t_comp.cell(1, 2))
        t_comp.cell(1, 3).merge(t_comp.cell(1, 5))
        self.format_cell(t_comp.cell(1, 0), "Operating Cash Flow (OCF)", font_size=Pt(11.5), bold=True, color=self.text_dark, align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)
        self.format_cell(t_comp.cell(1, 3), "Return on Capital Employed", font_size=Pt(11.5), bold=True, color=self.text_dark, align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)

        # Row 2: Subheaders ["Tier", "Fund", "BM", "Tier", "Fund", "BM"]
        sub_headers = ["Tier", "Fund", "BM", "Tier", "Fund", "BM"]
        for c_idx, sh in enumerate(sub_headers):
            self.format_cell(t_comp.cell(2, c_idx), sh, font_size=Pt(11), bold=True, color=self.text_dark, align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)

        # Rows 3-5: C1..C3 and R1..R3
        for i in range(3, 6):
            bg = self.zebra_grey if i % 2 == 1 else self.text_light
            # OCF
            ocf_idx = i - 3
            if ocf_idx < len(ocf_data):
                o = ocf_data[ocf_idx]
                self.format_cell(t_comp.cell(i, 0), o.get("tier", ""), font_size=Pt(11), bold=True, align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_comp.cell(i, 1), o.get("fund", ""), font_size=Pt(11), align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_comp.cell(i, 2), o.get("bm", ""), font_size=Pt(11), align=PP_ALIGN.CENTER, bg_color=bg)
            # RoCE
            roce_idx = i - 3
            if roce_idx < len(roce_data):
                r = roce_data[roce_idx]
                self.format_cell(t_comp.cell(i, 3), r.get("tier", ""), font_size=Pt(11), bold=True, align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_comp.cell(i, 4), r.get("fund", ""), font_size=Pt(11), align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_comp.cell(i, 5), r.get("bm", ""), font_size=Pt(11), align=PP_ALIGN.CENTER, bg_color=bg)

    # -------------------------------------------------------------------------
    # Layout 2: sector_analysis (Sector Allocation & LLM Commentary)
    # -------------------------------------------------------------------------
    def render_sector_analysis(
        self,
        page_cfg: Dict[str, Any],
        data_bundle: Dict[str, Any],
        llm_response: Dict[str, Any],
        chart_path: str,
        slide_num: int
    ):
        slide = self.prs.slides.add_slide(self.blank_layout)
        title = page_cfg.get("title", "Portfolio - Sectoral Allocation")
        disp_date = self.report_cfg.get("display_date_page2", "March 30, 2026")
        bm = self.report_cfg.get("benchmark", "BSE 100 TRI")

        footer_text = f"Source: Internal Portfolio Analytics, FactSet, {bm} as of {disp_date}. Active Weight = Fund Weight - Benchmark Weight."
        self.add_slide_chrome(slide, title, slide_num, footer_text, is_page_1=False)

        shapes = slide.shapes

        # 1. Left Section Banner (0.25", 0.80", 6.20", 0.42")
        b_left = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.25), Inches(0.80), Inches(6.20), Inches(0.42))
        b_left.fill.solid()
        b_left.fill.fore_color.rgb = self.banner_blue
        b_left.line.fill.background()
        tf = b_left.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = " Sectoral Breakdown & Active Position (%)"
        run.font.name = self.font_family
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = self.text_light

        # 2. Left Sector Chart (0.25", 1.28", 6.20", 5.38")
        if os.path.exists(chart_path):
            pic = shapes.add_picture(chart_path, Inches(0.25), Inches(1.28), width=Inches(6.20), height=Inches(5.38))
            pic.line.color.rgb = RGBColor(0x88, 0x88, 0x88)
            pic.line.width = Pt(0.75)

        # 3. Upper-Right Banner: Overweight (Top 3) at (6.62", 0.80", 6.45", 0.42")
        b_ow = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.62), Inches(0.80), Inches(6.45), Inches(0.42))
        b_ow.fill.solid()
        b_ow.fill.fore_color.rgb = self.banner_blue
        b_ow.line.fill.background()
        tf_ow = b_ow.text_frame
        p_ow = tf_ow.paragraphs[0]
        run_ow = p_ow.add_run()
        run_ow.text = " Overweight (Top 3)"
        run_ow.font.name = self.font_family
        run_ow.font.size = Pt(13)
        run_ow.font.bold = True
        run_ow.font.color.rgb = self.text_light

        # 4. Overweight Table (6.62", 1.28", 6.45", 2.55") - 4 x 2 table
        ow_data = llm_response.get("overweight_commentary", [])
        t_ow = shapes.add_table(4, 2, Inches(6.62), Inches(1.28), Inches(6.45), Inches(2.55)).table
        t_ow.columns[0].width = Inches(1.85)
        t_ow.columns[1].width = Inches(4.60)

        # Row 0 Header
        self.format_cell(t_ow.cell(0, 0), "Sector (%NAV / %AW)", font_size=Pt(11), bold=True, color=self.text_light, align=PP_ALIGN.CENTER, bg_color=self.banner_blue)
        self.format_cell(t_ow.cell(0, 1), "Sector Overview", font_size=Pt(11), bold=True, color=self.text_light, align=PP_ALIGN.LEFT, bg_color=self.banner_blue)

        for i in range(1, 4):
            bg = self.zebra_grey if i % 2 == 1 else self.text_light
            if i - 1 < len(ow_data):
                item = ow_data[i - 1]
                sec_text = f"{item.get('sector', '')}\n({item.get('fund_weight', 0):.1f}% / +{item.get('active_weight', 0):.1f}%)"
                self.format_cell(t_ow.cell(i, 0), sec_text, font_size=Pt(10.5), bold=True, color=self.text_dark, align=PP_ALIGN.CENTER, bg_color=bg)

                # Col 1: Observations bullet list
                obs_list = item.get("observations", [])
                self._format_bullet_cell(t_ow.cell(i, 1), obs_list, bg_color=bg)
            else:
                self.format_cell(t_ow.cell(i, 0), "", bg_color=bg)
                self.format_cell(t_ow.cell(i, 1), "", bg_color=bg)

        # 5. Lower-Right Banner: Underweight (Top 3) at (6.62", 3.96", 6.45", 0.42")
        b_uw = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.62), Inches(3.96), Inches(6.45), Inches(0.42))
        b_uw.fill.solid()
        b_uw.fill.fore_color.rgb = self.banner_blue
        b_uw.line.fill.background()
        tf_uw = b_uw.text_frame
        p_uw = tf_uw.paragraphs[0]
        run_uw = p_uw.add_run()
        run_uw.text = " Underweight (Top 3)"
        run_uw.font.name = self.font_family
        run_uw.font.size = Pt(13)
        run_uw.font.bold = True
        run_uw.font.color.rgb = self.text_light

        # 6. Underweight Table (6.62", 4.44", 6.45", 2.45") - 4 x 2 table
        uw_data = llm_response.get("underweight_commentary", [])
        t_uw = shapes.add_table(4, 2, Inches(6.62), Inches(4.44), Inches(6.45), Inches(2.45)).table
        t_uw.columns[0].width = Inches(1.85)
        t_uw.columns[1].width = Inches(4.60)

        # Row 0 Header
        self.format_cell(t_uw.cell(0, 0), "Sector (%NAV / %AW)", font_size=Pt(11), bold=True, color=self.text_light, align=PP_ALIGN.CENTER, bg_color=self.banner_blue)
        self.format_cell(t_uw.cell(0, 1), "Sector Overview", font_size=Pt(11), bold=True, color=self.text_light, align=PP_ALIGN.LEFT, bg_color=self.banner_blue)

        for i in range(1, 4):
            bg = self.zebra_grey if i % 2 == 1 else self.text_light
            if i - 1 < len(uw_data):
                item = uw_data[i - 1]
                sec_text = f"{item.get('sector', '')}\n({item.get('fund_weight', 0):.1f}% / {item.get('active_weight', 0):.1f}%)"
                self.format_cell(t_uw.cell(i, 0), sec_text, font_size=Pt(10.5), bold=True, color=self.text_dark, align=PP_ALIGN.CENTER, bg_color=bg)

                obs_list = item.get("observations", [])
                self._format_bullet_cell(t_uw.cell(i, 1), obs_list, bg_color=bg)
            else:
                self.format_cell(t_uw.cell(i, 0), "", bg_color=bg)
                self.format_cell(t_uw.cell(i, 1), "", bg_color=bg)

    def _format_bullet_cell(self, cell, bullets: List[str], bg_color: Optional[RGBColor] = None):
        """Formats cell with bullet points ensuring >= 10pt font."""
        if bg_color:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_color
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.06)
        cell.margin_top = Inches(0.04)
        cell.margin_bottom = Inches(0.04)

        tf = cell.text_frame
        tf.word_wrap = True
        tf.clear()

        if not bullets:
            bullets = ["No specific notes recorded."]

        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = f"•  {b}"
            run.font.name = self.font_family
            run.font.size = Pt(10)
            run.font.color.rgb = self.text_dark

        self.set_cell_border(cell, color="D0D0D0", width="12700")

    # -------------------------------------------------------------------------
    # Layout 3: holdings_dashboard (Top 10 Holdings & Investment Insights)
    # -------------------------------------------------------------------------
    def render_holdings_dashboard(
        self,
        page_cfg: Dict[str, Any],
        top_holdings_summary: Optional[pd.DataFrame] = None,
        top_overweight_stocks: Optional[pd.DataFrame] = None,
        top_underweight_stocks: Optional[pd.DataFrame] = None,
        slide_num: int = 3,
        kpi_callout_text: Optional[str] = None,
        data_bundle: Optional[Dict[str, Any]] = None,
        llm_response: Optional[Dict[str, Any]] = None,
    ):
        slide = self.prs.slides.add_slide(self.blank_layout)
        title = page_cfg.get("title", "Portfolio - Stock Holdings")
        disp_date = self.report_cfg.get("display_date_page2", "March 30, 2026")
        bm = self.report_cfg.get("benchmark", "BSE 100 TRI")

        footer_text = f"Source: Internal Portfolio Analytics, {bm} as of {disp_date}."
        self.add_slide_chrome(slide, title, slide_num, footer_text, is_page_1=False)

        shapes = slide.shapes
        sections = page_cfg.get("sections", [])

        # Fallback to data_bundle if dataframes were passed via bundle
        if top_holdings_summary is None and data_bundle:
            top_holdings_summary = data_bundle.get("top_holdings_summary")
        if top_overweight_stocks is None and data_bundle:
            top_overweight_stocks = data_bundle.get("top_overweight_stocks")
        if top_underweight_stocks is None and data_bundle:
            top_underweight_stocks = data_bundle.get("top_underweight_stocks")

        # -------------------------------------------------------------
        # 1. Left Section: Top 15 Holdings (0.25", 0.80", 6.20", 0.38")
        # -------------------------------------------------------------
        sec_h = next((s for s in sections if s.get("id") == "top_holdings"), {})
        title_h = sec_h.get("title", "Top 15 Holdings")
        limit_h = sec_h.get("limit", 15)

        b_h = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.25), Inches(0.80), Inches(6.20), Inches(0.38))
        b_h.fill.solid()
        b_h.fill.fore_color.rgb = self.banner_blue
        b_h.line.fill.background()
        tf_h = b_h.text_frame
        p_h = tf_h.paragraphs[0]
        run_h = p_h.add_run()
        run_h.text = f" {title_h}"
        run_h.font.name = self.font_family
        run_h.font.size = Pt(13)
        run_h.font.bold = True
        run_h.font.color.rgb = self.text_light

        # Left Table: (Header + limit_h + Total row) = limit_h + 2 rows
        num_h_rows = limit_h + 2
        t_h = shapes.add_table(num_h_rows, 5, Inches(0.25), Inches(1.22), Inches(6.20), Inches(5.60)).table
        t_h.columns[0].width = Inches(2.25)
        t_h.columns[1].width = Inches(1.45)
        t_h.columns[2].width = Inches(0.80)
        t_h.columns[3].width = Inches(0.80)
        t_h.columns[4].width = Inches(0.90)

        h_cols = ["Stock Name", "Sector", "Fund %", "BM %", "Active %"]
        for c_idx, hc in enumerate(h_cols):
            self.format_cell(t_h.cell(0, c_idx), hc, font_size=Pt(10.5), bold=True, color=self.text_light, align=PP_ALIGN.CENTER if c_idx >= 2 else PP_ALIGN.LEFT, bg_color=self.banner_blue)

        df_top_h = top_holdings_summary if top_holdings_summary is not None else pd.DataFrame()
        sum_fund = 0.0
        sum_bm = 0.0
        for i in range(1, limit_h + 1):
            bg = self.zebra_grey if i % 2 == 1 else self.text_light
            if i - 1 < len(df_top_h):
                r = df_top_h.iloc[i - 1]
                sname = str(r.get("Stock_Name", r.get("securityname", "")))
                sec = str(r.get("Sector", r.get("sector", "")))
                fw = float(r.get("Fund_Weight", 0.0))
                bw = float(r.get("BM_Weight", 0.0))
                aw = float(r.get("Active_Weight", fw - bw))
                sum_fund += fw
                sum_bm += bw

                self.format_cell(t_h.cell(i, 0), sname, font_size=Pt(9.5), bold=True, align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_h.cell(i, 1), sec, font_size=Pt(9.5), align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_h.cell(i, 2), f"{fw:.2f}%", font_size=Pt(9.5), align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_h.cell(i, 3), f"{bw:.2f}%" if bw > 0 else "-", font_size=Pt(9.5), align=PP_ALIGN.CENTER, bg_color=bg)

                aw_color = self.positive_green if aw >= 0 else self.negative_red
                aw_sign = "+" if aw > 0 else ""
                self.format_cell(t_h.cell(i, 4), f"{aw_sign}{aw:.2f}%", font_size=Pt(9.5), bold=True, color=aw_color, align=PP_ALIGN.CENTER, bg_color=bg)
            else:
                for c in range(5):
                    self.format_cell(t_h.cell(i, c), "", bg_color=bg)

        # Row limit_h + 1: Total Row
        tot_row_idx = limit_h + 1
        sum_aw = sum_fund - sum_bm
        sum_aw_color = self.positive_green if sum_aw >= 0 else self.negative_red
        sum_aw_sign = "+" if sum_aw > 0 else ""
        self.format_cell(t_h.cell(tot_row_idx, 0), f"Top {limit_h} Total", font_size=Pt(10), bold=True, align=PP_ALIGN.LEFT, bg_color=self.subheader_grey)
        self.format_cell(t_h.cell(tot_row_idx, 1), "-", font_size=Pt(10), align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)
        self.format_cell(t_h.cell(tot_row_idx, 2), f"{sum_fund:.2f}%", font_size=Pt(10), bold=True, align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)
        self.format_cell(t_h.cell(tot_row_idx, 3), f"{sum_bm:.2f}%", font_size=Pt(10), bold=True, align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)
        self.format_cell(t_h.cell(tot_row_idx, 4), f"{sum_aw_sign}{sum_aw:.2f}%", font_size=Pt(10), bold=True, color=sum_aw_color, align=PP_ALIGN.CENTER, bg_color=self.subheader_grey)

        # -------------------------------------------------------------
        # 2. Upper-Right Section: Overweight (Top 5)
        # -------------------------------------------------------------
        sec_ow = next((s for s in sections if s.get("id") == "top_overweight"), {})
        title_ow = sec_ow.get("title", "Overweight (Top 5)")
        limit_ow = sec_ow.get("limit", 5)

        b_ow = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.62), Inches(0.80), Inches(6.45), Inches(0.38))
        b_ow.fill.solid()
        b_ow.fill.fore_color.rgb = self.banner_blue
        b_ow.line.fill.background()
        tf_ow = b_ow.text_frame
        p_ow = tf_ow.paragraphs[0]
        run_ow = p_ow.add_run()
        run_ow.text = f" {title_ow}"
        run_ow.font.name = self.font_family
        run_ow.font.size = Pt(13)
        run_ow.font.bold = True
        run_ow.font.color.rgb = self.text_light

        t_ow = shapes.add_table(limit_ow + 1, 4, Inches(6.62), Inches(1.22), Inches(6.45), Inches(2.15)).table
        t_ow.columns[0].width = Inches(2.40)
        t_ow.columns[1].width = Inches(1.75)
        t_ow.columns[2].width = Inches(1.15)
        t_ow.columns[3].width = Inches(1.15)

        ow_cols = ["Stock Name", "Sector", "Fund %", "Active %"]
        for c_idx, oc in enumerate(ow_cols):
            self.format_cell(t_ow.cell(0, c_idx), oc, font_size=Pt(10.5), bold=True, color=self.text_light, align=PP_ALIGN.CENTER if c_idx >= 2 else PP_ALIGN.LEFT, bg_color=self.banner_blue)

        df_ow = top_overweight_stocks if top_overweight_stocks is not None else pd.DataFrame()
        for i in range(1, limit_ow + 1):
            bg = self.zebra_grey if i % 2 == 1 else self.text_light
            if i - 1 < len(df_ow):
                r = df_ow.iloc[i - 1]
                sname = str(r.get("Stock_Name", r.get("securityname", "")))
                sec = str(r.get("Sector", r.get("sector", "")))
                fw = float(r.get("Fund_Weight", 0.0))
                aw = float(r.get("Active_Weight", 0.0))

                self.format_cell(t_ow.cell(i, 0), sname, font_size=Pt(9.5), bold=True, align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_ow.cell(i, 1), sec, font_size=Pt(9.5), align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_ow.cell(i, 2), f"{fw:.2f}%", font_size=Pt(9.5), align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_ow.cell(i, 3), f"+{aw:.2f}%", font_size=Pt(9.5), bold=True, color=self.positive_green, align=PP_ALIGN.CENTER, bg_color=bg)
            else:
                for c in range(4):
                    self.format_cell(t_ow.cell(i, c), "", bg_color=bg)

        # -------------------------------------------------------------
        # 3. Lower-Right Section: Underweight (Top 5)
        # -------------------------------------------------------------
        sec_uw = next((s for s in sections if s.get("id") == "top_underweight"), {})
        title_uw = sec_uw.get("title", "Underweight (Top 5)")
        limit_uw = sec_uw.get("limit", 5)

        b_uw = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.62), Inches(3.50), Inches(6.45), Inches(0.38))
        b_uw.fill.solid()
        b_uw.fill.fore_color.rgb = self.banner_blue
        b_uw.line.fill.background()
        tf_uw = b_uw.text_frame
        p_uw = tf_uw.paragraphs[0]
        run_uw = p_uw.add_run()
        run_uw.text = f" {title_uw}"
        run_uw.font.name = self.font_family
        run_uw.font.size = Pt(13)
        run_uw.font.bold = True
        run_uw.font.color.rgb = self.text_light

        t_uw = shapes.add_table(limit_uw + 1, 4, Inches(6.62), Inches(3.92), Inches(6.45), Inches(2.15)).table
        t_uw.columns[0].width = Inches(2.40)
        t_uw.columns[1].width = Inches(1.75)
        t_uw.columns[2].width = Inches(1.15)
        t_uw.columns[3].width = Inches(1.15)

        uw_cols = ["Stock Name", "Sector", "Fund %", "Active %"]
        for c_idx, uc in enumerate(uw_cols):
            self.format_cell(t_uw.cell(0, c_idx), uc, font_size=Pt(10.5), bold=True, color=self.text_light, align=PP_ALIGN.CENTER if c_idx >= 2 else PP_ALIGN.LEFT, bg_color=self.banner_blue)

        df_uw = top_underweight_stocks if top_underweight_stocks is not None else pd.DataFrame()
        for i in range(1, limit_uw + 1):
            bg = self.zebra_grey if i % 2 == 1 else self.text_light
            if i - 1 < len(df_uw):
                r = df_uw.iloc[i - 1]
                sname = str(r.get("Stock_Name", r.get("securityname", "")))
                sec = str(r.get("Sector", r.get("sector", "")))
                fw = float(r.get("Fund_Weight", 0.0))
                aw = float(r.get("Active_Weight", 0.0))

                self.format_cell(t_uw.cell(i, 0), sname, font_size=Pt(9.5), bold=True, align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_uw.cell(i, 1), sec, font_size=Pt(9.5), align=PP_ALIGN.LEFT, bg_color=bg)
                self.format_cell(t_uw.cell(i, 2), f"{fw:.2f}%" if fw > 0 else "-", font_size=Pt(9.5), align=PP_ALIGN.CENTER, bg_color=bg)
                self.format_cell(t_uw.cell(i, 3), f"{aw:.2f}%", font_size=Pt(9.5), bold=True, color=self.negative_red, align=PP_ALIGN.CENTER, bg_color=bg)
            else:
                for c in range(4):
                    self.format_cell(t_uw.cell(i, c), "", bg_color=bg)

        # -------------------------------------------------------------
        # 4. KPI Callout Box (0.25" below Underweight table)
        # -------------------------------------------------------------
        sec_kpi = next((s for s in sections if s.get("id") == "kpi_callout"), {})
        kpi_msg = kpi_callout_text or sec_kpi.get("text", "Active share at 32% with portfolio turnover at 41%")

        callout = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.62), Inches(6.20), Inches(6.45), Inches(0.58))
        callout.fill.solid()
        callout.fill.fore_color.rgb = RGBColor(0xEA, 0xF2, 0xF9)
        callout.line.color.rgb = self.banner_blue
        callout.line.width = Pt(1.0)

        tf_c = callout.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.alignment = PP_ALIGN.CENTER
        run_c = p_c.add_run()
        run_c.text = f"Key Metric Callout:  {kpi_msg}"
        run_c.font.name = self.font_family
        run_c.font.size = Pt(11)
        run_c.font.bold = True
        run_c.font.color.rgb = self.banner_blue

    def save(self, output_path: str):
        """Saves final presentation with file lock resilience."""
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        try:
            self.prs.save(output_path)
            print(f"Presentation saved successfully to {output_path}")
        except PermissionError:
            base, ext = os.path.splitext(output_path)
            alt_path = f"{base}_new{ext}"
            print(f"Warning: '{output_path}' is currently open in PowerPoint or locked. Saving to '{alt_path}' instead.")
            self.prs.save(alt_path)
            print(f"Presentation saved successfully to {alt_path}")
