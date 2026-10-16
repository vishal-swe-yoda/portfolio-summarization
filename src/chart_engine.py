"""
ChartGenerator: Generates high-resolution (300 DPI) dual-band horizontal bar charts
displaying both absolute allocation (% NAV) and relative active positions (Active Weight %)
on a unified horizontal axis.
"""

from typing import Dict, Any, Optional
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


class ChartGenerator:
    def __init__(self, theme_config: Optional[Dict[str, Any]] = None):
        theme = theme_config or {}
        self.chart_navy = theme.get("chart_navy", "#002060")
        self.positive_green = theme.get("positive_green", "#00B050")
        self.negative_red = theme.get("negative_red", "#FF0000")
        self.text_dark = theme.get("text_dark", "#000000")
        self.font_family = theme.get("font_family", "Calibri")

    def render_sector_chart(
        self,
        sector_summary: pd.DataFrame,
        output_path: str = "output/intermediate/chart_sector_allocation.png"
    ) -> str:
        """
        Renders the dual-band horizontal bar chart and saves as 300 DPI PNG.
        
        Canvas: figsize=(6.20, 5.38), dpi=300
        Unified X-Axis: 0 to 70
        Left: % NAV (navy bars from 0.0)
        Right: Active Weight % (baseline at 61.0, green for >=0, red for <0)
        """
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # Ensure we filter out totals and sort descending by Fund_Weight
        df = sector_summary.copy()
        if "Sector" in df.columns:
            df = df[~df["Sector"].astype(str).str.contains("Total", case=False, na=False)]
        
        # Ensure sorted descending so largest is top
        df = df.sort_values(by="Fund_Weight", ascending=False).reset_index(drop=True)
        
        n = len(df)
        if n == 0:
            # Fallback blank image
            fig, ax = plt.subplots(figsize=(6.20, 5.38), dpi=300)
            ax.text(0.5, 0.5, "No Sector Data Available", ha="center", va="center")
            fig.savefig(output_path, dpi=300, bbox_inches="tight")
            plt.close(fig)
            return output_path

        # Layout parameters
        fig, ax = plt.subplots(figsize=(6.20, 5.38), dpi=300)
        fig.patch.set_facecolor("#FFFFFF")
        ax.set_facecolor("#FFFFFF")

        y_positions = np.arange(n)
        sectors = [str(s) for s in df["Sector"]]
        fund_weights = [float(v) for v in df["Fund_Weight"]]
        active_weights = [float(v) for v in df["Active_Weight"]]

        aw_base = 61.0
        bar_height = 0.68

        # 1. Plot % NAV bars
        ax.barh(
            y_positions,
            fund_weights,
            left=0.0,
            height=bar_height,
            color=self.chart_navy,
            edgecolor="none",
            zorder=3
        )

        # 2. Plot Active Weight % bars
        colors_aw = [self.positive_green if aw >= 0 else self.negative_red for aw in active_weights]
        ax.barh(
            y_positions,
            active_weights,
            left=aw_base,
            height=bar_height,
            color=colors_aw,
            edgecolor="none",
            zorder=3
        )

        # 3. Add Data Labels
        for idx in range(n):
            fw = fund_weights[idx]
            aw = active_weights[idx]

            # % NAV label
            if fw > 0.05:
                ax.text(
                    fw + 0.8,
                    idx,
                    f"{fw:.1f}",
                    va="center",
                    ha="left",
                    fontsize=8.0,
                    fontweight="bold",
                    color=self.chart_navy,
                    fontfamily=self.font_family,
                    zorder=4
                )

            # Active Weight label
            if aw >= 0:
                ax.text(
                    aw_base + aw + 0.8,
                    idx,
                    f"{aw:.1f}",
                    va="center",
                    ha="left",
                    fontsize=7.8,
                    fontweight="bold",
                    color=self.positive_green,
                    fontfamily=self.font_family,
                    zorder=4
                )
            else:
                ax.text(
                    aw_base + aw - 0.8,
                    idx,
                    f"{aw:.1f}",
                    va="center",
                    ha="right",
                    fontsize=7.8,
                    fontweight="bold",
                    color=self.negative_red,
                    fontfamily=self.font_family,
                    zorder=4
                )

        # 4. Horizontal Row Gridlines
        for y_sep in np.arange(-0.5, n, 1.0):
            ax.axhline(
                y=y_sep,
                color="#C0C0C0",
                linewidth=0.55,
                linestyle="-",
                xmin=0.0,
                xmax=1.0,
                zorder=1
            )

        # 5. Vertical Reference Lines (stopping at y = -0.5 so headers stay clean)
        # Left baseline at x = 0.0
        ax.plot([0.0, 0.0], [-0.5, n - 0.5], color="#888888", linewidth=0.9, zorder=2)
        # Active weight baseline at x = 61.0
        ax.plot([aw_base, aw_base], [-0.5, n - 0.5], color="#CCCCCC", linewidth=0.9, zorder=2)

        # Invert Y so highest Fund_Weight is at top
        ax.invert_yaxis()

        # 6. Column Headers inside top area
        ax.text(
            14.5,
            -0.85,
            "% NAV",
            ha="center",
            va="bottom",
            fontsize=9.5,
            fontweight="bold",
            color=self.chart_navy,
            fontfamily=self.font_family
        )
        ax.text(
            aw_base,
            -0.85,
            "Active Weight (%)",
            ha="center",
            va="bottom",
            fontsize=9.5,
            fontweight="bold",
            color=self.text_dark,
            fontfamily=self.font_family
        )

        # Axes styling
        ax.set_yticks(y_positions)
        ax.set_yticklabels(sectors, fontsize=8.0, fontweight="bold", fontfamily=self.font_family, color="#1A1A1A")
        ax.set_xlim(-1.0, 71.0)
        ax.set_ylim(n - 0.4, -1.3)

        # Hide spines and x-axis ticks
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.xaxis.set_visible(False)
        ax.tick_params(left=False, length=0)

        plt.tight_layout(pad=0.2)
        plt.savefig(output_path, dpi=300, bbox_inches="tight", facecolor=fig.get_facecolor(), edgecolor="none")
        plt.close(fig)

        return output_path
