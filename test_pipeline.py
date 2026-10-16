"""
Unit and smoke tests for the Portfolio PPT Generator pipeline components.
"""

import os
import unittest
import pandas as pd
import yaml

from src.data_engine import PortfolioDataEngine
from src.llm_engine import LLMCommentaryEngine
from src.chart_engine import ChartGenerator
from src.ppt_renderer import PowerPointRenderer, hex_to_rgb


class TestPortfolioComponents(unittest.TestCase):
    def setUp(self):
        with open("config/report_config.yaml", "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f)
        with open("config/metric_rules.yaml", "r", encoding="utf-8") as f:
            self.rules = yaml.safe_load(f)

    def test_hex_to_rgb(self):
        rgb = hex_to_rgb("#1B4F9C")
        self.assertEqual((rgb[0], rgb[1], rgb[2]), (0x1B, 0x4F, 0x9C))

    def test_chart_generator_empty(self):
        theme_cfg = self.config.get("themes", {}).get("uti_blue", {})
        gen = ChartGenerator(theme_config=theme_cfg)
        os.makedirs("output/intermediate", exist_ok=True)
        out_path = "output/intermediate/test_chart.png"
        df = pd.DataFrame({
            "Sector": ["Financial Services", "IT", "Consumer Services"],
            "Fund_Weight": [32.5, 10.9, 6.7],
            "BM_Weight": [34.2, 8.6, 3.1],
            "Active_Weight": [-1.7, 2.3, 3.6]
        })
        res = gen.render_sector_chart(df, output_path=out_path)
        self.assertTrue(os.path.exists(res))
        self.assertGreater(os.path.getsize(res), 1000)

    def test_llm_fallback_generation(self):
        engine = LLMCommentaryEngine(self.config)
        ow_df = pd.DataFrame([
            {"Sector": "Consumer Services", "Fund_Weight": 6.7, "BM_Weight": 3.1, "Active_Weight": 3.6, "Research_Commentary": "Changing lifestyles|Organized shift"},
            {"Sector": "Consumer Durables", "Fund_Weight": 4.5, "BM_Weight": 1.4, "Active_Weight": 3.1, "Research_Commentary": "Brand recall|Distribution footprint"},
            {"Sector": "IT", "Fund_Weight": 10.9, "BM_Weight": 8.6, "Active_Weight": 2.3, "Research_Commentary": "Cloud and AI deals|Margin improvement"}
        ])
        uw_df = pd.DataFrame([
            {"Sector": "Power", "Fund_Weight": 0.2, "BM_Weight": 3.4, "Active_Weight": -3.2, "Research_Commentary": "Capital intensive|Moderate RoCE"},
            {"Sector": "FMCG", "Fund_Weight": 6.8, "BM_Weight": 8.7, "Active_Weight": -1.9, "Research_Commentary": "Uneven recovery|Margin pressure"},
            {"Sector": "Financial Services", "Fund_Weight": 32.5, "BM_Weight": 34.2, "Active_Weight": -1.7, "Research_Commentary": "NIM compression|Cautious retail lending"}
        ])
        data_bundle = {
            "top_overweight_sectors": ow_df,
            "top_underweight_sectors": uw_df,
            "composition_tiers": {
                "ocf": [{"tier": "C1", "fund": "89%", "bm": "88%"}],
                "roce": [{"tier": "R1", "fund": "63%", "bm": "56%"}]
            }
        }
        res = engine.generate_commentary(data_bundle, output_audit_dir="output/intermediate")
        self.assertIn("overweight_commentary", res)
        self.assertIn("underweight_commentary", res)
        self.assertIn("portfolio_observations", res)
        self.assertEqual(len(res["overweight_commentary"]), 3)
        self.assertEqual(len(res["underweight_commentary"]), 3)

    def test_ppt_renderer_assembly(self):
        renderer = PowerPointRenderer(self.config)
        ow_df = pd.DataFrame([
            {"Sector": "Consumer Services", "Fund_Weight": 6.7, "BM_Weight": 3.1, "Active_Weight": 3.6, "Research_Commentary": "Changing lifestyles"},
            {"Sector": "Consumer Durables", "Fund_Weight": 4.5, "BM_Weight": 1.4, "Active_Weight": 3.1, "Research_Commentary": "Brand recall"},
            {"Sector": "IT", "Fund_Weight": 10.9, "BM_Weight": 8.6, "Active_Weight": 2.3, "Research_Commentary": "Cloud and AI"}
        ])
        uw_df = pd.DataFrame([
            {"Sector": "Power", "Fund_Weight": 0.2, "BM_Weight": 3.4, "Active_Weight": -3.2, "Research_Commentary": "Capital intensive"},
            {"Sector": "FMCG", "Fund_Weight": 6.8, "BM_Weight": 8.7, "Active_Weight": -1.9, "Research_Commentary": "Uneven recovery"},
            {"Sector": "Financial Services", "Fund_Weight": 32.5, "BM_Weight": 34.2, "Active_Weight": -1.7, "Research_Commentary": "NIM compression"}
        ])
        holdings = pd.DataFrame([
            {"Stock_Name": f"Stock {i}", "Sector": "Financial Services", "Fund_Weight": 5.0, "BM_Weight": 4.0, "Active_Weight": 1.0, "Position_Action": "Core Holding"}
            for i in range(1, 15)
        ])
        data_bundle = {
            "quants_ratios": [
                {"metric": "Month-end AUM", "display_value": "INR 11,203 Cr"},
                {"metric": "Total Stock Count", "display_value": "58"},
                {"metric": "Active Share", "display_value": "32%"},
                {"metric": "Top-10 Holdings", "display_value": "49%"},
                {"metric": "Sharpe Ratio", "display_value": "0.18"},
                {"metric": "Beta", "display_value": "0.89"},
                {"metric": "Standard Deviation", "display_value": "11.72%"},
                {"metric": "Portfolio Turnover Ratio", "display_value": "41%"}
            ],
            "market_cap_summary": [
                {"segment": "Large Cap", "fund": "89%", "bm": "91%"},
                {"segment": "Mid Cap", "fund": "9%", "bm": "7%"},
                {"segment": "Small Cap", "fund": "2%", "bm": "2%"}
            ],
            "valuation_metrics": [
                {"parameter": "Price to Book", "fund": "3.38", "bm": "2.96"},
                {"parameter": "Price to Earnings", "fund": "23.63", "bm": "21.16"},
                {"parameter": "Return on Equity", "fund": "18.81%", "bm": "20.11%"}
            ],
            "composition_tiers": {
                "ocf": [
                    {"tier": "C1", "fund": "89%", "bm": "88%"},
                    {"tier": "C2", "fund": "2%", "bm": "3%"},
                    {"tier": "C3", "fund": "9%", "bm": "9%"}
                ],
                "roce": [
                    {"tier": "R1", "fund": "63%", "bm": "56%"},
                    {"tier": "R2", "fund": "31%", "bm": "37%"},
                    {"tier": "R3", "fund": "6%", "bm": "7%"}
                ]
            },
            "sector_summary": pd.concat([ow_df, uw_df], ignore_index=True),
            "top_overweight_sectors": ow_df,
            "top_underweight_sectors": uw_df,
            "top_holdings_summary": holdings.head(10),
            "active_stock_positions": holdings.head(5)
        }

        llm_engine = LLMCommentaryEngine(self.config)
        llm_res = llm_engine.generate_commentary(data_bundle, output_audit_dir="output/intermediate")

        pages = self.config.get("pages", [])
        slide_idx = 1
        for p in pages:
            layout = p.get("layout")
            if layout == "three_table_dashboard":
                renderer.render_three_table_dashboard(p, data_bundle, slide_idx)
                slide_idx += 1
            elif layout == "sector_analysis":
                renderer.render_sector_analysis(p, data_bundle, llm_res, "output/intermediate/test_chart.png", slide_idx)
                slide_idx += 1
            elif layout == "holdings_dashboard":
                renderer.render_holdings_dashboard(p, data_bundle, llm_res, slide_idx)
                slide_idx += 1

        test_out = "output/intermediate/test_presentation.pptx"
        renderer.save(test_out)
        self.assertTrue(os.path.exists(test_out))
        self.assertGreater(os.path.getsize(test_out), 5000)


if __name__ == "__main__":
    unittest.main()
