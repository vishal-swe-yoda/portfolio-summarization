"""
Unit and smoke tests for the Portfolio PPT Generator pipeline components.
Tests cover:
1. Data Engine: Excel loading, extract_slide_1_tables, 3 stock dataframes, and pandas audit export.
2. LLM Engine: Strategy Pattern (Gemini, DeepSeek, OpenAI, OfflineFallback) & fallback generation.
3. Chart Engine: Dual-band 300 DPI sector allocation chart rendering.
4. PowerPoint Renderer: 3-slide declarative assembly and Slide 3 holdings dashboard.
"""

import json
import os
import unittest
import openpyxl
import pandas as pd
import yaml

from src.data_engine import PortfolioDataEngine, extract_slide_1_tables
from src.llm_engine import (
    LLMCommentaryEngine,
    LLMStrategyFactory,
    GeminiStrategy,
    DeepSeekStrategy,
    OpenAIStrategy,
    OfflineFallbackStrategy,
)
from src.chart_engine import ChartGenerator
from src.ppt_renderer import PowerPointRenderer, hex_to_rgb


class TestPortfolioComponents(unittest.TestCase):
    def setUp(self):
        with open("config/report_config.yaml", "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f)
        with open("config/metric_rules.yaml", "r", encoding="utf-8") as f:
            self.rules = yaml.safe_load(f)
        self.excel_path = "data/sample_portfolio_data.xlsx"

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

    def test_llm_strategy_factory(self):
        data_bundle = {
            "top_overweight_sectors": pd.DataFrame(),
            "top_underweight_sectors": pd.DataFrame(),
            "composition_tiers": {}
        }
        strat_gemini = LLMStrategyFactory.get_strategy("gemini", self.config, data_bundle)
        self.assertIsInstance(strat_gemini, GeminiStrategy)

        strat_deepseek = LLMStrategyFactory.get_strategy("deepseek", self.config, data_bundle)
        self.assertIsInstance(strat_deepseek, DeepSeekStrategy)

        strat_openai = LLMStrategyFactory.get_strategy("openai", self.config, data_bundle)
        self.assertIsInstance(strat_openai, OpenAIStrategy)

        strat_fallback = LLMStrategyFactory.get_strategy("fallback", self.config, data_bundle)
        self.assertIsInstance(strat_fallback, OfflineFallbackStrategy)

    def test_gemini_strategy_api_key_loading(self):
        """Verifies that GeminiStrategy properly loads the API key from .env outside debugger mode."""
        strat = GeminiStrategy(self.config)
        self.assertIsNotNone(strat.api_key)
        self.assertEqual(strat.model, "gemini-3.8-flash")

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

    def test_extract_slide_1_tables(self):
        if os.path.exists(self.excel_path):
            res = extract_slide_1_tables(self.excel_path, "UTI Multi Cap Fund", "BSE100")
            self.assertIn("quants", res)
            self.assertIn("mcap", res)
            self.assertIn("valuation", res)
            self.assertIn("composition", res)
            self.assertFalse(res["quants"].empty)
            self.assertFalse(res["mcap"].empty)
            self.assertFalse(res["valuation"].empty)
            self.assertFalse(res["composition"].empty)

    def test_stock_sheet_three_dataframes(self):
        """Verifies that _process_stock_sheet creates the 3 required stock DataFrames."""
        if os.path.exists(self.excel_path):
            engine = PortfolioDataEngine(excel_path=self.excel_path)
            engine.load_data()

            # 1. Top 15 Holdings
            self.assertEqual(len(engine.top_holdings_summary), 15)
            self.assertTrue({"Stock_Name", "Sector", "Fund_Weight", "BM_Weight", "Active_Weight"}.issubset(engine.top_holdings_summary.columns))
            # Verify sorted descending by Fund_Weight
            fund_weights = engine.top_holdings_summary["Fund_Weight"].tolist()
            self.assertEqual(fund_weights, sorted(fund_weights, reverse=True))

            # 2. Top 5 Overweight Stocks
            self.assertEqual(len(engine.top_overweight_stocks), 5)
            self.assertTrue((engine.top_overweight_stocks["Active_Weight"] > 0).all())
            ow_weights = engine.top_overweight_stocks["Active_Weight"].tolist()
            self.assertEqual(ow_weights, sorted(ow_weights, reverse=True))

            # 3. Top 5 Underweight Stocks
            self.assertEqual(len(engine.top_underweight_stocks), 5)
            self.assertTrue((engine.top_underweight_stocks["Active_Weight"] < 0).all())
            uw_weights = engine.top_underweight_stocks["Active_Weight"].tolist()
            self.assertEqual(uw_weights, sorted(uw_weights, reverse=False))

            # Check get_data_bundle includes the 3 dataframes
            bundle = engine.get_data_bundle()
            self.assertIn("top_holdings_summary", bundle)
            self.assertIn("top_overweight_stocks", bundle)
            self.assertIn("top_underweight_stocks", bundle)

    def test_save_pandas_output_only_three_dataframes(self):
        """Verifies that save_pandas_output exports ONLY the 3 stock DataFrames."""
        if os.path.exists(self.excel_path):
            engine = PortfolioDataEngine(excel_path=self.excel_path)
            engine.load_data()
            out_files = engine.save_pandas_output(output_dir="output/intermediate", filename_base="pandas_output")
            
            # Check file existence
            self.assertTrue(os.path.exists(out_files["excel"]))
            self.assertTrue(os.path.exists(out_files["txt"]))
            self.assertTrue(os.path.exists(out_files["json"]))

            # Verify Excel workbook has exactly 3 tabs
            wb = openpyxl.load_workbook(out_files["excel"])
            expected_sheets = {"Top_Holdings", "Top_Overweight_Stocks", "Top_Underweight_Stocks"}
            self.assertEqual(set(wb.sheetnames), expected_sheets)

            # Verify JSON has exactly the 3 keys
            with open(out_files["json"], "r", encoding="utf-8") as f:
                json_data = json.load(f)
            self.assertEqual(set(json_data.keys()), expected_sheets)
            self.assertEqual(len(json_data["Top_Holdings"]), 15)
            self.assertEqual(len(json_data["Top_Overweight_Stocks"]), 5)
            self.assertEqual(len(json_data["Top_Underweight_Stocks"]), 5)

            # Verify text file contains sections for all 3
            with open(out_files["txt"], "r", encoding="utf-8") as f:
                txt_content = f.read()
            self.assertIn("[TOP_HOLDINGS]", txt_content)
            self.assertIn("[TOP_OVERWEIGHT_STOCKS]", txt_content)
            self.assertIn("[TOP_UNDERWEIGHT_STOCKS]", txt_content)

    def test_ppt_renderer_assembly(self):
        """End-to-end assembly test with Slide 3 rendered using the 3 stock DataFrames."""
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

        # Generate 15 holdings for Slide 3
        top_holdings = pd.DataFrame([
            {
                "Stock_Name": f"Stock {i}",
                "Sector": "Financial Services" if i % 2 == 0 else "Information Technology",
                "Fund_Weight": round(5.0 - (i * 0.2), 2),
                "BM_Weight": round(4.0 - (i * 0.15), 2),
                "Active_Weight": round((5.0 - (i * 0.2)) - (4.0 - (i * 0.15)), 2)
            }
            for i in range(1, 16)
        ])

        top_ow = pd.DataFrame([
            {"Stock_Name": f"OW Stock {i}", "Sector": "Financial Services", "Fund_Weight": 4.0, "BM_Weight": 2.0, "Active_Weight": 2.0}
            for i in range(1, 6)
        ])

        top_uw = pd.DataFrame([
            {"Stock_Name": f"UW Stock {i}", "Sector": "Energy", "Fund_Weight": 1.0, "BM_Weight": 4.0, "Active_Weight": -3.0}
            for i in range(1, 6)
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
            "top_holdings_summary": top_holdings,
            "top_overweight_stocks": top_ow,
            "top_underweight_stocks": top_uw,
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
                kpi_sec = next((s for s in p.get("sections", []) if s.get("id") == "kpi_callout"), {})
                kpi_text = kpi_sec.get("text", "Active share at 32% with portfolio turnover at 41%")
                renderer.render_holdings_dashboard(
                    page_cfg=p,
                    top_holdings_summary=top_holdings,
                    top_overweight_stocks=top_ow,
                    top_underweight_stocks=top_uw,
                    slide_num=slide_idx,
                    kpi_callout_text=kpi_text,
                )
                slide_idx += 1

        test_out = "output/intermediate/test_presentation.pptx"
        saved = renderer.save(test_out)
        self.assertIsNotNone(saved)
        self.assertTrue(os.path.exists(saved))

    def test_save_resilience_and_re_saving(self):
        """Verifies that PermissionError on primary save triggers re-saving to alternative path."""
        from unittest.mock import MagicMock
        renderer = PowerPointRenderer(self.config)
        
        # Mock prs.save so 1st call raises PermissionError, 2nd call succeeds
        mock_save = MagicMock()
        mock_save.side_effect = [PermissionError("File locked"), None]
        renderer.prs.save = mock_save

        out_path = "output/intermediate/test_locked.pptx"
        saved = renderer.save(out_path)

        # Should have attempted primary, then alternative path
        self.assertEqual(mock_save.call_count, 2)
        self.assertEqual(saved, "output/intermediate/test_locked_new.pptx")

    def test_save_graceful_re_saving_failure_handling(self):
        """Verifies that re-saving failure is handled gracefully with error logging and returns None."""
        from unittest.mock import MagicMock
        renderer = PowerPointRenderer(self.config)

        # Mock prs.save so all save attempts raise PermissionError
        mock_save = MagicMock()
        mock_save.side_effect = PermissionError("Permanent lock")
        renderer.prs.save = mock_save

        out_path = "output/intermediate/test_permanently_locked.pptx"
        # Should NOT raise an exception; must handle gracefully
        saved = renderer.save(out_path)
        self.assertIsNone(saved)

    def test_run_pipeline_function(self):
        """Verifies programmatic execution of run_pipeline with sequence numbering."""
        from main import run_pipeline
        if os.path.exists(self.excel_path):
            test_seq = 888
            res = run_pipeline(self.excel_path, sequence_num=test_seq)
            self.assertEqual(res["status"], "success")
            self.assertEqual(res["output_file"], f"output_{test_seq}.pptx")
            self.assertTrue(os.path.exists(res["output_path"]))

    def test_allocate_next_sequence_number(self):
        """Verifies thread-safe sequence number allocation."""
        from app import allocate_next_sequence_number
        seq = allocate_next_sequence_number("data", "output")
        self.assertIsInstance(seq, int)
        self.assertGreaterEqual(seq, 1)

    def test_api_upload_invalid_extension(self):
        """Verifies that non-Excel extensions return HTTP 400."""
        from fastapi.testclient import TestClient
        from app import app

        client = TestClient(app)
        res = client.post(
            "/upload",
            files={"file": ("test.pdf", b"%PDF-dummy-content", "application/pdf")}
        )
        self.assertEqual(res.status_code, 400)
        json_resp = res.json()
        self.assertEqual(json_resp["status"], "error")
        self.assertIn("Invalid file format", json_resp["message"])

    def test_api_upload_empty_file(self):
        """Verifies that empty file uploads return HTTP 400."""
        from fastapi.testclient import TestClient
        from app import app

        client = TestClient(app)
        res = client.post(
            "/upload",
            files={"file": ("empty.xlsx", b"", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        )
        self.assertEqual(res.status_code, 400)
        json_resp = res.json()
        self.assertEqual(json_resp["status"], "error")
        self.assertIn("empty", json_resp["message"].lower())


    def test_swagger_ui_and_openapi(self):
        """Verifies that Swagger UI and OpenAPI specification endpoints are active."""
        from fastapi.testclient import TestClient
        from app import app

        client = TestClient(app)
        res_docs = client.get("/docs")
        self.assertEqual(res_docs.status_code, 200)

        res_openapi = client.get("/openapi.json")
        self.assertEqual(res_openapi.status_code, 200)
        openapi_data = res_openapi.json()
        self.assertIn("paths", openapi_data)
        self.assertIn("/upload", openapi_data["paths"])
        self.assertIn("/presentations", openapi_data["paths"])
        self.assertIn("/download/{filename}", openapi_data["paths"])

    def test_root_redirect_to_docs(self):
        """Verifies that accessing root redirects to the interactive Swagger UI."""
        from fastapi.testclient import TestClient
        from app import app

        client = TestClient(app)
        res = client.get("/", follow_redirects=False)
        self.assertIn(res.status_code, [302, 307])
        self.assertEqual(res.headers["location"], "/docs")

    def test_presentations_endpoint(self):
        """Verifies the presentations list endpoint."""
        from fastapi.testclient import TestClient
        from app import app

        client = TestClient(app)
        res = client.get("/presentations")
        self.assertEqual(res.status_code, 200)
        self.assertIsInstance(res.json(), list)


if __name__ == "__main__":
    unittest.main()
