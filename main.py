"""
Main CLI entrypoint orchestrating the 5-stage portfolio presentation pipeline:
1. Configuration loading (report_config.yaml, metric_rules.yaml)
2. PortfolioDataEngine (Excel reading, cleaning, metric calculations)
3. LLMCommentaryEngine (AI observations with offline fallback)
4. ChartGenerator (Matplotlib 300 DPI dual-band sector allocation chart)
5. PowerPointRenderer (python-pptx declarative slide assembly)
"""

import argparse
import os
import sys
import yaml

from src.data_engine import PortfolioDataEngine
from src.llm_engine import LLMCommentaryEngine
from src.chart_engine import ChartGenerator
from src.ppt_renderer import PowerPointRenderer


def parse_args():
    parser = argparse.ArgumentParser(description="Institutional Portfolio PPT Generator")
    parser.add_argument(
        "--config",
        default="config/report_config.yaml",
        help="Path to report_config.yaml (default: config/report_config.yaml)"
    )
    parser.add_argument(
        "--rules",
        default="config/metric_rules.yaml",
        help="Path to metric_rules.yaml (default: config/metric_rules.yaml)"
    )
    parser.add_argument(
        "--excel",
        default="data/sample_portfolio_data.xlsx",
        help="Path to portfolio Excel workbook (default: data/sample_portfolio_data.xlsx)"
    )
    parser.add_argument(
        "--output",
        default="output/uti_portfolio_review.pptx",
        help="Path for output presentation (default: output/uti_portfolio_review.pptx)"
    )
    return parser.parse_args()


def load_yaml(file_path: str) -> dict:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Configuration file not found: {file_path}")
    with open(file_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def main():
    args = parse_args()
    print("=" * 70)
    print("  CONFIG-DRIVEN INSTITUTIONAL PORTFOLIO PPT GENERATOR")
    print("=" * 70)

    # 1. Load Configurations
    print(f"\n[Stage 1/5] Loading configuration files...")
    config = load_yaml(args.config)
    rules_path = args.rules if os.path.exists(args.rules) else None
    print(f"  - Config loaded: {args.config}")
    if rules_path:
        print(f"  - Metric rules loaded: {rules_path}")

    # 2. Ingest and Process Data
    print(f"\n[Stage 2/5] Reading and processing Excel workbook: {args.excel}")
    if not os.path.exists(args.excel):
        print(f"Error: Excel file does not exist at '{args.excel}'.")
        print("Please place the Excel file at that path or provide --excel <path>.")
        sys.exit(1)

    report_cfg = config.get("report", {})
    fund_name = report_cfg.get("fund_name", "UTI Multi Cap Fund")
    bm_name = report_cfg.get("bm_name", "BSE100")

    data_engine = PortfolioDataEngine(
        excel_path=args.excel,
        metric_rules_path=rules_path,
        fund_name=fund_name,
        bm_name=bm_name
    )
    data_engine.load_data()
    data_bundle = data_engine.get_data_bundle()
    print(f"  - Sectors processed: {len(data_bundle['sector_summary'])}")
    print(f"  - Stock DataFrames for Slide 3 (from _process_stock_sheet):")
    print(f"      * Top Holdings: {len(data_engine.top_holdings_summary)} rows")
    print(f"      * Top Overweight Stocks: {len(data_engine.top_overweight_stocks)} rows")
    print(f"      * Top Underweight Stocks: {len(data_engine.top_underweight_stocks)} rows")
    print(f"  - Exported 3 stock DataFrames to: output/intermediate/pandas_output.[xlsx|txt|json]")

    # 3. Generate AI Commentary
    print(f"\n[Stage 3/5] Generating qualitative commentary & portfolio observations...")
    intermediate_dir = os.path.join(os.path.dirname(args.output), "intermediate")
    os.makedirs(intermediate_dir, exist_ok=True)
    
    llm_engine = LLMCommentaryEngine(config=config)
    llm_response = llm_engine.generate_commentary(data_bundle=data_bundle, output_audit_dir=intermediate_dir)
    print(f"  - Overweight commentary items: {len(llm_response.get('overweight_commentary', []))}")
    print(f"  - Underweight commentary items: {len(llm_response.get('underweight_commentary', []))}")
    print(f"  - Portfolio observations logged to {intermediate_dir}/llm_payload_and_response.json")

    # 4. Generate Composite Chart
    print(f"\n[Stage 4/5] Rendering dual-band horizontal bar chart (300 DPI)...")
    theme_cfg = config.get("themes", {}).get(config.get("report", {}).get("theme", "uti_blue"), {})
    chart_gen = ChartGenerator(theme_config=theme_cfg)
    chart_output = os.path.join(intermediate_dir, "chart_sector_allocation.png")
    chart_gen.render_sector_chart(data_bundle["sector_summary"], output_path=chart_output)
    print(f"  - Chart saved: {chart_output}")

    # 5. Assemble PowerPoint Presentation
    print(f"\n[Stage 5/5] Assembling 16:9 PowerPoint presentation...")
    renderer = PowerPointRenderer(config=config)

    pages = config.get("pages", [])
    slide_counter = 1
    for page in pages:
        if not page.get("enabled", True):
            continue

        layout = page.get("layout")
        print(f"  - Rendering Slide {slide_counter}: [{layout}] - {page.get('title', '')}")

        if layout == "three_table_dashboard":
            renderer.render_three_table_dashboard(page_cfg=page, data_bundle=data_bundle, slide_num=slide_counter)
            slide_counter += 1
        elif layout == "sector_analysis":
            renderer.render_sector_analysis(
                page_cfg=page,
                data_bundle=data_bundle,
                llm_response=llm_response,
                chart_path=chart_output,
                slide_num=slide_counter
            )
            slide_counter += 1
        elif layout == "holdings_dashboard":
            kpi_sec = next((s for s in page.get("sections", []) if s.get("id") == "kpi_callout"), {})
            kpi_text = kpi_sec.get("text", "Active share at 32% with portfolio turnover at 41%")
            renderer.render_holdings_dashboard(
                page_cfg=page,
                top_holdings_summary=data_engine.top_holdings_summary,
                top_overweight_stocks=data_engine.top_overweight_stocks,
                top_underweight_stocks=data_engine.top_underweight_stocks,
                slide_num=slide_counter,
                kpi_callout_text=kpi_text,
            )
            slide_counter += 1
        else:
            print(f"  - Warning: Unknown layout archetype '{layout}'. Skipping.")

    renderer.save(args.output)
    print(f"\nSuccessfully generated presentation: {args.output}")
    print("=" * 70)


if __name__ == "__main__":
    main()
