"""
Main CLI entrypoint and programmatic pipeline orchestrator:
1. Configuration loading (report_config.yaml, metric_rules.yaml)
2. PortfolioDataEngine (Excel reading, cleaning, metric calculations)
3. LLMCommentaryEngine (AI observations with offline fallback)
4. ChartGenerator (Matplotlib 300 DPI dual-band sector allocation chart)
5. PowerPointRenderer (python-pptx declarative slide assembly)
"""

import argparse
import os
import sys
from typing import Dict, Any, Optional
import yaml

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

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
        default=None,
        help="Path for output presentation (default: output/uti_portfolio_review.pptx or output_{seq}.pptx)"
    )
    return parser.parse_args()


def load_yaml(file_path: str) -> dict:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Configuration file not found: {file_path}")
    with open(file_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def run_pipeline(
    excel_path: str,
    output_path: Optional[str] = None,
    config_path: str = "config/report_config.yaml",
    rules_path: Optional[str] = "config/metric_rules.yaml",
    sequence_num: Optional[int] = None
) -> Dict[str, Any]:
    """
    Executes the 5-stage portfolio PPT generation pipeline for a given Excel file.
    
    Args:
        excel_path: Path to the input Excel workbook (.xlsx or .xls)
        output_path: Optional custom output presentation path
        config_path: Path to report_config.yaml
        rules_path: Path to metric_rules.yaml
        sequence_num: Optional sequence number to align input and output filenames
        
    Returns:
        Dict containing execution status, input_file, output_file, and output_path.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel workbook not found at: {excel_path}")

    # Determine output presentation path
    if not output_path:
        if sequence_num is not None:
            output_path = os.path.join("output", f"output_{sequence_num}.pptx")
        else:
            output_path = os.path.join("output", "uti_portfolio_review.pptx")

    output_dir = os.path.dirname(output_path) or "output"
    os.makedirs(output_dir, exist_ok=True)

    intermediate_dir = os.path.join(output_dir, "intermediate")
    os.makedirs(intermediate_dir, exist_ok=True)

    # 1. Load Configurations
    config = load_yaml(config_path)
    active_rules_path = rules_path if (rules_path and os.path.exists(rules_path)) else None

    # 2. Ingest and Process Data
    report_cfg = config.get("report", {})
    fund_name = report_cfg.get("fund_name", "UTI Multi Cap Fund")
    bm_name = report_cfg.get("bm_name", "BSE100")

    data_engine = PortfolioDataEngine(
        excel_path=excel_path,
        metric_rules_path=active_rules_path,
        fund_name=fund_name,
        bm_name=bm_name
    )
    data_engine.load_data()
    data_bundle = data_engine.get_data_bundle()

    # Intermediate filename prefix (namespaced by sequence number if present)
    prefix = f"pandas_output_{sequence_num}" if sequence_num is not None else "pandas_output"
    data_engine.save_pandas_output(output_dir=intermediate_dir, filename_base=prefix)

    # 3. Generate Qualitative Commentary
    llm_engine = LLMCommentaryEngine(config=config)
    llm_response = llm_engine.generate_commentary(data_bundle=data_bundle, output_audit_dir=intermediate_dir)

    # 4. Generate Composite Sector Allocation Chart
    theme_cfg = config.get("themes", {}).get(config.get("report", {}).get("theme", "uti_blue"), {})
    chart_gen = ChartGenerator(theme_config=theme_cfg)
    chart_filename = f"chart_sector_{sequence_num}.png" if sequence_num is not None else "chart_sector_allocation.png"
    chart_output = os.path.join(intermediate_dir, chart_filename)
    chart_gen.render_sector_chart(data_bundle["sector_summary"], output_path=chart_output)

    # 5. Assemble PowerPoint Presentation
    renderer = PowerPointRenderer(config=config)
    pages = config.get("pages", [])
    slide_counter = 1

    for page in pages:
        if not page.get("enabled", True):
            continue

        layout = page.get("layout")
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

    saved_path = renderer.save(output_path)
    if not saved_path or not os.path.exists(saved_path):
        raise IOError(f"Failed to generate or save presentation to '{output_path}'.")

    return {
        "status": "success",
        "input_file": os.path.basename(excel_path),
        "output_file": os.path.basename(saved_path),
        "output_path": saved_path,
        "message": "File processed successfully"
    }


def main():
    args = parse_args()
    print("=" * 70)
    print("  CONFIG-DRIVEN INSTITUTIONAL PORTFOLIO PPT GENERATOR")
    print("=" * 70)

    try:
        result = run_pipeline(
            excel_path=args.excel,
            output_path=args.output,
            config_path=args.config,
            rules_path=args.rules
        )
        print(f"\nSuccessfully generated presentation: {result['output_path']}")
    except Exception as e:
        print(f"\nExecution failed: {e}")
        sys.exit(1)

    print("=" * 70)


if __name__ == "__main__":
    main()
