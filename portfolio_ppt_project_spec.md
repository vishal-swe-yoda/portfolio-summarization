# Project Plan & Technical Specification: Config-Driven Portfolio PPT Generator

## 1. Project Overview & Core Objectives

Build a modular, configuration-driven Python application that generates fixed-format institutional portfolio review presentations (`.pptx`) from Excel workbooks (`.xlsx`).

The system separates quantitative data processing from qualitative analysis and slide rendering:
- **Deterministic Data Engine (`pandas`)**: Reads portfolio holdings, benchmark weights, and summary metrics from Excel; cleans data; and computes sector exposures, active weights, top overweight/underweight positions, and concentration ratios.
- **Qualitative AI Analyst (`Ollama` / Low-Cost LLM API)**: Receives only pre-calculated metrics in a strict JSON payload and returns structured, evidence-based sector commentary and portfolio risk observations in JSON format. The LLM never performs mathematical calculations.
- **Composite Chart Engine (`matplotlib`)**: Generates high-resolution (300 DPI) dual-band horizontal bar charts displaying both absolute allocation (`% NAV`) and relative active positions (`Active Weight %` with green/red directional bars) on a unified horizontal axis.
- **Declarative Slide Renderer (`python-pptx` + `PyYAML`)**: Reads `report_config.yaml` and `metric_rules.yaml` to assemble 16:9 widescreen slides using reusable layout renderers (`three_table_dashboard`, `sector_analysis`, and `holdings_dashboard`). Adding or reordering sections, changing row limits, or toggling optional slides requires only YAML edits unless a brand-new layout archetype is introduced.

## 2. System Architecture & Directory Structure

### Data Flow Pipeline

| Stage | Component | Input | Output | Responsibility |
| :--- | :--- | :--- | :--- | :--- |
| **1. Configuration** | `ConfigLoader` (`PyYAML`) | `report_config.yaml`, `metric_rules.yaml` | Validated Python config dicts | Defines active pages, layouts, section data sources, row limits, theme palette, and metric formatting rules. |
| **2. Data Engine** | `PortfolioDataEngine` (`pandas`) | Portfolio `.xlsx` + `metric_rules.yaml` | Normalized DataFrames & metric tuples | Cleans raw sheets, computes `Active_Weight = Fund_Weight - BM_Weight`, ranks overweight/underweight sectors, and formats summary metrics. |
| **3. AI Commentary** | `LLMCommentaryEngine` (`urllib` / SDK) | Pre-calculated JSON payload + optional research notes | Structured commentary JSON + audit log | Generates concise bulleted observations and risks for top overweight/underweight sectors and overall portfolio positioning. |
| **4. Charting** | `ChartGenerator` (`matplotlib`) | `sector_summary` DataFrame + theme config | `chart_sector_allocation.png` (300 DPI) | Renders the dual-band `% NAV` and green/red `Active Weight (%)` horizontal bar chart with row gridlines. |
| **5. Rendering** | `PowerPointRenderer` (`python-pptx`) | Config + DataFrames + LLM JSON + Chart PNG | Final `.pptx` presentation | Dispatches layout renderers, applies UTI Mutual Fund theme chrome, builds styled tables, embeds charts, and saves the deck. |

### Recommended Repository Layout

```text
portfolio-ppt-generator/
├── config/
│   ├── report_config.yaml          # Page definitions, layouts, sections, theme, LLM config
│   └── metric_rules.yaml           # Formatting and aggregation rules for portfolio metrics
├── data/
│   └── sample_portfolio_data.xlsx  # Input Excel workbook (Summary, Sector, Holdings tabs)
├── assets/
│   └── uti_logo.png                # Brand logo placed in top-right slide header
├── output/
│   ├── intermediate/               # Generated chart PNGs and LLM JSON audit payloads
│   └── uti_portfolio_review.pptx   # Final generated PowerPoint presentation
├── src/
│   ├── __init__.py
│   ├── data_engine.py              # PortfolioDataEngine (pandas reading, cleaning, aggregation)
│   ├── llm_engine.py               # LLMCommentaryEngine (JSON prompt builder, Ollama/OpenAI client)
│   ├── chart_engine.py             # ChartGenerator (matplotlib dual-band horizontal bar chart)
│   └── ppt_renderer.py             # PowerPointRenderer (layout registry & python-pptx builders)
├── main.py                         # CLI entrypoint orchestrating the 5-stage pipeline
└── requirements.txt                # pandas, openpyxl, matplotlib, python-pptx, PyYAML
```

## 3. Technology Stack & Requirements

| Requirement | Library / Tool | Version / Notes |
| :--- | :--- | :--- |
| **Read & Aggregate Data** | `pandas`, `openpyxl` | `pandas>=2.0.0`, `openpyxl>=3.1.0` for `.xlsx` ingestion |
| **Charts & Diagrams** | `matplotlib`, `numpy` | `matplotlib>=3.7.0` using headless `Agg` backend at 300 DPI |
| **Create Slides** | `python-pptx`, `lxml` | `python-pptx>=0.6.21` with custom XML cell border helpers |
| **Parse Configuration** | `PyYAML` | `PyYAML>=6.0` for `report_config.yaml` and `metric_rules.yaml` |
| **AI Commentary** | Local `Ollama` or OpenAI/Gemini API | Standard library `urllib.request` (zero extra dependencies) with structured JSON output mode and deterministic offline fallback |

## 4. Input Excel Data Specification (`sample_portfolio_data.xlsx`)

The Excel workbook contains three structured tabs. `PortfolioDataEngine` must support reading these sheets and dynamically recalculating active weights in `pandas`.

### Tab 1: `Portfolio_Summary`

Contains four distinct summary blocks used by Page 1 (`three_table_dashboard`):

1. **Quants & Ratios (`Rows 5–12, Columns A–C`)**:
   - Columns: `Metric`, `Raw_Value`, `Display_Value`
   - Rows:
     - `Month-end AUM`: `11203` -> `"INR 11,203 Cr"`
     - `Total Stock Count`: `58` -> `"58"`
     - `Active Share`: `0.32` -> `"32%"`
     - `Top-10 Holdings`: `0.49` -> `"49%"`
     - `Sharpe Ratio`: `0.18` -> `"0.18"`
     - `Beta`: `0.89` -> `"0.89"`
     - `Standard Deviation`: `0.1172` -> `"11.72%"`
     - `Portfolio Turnover Ratio`: `0.41` -> `"41%"`
2. **Market Capitalization (`Rows 17–19, Columns A–C`)**:
   - Columns: `Segment`, `Fund`, `BM`
   - Rows: `Large Cap` (`89%` vs `91%`), `Mid Cap` (`9%` vs `7%`), `Small Cap` (`2%` vs `2%`).
3. **Valuation Parameters (`Rows 24–26, Columns A–C`)**:
   - Columns: `Parameters`, `Fund`, `BM`
   - Rows: `Price to Book` (`3.38` vs `2.96`), `Price to Earnings` (`23.63` vs `21.16`), `Return on Equity` (`18.81` vs `20.11`).
4. **Portfolio Composition — OCF & RoCE Tiers (`Rows 6–8, Columns E–J`)**:
   - Columns: `OCF_Tier`, `OCF_Fund`, `OCF_BM`, `RoCE_Tier`, `RoCE_Fund`, `RoCE_BM`
   - Rows:
     - `C1`: `89%` Fund, `88%` BM | `R1`: `63%` Fund, `56%` BM
     - `C2`: `2%` Fund, `3%` BM | `R2`: `31%` Fund, `37%` BM
     - `C3`: `9%` Fund, `9%` BM | `R3`: `6%` Fund, `7%` BM

### Tab 2: `Sector_Allocation`

Contains sector-level weights and optional analyst research notes (starting at Row 3 header):

| Column Name | Type | Description |
| :--- | :--- | :--- |
| `Sector` | `str` | Sector name (e.g., `Financial Services`, `IT`, `Consumer Services`, `Power`). |
| `Fund_Weight` | `float` | Portfolio weight as `% NAV` (e.g., `32.5`, `10.9`). |
| `BM_Weight` | `float` | Benchmark weight in `%` (e.g., `34.2`, `8.6`). |
| `Active_Weight` | `float` | Formula `=B{row}-C{row}` in Excel, recalculated in `pandas` as `(Fund_Weight - BM_Weight).round(1)`. |
| `Research_Commentary` | `str` | Optional pipe-delimited (`\|`) analyst notes used to ground LLM commentary or serve as offline fallback bullets. |

### Tab 3: `Holdings_Data`

Contains stock-level holdings for Page 3 (`holdings_dashboard`) and optional bottom-up aggregation:

| Column Name | Type | Description |
| :--- | :--- | :--- |
| `Stock_Name` | `str` | Company name (e.g., `HDFC Bank Ltd.`, `ICICI Bank Ltd.`, `Infosys Ltd.`). |
| `Sector` | `str` | Mapped GICS/AMFI sector. |
| `Market_Cap_Tier` | `str` | `Large Cap`, `Mid Cap`, or `Small Cap`. |
| `Fund_Weight` | `float` | Portfolio weight (`% NAV`). |
| `BM_Weight` | `float` | Benchmark weight (`%`). |
| `Active_Weight` | `float` | Active weight (`Fund_Weight - BM_Weight`). |
| `Position_Action` | `str` | Recent positioning note (e.g., `Core Holding`, `Added +45 bps QoQ`, `High Conviction Overweight`). |

## 5. Configuration Specification (`YAML`)

### `config/report_config.yaml`

```yaml
report:
  title: "Portfolio Review"
  report_date: "2026-03-30"
  display_date_page1: "30/Mar/2026"
  display_date_page2: "March 30, 2026"
  benchmark: "BSE 100 TRI"
  website: "www.utimf.com"
  page_size: "16:9"
  theme: "uti_blue"
  logo_path: "assets/uti_logo.png"

llm:
  provider: "ollama"  # Options: "ollama", "openai", "fallback"
  model: "llama3.1:8b"
  endpoint: "http://localhost:11434/api/generate"
  temperature: 0.2
  instruction: "Write concise, evidence-based portfolio observations."

themes:
  uti_blue:
    banner_blue: "#1B4F9C"
    header_blue: "#004C97"
    chart_navy: "#002060"
    positive_green: "#00B050"
    negative_red: "#FF0000"
    subheader_grey: "#D9D9D9"
    zebra_grey: "#E7E6E6"
    accent_cyan: "#0072BC"
    accent_orange: "#F58220"
    text_dark: "#000000"
    text_light: "#FFFFFF"
    font_family: "Carlito"

pages:
  - id: portfolio_summary
    enabled: true
    title: "Portfolio - Statistics"
    layout: "three_table_dashboard"
    sections:
      - id: quants_ratios
        title: "Quants & Ratios"
        type: table
        source: calculated_metrics
        rules: "metric_rules.yaml"
      - id: market_cap
        title: "Market Capitalization"
        type: table
        source: market_cap_summary
      - id: valuation
        title: "Parameters"
        type: table
        source: valuation_metrics
      - id: portfolio_composition
        title: "Portfolio Composition"
        type: split_tier_table
        source: composition_tiers

  - id: sectoral_allocation
    enabled: true
    title: "Portfolio - Sectoral Allocation"
    layout: "sector_analysis"
    sections:
      - id: sector_chart
        title: "Sectoral Breakdown & Active Position (%)"
        type: horizontal_bar_chart
        source: sector_summary
        show_active_weight: true
      - id: overweight
        title: "Overweight (Top 3)"
        type: table_with_commentary
        source: top_overweight_sectors
        limit: 3
        commentary: "llm_analysis"
      - id: underweight
        title: "Underweight (Top 3)"
        type: table_with_commentary
        source: top_underweight_sectors
        limit: 3
        commentary: "llm_analysis"

  - id: holdings_insights
    enabled: true
    title: "Portfolio - Holdings & Investment Insights"
    layout: "holdings_dashboard"
    sections:
      - id: top_holdings
        title: "Top 10 Holdings & Benchmark Comparison"
        type: table
        source: top_holdings_summary
        limit: 10
      - id: portfolio_changes
        title: "Key Active Stock Positions & Recent Activity"
        type: table
        source: active_stock_positions
        limit: 5
      - id: risk_observations
        title: "AI-Generated Portfolio Observations & Key Risks"
        type: commentary_box
        source: portfolio_insights
        commentary: "llm_analysis"
```

### `config/metric_rules.yaml`

```yaml
metrics:
  - id: month_end_aum
    label: "Month-end AUM"
    source_column: "Month-end AUM"
    format: "INR {:,} Cr"
    aggregation: "direct"
  - id: total_stock_count
    label: "Total Stock Count"
    source_column: "Total Stock Count"
    format: "{:d}"
    aggregation: "count_holdings"
  - id: active_share
    label: "Active Share"
    source_column: "Active Share"
    format: "{:.0f}%"
    aggregation: "half_sum_abs_active_weight"
  - id: top_10_holdings
    label: "Top-10 Holdings"
    source_column: "Top-10 Holdings"
    format: "{:.0f}%"
    aggregation: "sum_top_10_weights"
  - id: sharpe_ratio
    label: "Sharpe Ratio"
    source_column: "Sharpe Ratio"
    format: "{:.2f}"
    aggregation: "direct"
  - id: beta
    label: "Beta"
    source_column: "Beta"
    format: "{:.2f}"
    aggregation: "direct"
  - id: standard_deviation
    label: "Standard Deviation"
    source_column: "Standard Deviation"
    format: "{:.2f}%"
    aggregation: "direct"
  - id: portfolio_turnover_ratio
    label: "Portfolio Turnover Ratio"
    source_column: "Portfolio Turnover Ratio"
    format: "{:.0f}%"
    aggregation: "direct"
```

## 6. Detailed Component Specifications

### Component 1: `PortfolioDataEngine` (`src/data_engine.py`)

1. **Sector Processing**:
   - Filter out summary/total rows (`~df["Sector"].str.contains("Total", case=False)`).
   - Coerce `Fund_Weight` and `BM_Weight` to float, filling `NaN` with `0.0`.
   - Compute `Active_Weight = (df["Fund_Weight"] - df["BM_Weight"]).round(1)`.
   - Sort `sector_summary` by `Fund_Weight` descending (from `Financial Services` `32.5%` down to `Realty` `0.0%`).
2. **Overweight / Underweight Extraction**:
   - `top_overweight_sectors`: Filter `Active_Weight > 0`, sort by `Active_Weight` descending, take top `limit` (default `3`: `Consumer Services` `+3.6%`, `Consumer Durables` `+3.1%`, `IT` `+2.3%`).
   - `top_underweight_sectors`: Filter `Active_Weight < 0`, sort by `Active_Weight` ascending, take top `limit` (default `3`: `Power` `-3.2%`, `FMCG` `-1.9%`, `Financial Services` `-1.7%`).
3. **Holdings Processing**:
   - Filter out subtotal rows, compute `Active_Weight`, sort `top_holdings_summary` by `Fund_Weight` descending, and sort `active_stock_positions` by `Active_Weight.abs()` descending.

### Component 2: `LLMCommentaryEngine` (`src/llm_engine.py`)

1. **Strict Input Payload Contract**:
   Construct and serialize the following JSON object from `PortfolioDataEngine` results (never ask the LLM to compute weights):
   ```json
   {
     "report_date": "2026-03-30",
     "benchmark": "BSE 100 TRI",
     "top_overweight": [
       {
         "sector": "Consumer Services",
         "fund_weight": 6.7,
         "benchmark_weight": 3.1,
         "active_weight": 3.6
       }
     ],
     "top_underweight": [
       {
         "sector": "Power",
         "fund_weight": 0.2,
         "benchmark_weight": 3.4,
         "active_weight": -3.2
       }
     ],
     "instruction": "Write concise, evidence-based portfolio observations."
   }
   ```
2. **Strict Output JSON Contract**:
   ```json
   {
     "overweight_commentary": [
       {
         "sector": "Consumer Services",
         "fund_weight": 6.7,
         "benchmark_weight": 3.1,
         "active_weight": 3.6,
         "observations": [
           "Potential for a long growth runway driven by: Changing lifestyles",
           "Improving working-age population",
           "Unorganized to organized shift"
         ],
         "risks": [
           "Sector-specific earnings or valuation changes may affect relative performance."
         ]
       }
     ],
     "underweight_commentary": [
       {
         "sector": "Power",
         "fund_weight": 0.2,
         "benchmark_weight": 3.4,
         "active_weight": -3.2,
         "observations": [
           "Highly capital intensive & long gestation",
           "Moderate RoCEs being utility sector",
           "Near-term growth lower while valuations trading fair"
         ],
         "risks": [
           "Sharp cyclical rally in underweight constituents could create relative drag."
         ]
       }
     ],
     "portfolio_observations": [
       "High quality bias: 89% of portfolio companies reside in top Operating Cash Flow tier (C1) and 63% in top RoCE tier (R1 vs 56% for BSE 100 TRI)."
     ],
     "data_limitations": [
       "Sector commentary reflects point-in-time holdings as of March 30, 2026."
     ]
   }
   ```
3. **Resilience & Auditability**:
   - Save both `llm_input_payload` and `llm_output` to `output/intermediate/llm_payload_and_response.json`.
   - If the local Ollama server or external API is unreachable, automatically fall back to parsing the `Research_Commentary` column from `Sector_Allocation` so slide generation never fails offline.

### Component 3: `ChartGenerator` (`src/chart_engine.py`)

To match the reference Slide 2 chart (`Sectoral Breakdown & Active Position (%)`) on a single clean canvas:
- **Canvas Size & Resolution**: `figsize=(6.20, 5.38)` inches at `dpi=300`.
- **Unified X-Axis (`0` to `70`)**:
  - **Left Visual Element (`% NAV`)**: Plot horizontal bars (`ax.barh`) starting at `left=0.0` with height `0.68` in deep navy (`#002060`). Place data labels (`f"{fw:.1f}"`) at `x = fw + 0.8` for all sectors with `fw > 0.05`.
  - **Right Visual Element (`Active Weight %`)**: Use a vertical zero-baseline at `aw_base = 61.0`. Plot horizontal bars (`ax.barh(idx, aw, left=61.0, height=0.68)`) in green (`#00B050`) when `aw >= 0` and red (`#FF0000`) when `aw < 0`.
  - **Directional Data Labels**: For `aw >= 0`, place left-aligned label at `aw_base + aw + 0.8`. For `aw < 0`, place right-aligned label at `aw_base + aw - 0.8`.
- **Gridlines & Headers**:
  - Draw horizontal separator lines (`#A6A6A6`, `linewidth=0.65`) between every sector row (`y = -0.5, 0.5, 1.5, ..., n - 0.5`).
  - Draw vertical reference lines from `y = -0.5` to `y = n - 0.5` at `x = 0.0` (`#888888`) and `x = 61.0` (`#CCCCCC`) so they stop cleanly below the top column headers.
  - Place bold column headers inside the top plot area at `y = n - 0.15`: `"% NAV"` centered at `x = 14.5`, and `"Active Weight (%)"` centered at `x = 61.0`.

### Component 4: `PowerPointRenderer` (`src/ppt_renderer.py`)

#### Slide Dimensions & Global Chrome
- **Slide Dimensions**: Widescreen 16:9 (`13.333"` width x `7.50"` height).
- **Typography Floor**: Every text run must explicitly set `run.font.name` and `run.font.size >= Pt(10)` (never leave `font.size` as `None`).
- **Shared Chrome on Every Slide**:
  - **Top Banner**: Blue rectangle (`#1B4F9C`) at `(x=0.25", y=0.18", w=10.65", h=0.48")` with a parallelogram cap (`MSO_SHAPE.PARALLELOGRAM`) at `(x=10.75", y=0.18", w=0.32", h=0.48")` to create the slanted right edge. Title text in `Pt(21)` bold white.
  - **Top-Right Logo**: `uti_logo.png` placed at `(x=11.75", y=0.12", width=1.34")`.
  - **Slide Number**: Bottom-left textbox at `(x=0.18", y=6.95", w=0.35", h=0.32")` in `Pt(11)`.
  - **Footer Textbox**: Placed at `x=0.58", w=12.50"` (`y=6.55", h=0.75"` on Page 1 for 2-line methodology notes; `y=6.95", h=0.35"` on Pages 2–3 for 1-line source notes) in `Pt(10)`.
  - **Bottom Tri-Color Accent Strip**: Three flat rectangles at `y=7.36", h=0.14"`:
    - Segment 1: `x=0.00", w=6.20"` (`#1B4F9C` navy)
    - Segment 2: `x=6.20", w=3.05"` (`#0072BC` cyan-blue)
    - Segment 3: `x=9.25", w=4.083"` (`#F58220` orange)

#### Layout 1: `three_table_dashboard` (Page 1 — Portfolio Summary)

| Table Section | Position `(x, y)` | Size `(w, h)` | Grid (`Rows x Cols`) | Column Widths & Styling Details |
| :--- | :--- | :--- | :--- | :--- |
| **Quants & Ratios** | `(0.35", 1.02")` | `(4.15", 5.00")` | `9 x 2` | Cols: `[2.45", 1.70"]`. Row 0 merged header (`#004C97`, white bold `14pt`). Rows 1–8: metric label bold left-aligned (`12.5pt`), value centered (`12.5pt`). |
| **Market Capitalization** | `(4.80", 1.02")` | `(3.45", 2.45")` | `5 x 3` | Cols: `[1.35", 1.05", 1.05"]`. Row 0 merged header (`#004C97`, `14pt`). Row 1 subheader `["", "Fund", "BM"]` (`#D9D9D9` grey, black bold `12.5pt`). Rows 2–4: `Large Cap`, `Mid Cap`, `Small Cap`. |
| **Valuation Parameters** | `(4.80", 3.68")` | `(3.45", 2.34")` | `4 x 3` | Cols: `[1.65", 0.90", 0.90"]`. Row 0 header `["Parameters", "Fund", "BM"]` (`#004C97`, white bold `14pt`). Rows 1–3: `Price to Book`, `Price to Earnings`, `Return on Equity`. |
| **Portfolio Composition** | `(8.55", 1.02")` | `(4.43", 2.50")` | `6 x 6` | Cols: `[0.69", 0.77", 0.75", 0.69", 0.77", 0.76"]`. Row 0 merged `0..5` (`Portfolio Composition`, `#004C97`). Row 1 merged `0..2` (`OCF`, `#D9D9D9`) and `3..5` (`RoCE`, `#D9D9D9`). Row 2 subheaders `["Tier", "Fund", "BM", "Tier", "Fund", "BM"]`. Rows 3–5: `C1..C3` and `R1..R3`. |

#### Layout 2: `sector_analysis` (Page 2 — Portfolio Sectoral Allocation)

| Element | Position `(x, y)` | Size `(w, h)` | Specification |
| :--- | :--- | :--- | :--- |
| **Left Section Banner** | `(0.25", 0.80")` | `(6.20", 0.42")` | `#1B4F9C` fill, white bold `14pt`: `"Sectoral Breakdown & Active Position (%)"`. |
| **Left Sector Chart** | `(0.25", 1.28")` | `(6.20", 5.38")` | Embedded 300 DPI PNG (`chart_sector_allocation.png`) with `0.75pt` grey border (`#555555`). |
| **Upper-Right Banner** | `(6.62", 0.80")` | `(6.45", 0.42")` | `#1B4F9C` fill, white bold `14pt`: `"Overweight (Top 3)"`. |
| **Overweight Table** | `(6.62", 1.28")` | `(6.45", 1.20")` | `4 x 2` table, Cols `[1.75", 4.70"]`. Row 0 header (`#1B4F9C`, white bold `11pt`): `"Sector\n(%NAV/ %AW)"` \| `"Sector Overview"`. Rows 1–3: alternating `#E7E6E6` and `#FFFFFF` fills; Col 0 bold `10.5pt`; Col 1 bulleted observations `10pt`. |
| **Lower-Right Banner** | `(6.62", 3.96")` | `(6.45", 0.42")` | `#1B4F9C` fill, white bold `14pt`: `"Underweight (Top 3)"`. |
| **Underweight Table** | `(6.62", 4.44")` | `(6.45", 1.20")` | Same `4 x 2` structure and styling as Overweight table, populated with the top 3 underweight sectors (`Power`, `FMCG`, `Financial Services`). |

#### Layout 3: `holdings_dashboard` (Optional Page 3 — Holdings & Investment Insights)

| Element | Position `(x, y)` | Size `(w, h)` | Specification |
| :--- | :--- | :--- | :--- |
| **Left Section (`Top 10 Holdings`)** | Banner `(0.25", 0.80")` Table `(0.25", 1.28")` | Banner `(6.20", 0.42")` Table `(6.20", 3.60")` | `12 x 5` table (Header + 10 stocks + Top-10 Total row). Cols: `[2.30", 1.45", 0.80", 0.80", 0.85"]`. Color-codes `Active %` in green (`#00B050`) or red (`#FF0000`). |
| **Upper-Right (`Active Positions`)** | Banner `(6.62", 0.80")` Table `(6.62", 1.28")` | Banner `(6.45", 0.42")` Table `(6.45", 1.60")` | `6 x 4` table highlighting the 5 largest active stock positions (`Stock Name`, `Sector`, `Active %`, `Portfolio Positioning`). |
| **Lower-Right (`AI Observations`)** | Banner `(6.62", 3.82")` Table `(6.62", 4.30")` | Banner `(6.45", 0.42")` Table `(6.45", 1.40")` | `4 x 1` table displaying bulleted portfolio-level observations and valuation/concentration risks from `llm_output["portfolio_observations"]`. |

## 7. Implementation & Verification Plan for Antigravity

1. **Step 1 — Scaffold Project & Configs**:
   - Create directory structure (`config/`, `data/`, `assets/`, `src/`, `output/intermediate/`).
   - Write `config/report_config.yaml` and `config/metric_rules.yaml`.
2. **Step 2 — Generate Sample Input Workbook**:
   - Build `data/sample_portfolio_data.xlsx` with `Portfolio_Summary`, `Sector_Allocation`, and `Holdings_Data` sheets populated with the March 30, 2026 (`BSE 100 TRI`) dataset.
3. **Step 3 — Implement Data & LLM Engines**:
   - Implement `PortfolioDataEngine` in `src/data_engine.py` and verify via unit tests that `Active_Weight` calculations and top-3 overweight/underweight rankings match expected values (`Consumer Services +3.6%`, `Consumer Durables +3.1%`, `IT +2.3%`; `Power -3.2%`, `FMCG -1.9%`, `Financial Services -1.7%`).
   - Implement `LLMCommentaryEngine` in `src/llm_engine.py` with JSON schema validation and offline fallback.
4. **Step 4 — Implement Chart & PPTX Renderers**:
   - Implement `ChartGenerator.render_sector_chart()` in `src/chart_engine.py` and verify that the vertical reference lines stop cleanly below the `"% NAV"` and `"Active Weight (%)"` headers.
   - Implement `PowerPointRenderer` in `src/ppt_renderer.py` with custom XML cell border styling (`a:lnL`, `a:lnR`, `a:lnT`, `a:lnB`) and zero text/table collisions.
5. **Step 5 — End-to-End Execution & Visual QA**:
   - Run `python3 main.py --config config/report_config.yaml --rules config/metric_rules.yaml --excel data/sample_portfolio_data.xlsx --output output/uti_portfolio_review.pptx`.
   - Convert `output/uti_portfolio_review.pptx` to PDF/PNG via headless LibreOffice and verify that all text runs are `>= 10pt`, no table cells overflow or collide with footers, and all three slides match the target visual layout.
