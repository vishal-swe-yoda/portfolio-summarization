# Config-Driven Portfolio PPT Generator

A modular, configuration-driven Python application that generates institutional-quality, fixed-format portfolio review presentations (`.pptx`) from Excel workbooks (`.xlsx`).

---

## 1. Key Capabilities & Pipeline Architecture

The system cleanly decouples quantitative data processing, qualitative AI synthesis, charting, and declarative presentation assembly:

```
┌─────────────────────────────────┐
│     Excel Portfolio Data        │
│ (data/ValEqAnalysis_*.xlsx)    │
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│  Stage 1: PortfolioDataEngine   │ ──► Cleans raw data, computes Active Weights
│          (pandas)               │     (Fund - BM), extracts top active positions
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│  Stage 2: LLMCommentaryEngine   │ ──► Strategy Pattern:
│     (Gemini / DeepSeek / AI)    │     • GeminiStrategy (Google Gemini API)
└────────────────┬────────────────┘     • DeepSeekStrategy (DeepSeek API)
                 │                      • OpenAIStrategy (OpenAI API)
                 │                      • OfflineFallbackStrategy (Zero-API fallback)
                 ▼
┌─────────────────────────────────┐
│     Stage 3: ChartGenerator     │ ──► Dual-band 300 DPI horizontal bar chart
│          (matplotlib)           │     (% NAV & directional Active Weight %)
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│  Stage 4: PowerPointRenderer    │ ──► 16:9 Widescreen slide assembly with
│         (python-pptx)           │     custom XML cell borders & brand chrome
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│   Final Institutional PPTX      │
│  (output/uti_portfolio_review)  │
└─────────────────────────────────┘
```

1. **Deterministic Data Engine (`src/data_engine.py`)**:
   - Ingests user-provided portfolio workbooks (Summary, Sector Allocations, and Stock Holdings).
   - Dynamically recalculates `Active_Weight = (Fund_Weight - BM_Weight).round(1)`.
   - Computes concentration metrics, top overweight/underweight sectors, and quants ratios.

2. **LLM Qualitative Analyst with Strategy Pattern (`src/llm_engine.py`)**:
   - Implements the **Strategy Pattern** across multiple LLM providers:
     - **`GeminiStrategy`**: Google Gemini API (`gemini-3.8-flash`) via structured JSON schema mode.
     - **`DeepSeekStrategy`**: DeepSeek API (`deepseek-chat`) via OpenAI-compatible endpoints.
     - **`OpenAIStrategy`**: OpenAI API (`gpt-4o-mini`).
     - **`OfflineFallbackStrategy`**: Deterministic offline engine utilizing analyst notes and composition tiers—ensuring deck generation **never fails** if offline or without API keys.
   - Saves audit payloads and LLM responses to `output/intermediate/llm_payload_and_response.json`.

3. **Composite Chart Engine (`src/chart_engine.py`)**:
   - Renders high-resolution 300 DPI dual-band horizontal bar charts on a unified axis:
     - Left band: Deep navy bars for absolute allocation (`% NAV`).
     - Right band: Directional green (positive) and red (negative) bars for `Active Weight %`.
     - Distinct vertical baselines and horizontal row separator gridlines.

4. **Declarative PowerPoint Renderer (`src/ppt_renderer.py`)**:
   - 16:9 widescreen presentation assembly.
   - Enforces institutional typography standards (run font size floor $\ge 10\,\text{pt}$).
   - Applies low-level DrawingML XML cell borders (`a:lnL`, `a:lnR`, `a:lnT`, `a:lnB`) for crisp table styling.
   - Standardized layout archetypes:
     - `three_table_dashboard`: Portfolio quants, market cap breakdown, valuations, and OCF/RoCE tier matrices.
     - `sector_analysis`: Embedded 300 DPI sector allocation chart paired with top overweight and underweight commentary.
     - `holdings_dashboard`: Top-10 holdings comparison table, key active stock positions, and portfolio observations.

---

## 2. Directory Structure

```text
portfolio-ppt-generator/
├── config/
│   ├── report_config.yaml          # Slide definitions, theme palette, LLM provider config
│   └── metric_rules.yaml           # Metric formatting rules and aggregation mappings
├── data/
│   └── ValEqAnalysis_30.03.26 - Copy.xlsx  # User-provided portfolio input workbook
├── assets/
│   └── uti_logo.png                # High-res brand logo placed in slide header
├── output/
│   ├── intermediate/               # Generated 300 DPI chart PNGs and LLM audit JSON
│   └── uti_portfolio_review.pptx   # Final generated PowerPoint deck
├── src/
│   ├── __init__.py
│   ├── data_engine.py              # PortfolioDataEngine (pandas cleaning & calculation)
│   ├── llm_engine.py               # Strategy Pattern LLM engine (Gemini, DeepSeek, Offline)
│   ├── chart_engine.py             # ChartGenerator (matplotlib 300 DPI dual-band charts)
│   └── ppt_renderer.py             # PowerPointRenderer (python-pptx declarative layouts)
├── main.py                         # CLI entrypoint orchestrating the 5-stage pipeline
├── test_pipeline.py                # Comprehensive component test suite
├── requirements.txt                # Core dependencies (pandas, openpyxl, python-pptx, etc.)
└── README.md                       # Project documentation
```

---

## 3. Installation & Setup

### Prerequisites
- Python 3.9+ (Python 3.10 – 3.13 supported)

### Install Dependencies
```bash
pip install -r requirements.txt
```

Core packages installed:
- `pandas` & `openpyxl` (Data manipulation and Excel parsing)
- `matplotlib` & `numpy` (Dual-band charting at 300 DPI)
- `python-pptx` & `lxml` (Slide assembly and XML table styling)
- `PyYAML` (Configuration management)

---

## 4. Configuration & LLM Strategy Selection

Edit `config/report_config.yaml` to configure metadata, branding, and LLM provider:

```yaml
report:
  title: "Portfolio Review"
  report_date: "2026-03-30"
  display_date_page1: "30/Mar/2026"
  display_date_page2: "March 30, 2026"
  benchmark: "BSE 100 TRI"
  website: "www.utimf.com"
  theme: "uti_blue"
  logo_path: "assets/uti_logo.png"

llm:
  provider: "gemini"  # Options: "gemini", "deepseek", "openai", "fallback"
  model: "gemini-1.5-flash"
  endpoint: "https://generativelanguage.googleapis.com/v1beta/models"
  temperature: 0.2
  instruction: "Write concise, evidence-based portfolio observations."
```

### Setting API Keys

- **Google Gemini**:
  ```bash
  # Windows PowerShell
  $env:GEMINI_API_KEY="your-gemini-api-key"
  
  # Linux / macOS
  export GEMINI_API_KEY="your-gemini-api-key"
  ```

- **DeepSeek**:
  ```bash
  $env:DEEPSEEK_API_KEY="your-deepseek-api-key"
  ```

- **Offline Mode**:
  If no API key is set or the network is unavailable, the pipeline automatically activates `OfflineFallbackStrategy` so presentation generation completes without errors.

---

## 5. Running the Generator

### Basic Execution
Run the CLI orchestrator pointing to your Excel workbook:

```bash
python main.py --excel "data/ValEqAnalysis_30.03.26 - Copy.xlsx" --output "output/uti_portfolio_review.pptx"
```

### CLI Arguments
| Argument | Default | Description |
| :--- | :--- | :--- |
| `--excel` | `data/sample_portfolio_data.xlsx` | Path to the source portfolio Excel workbook. |
| `--config` | `config/report_config.yaml` | Path to the main YAML configuration file. |
| `--rules` | `config/metric_rules.yaml` | Path to the metric rules configuration file. |
| `--output` | `output/uti_portfolio_review.pptx` | Destination path for the `.pptx` presentation. |

---

## 6. Testing & Quality Assurance

Run the test suite to verify all pipeline components (data engine, chart generation, LLM strategy fallback, and PowerPoint assembly):

```bash
python test_pipeline.py
```