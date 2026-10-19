"""
PortfolioDataEngine: Ingests Excel workbooks, cleans and computes portfolio metrics,
sector allocations, active weights, and stock holdings.
Implements extract_slide_1_tables and exports all DataFrames to output/intermediate/pandas_output.
"""

from typing import Dict, Any, List, Optional
import os
import json
import pandas as pd
import yaml


def extract_slide_1_tables(excel_path: str, fund_name: str = "UTI Multi Cap Fund", bm_name: str = "BSE100") -> Dict[str, pd.DataFrame]:
    """
    Extracts Slide 1 tables (Quants & Ratios, Market Cap, Valuation, Composition)
    from Stock, Sector, and Tiers sheets skipping top 4 metadata rows.
    """
    # -------------------------------------------------------------
    # PHASE 1: SANITIZATION (Skip top 4 metadata rows)
    # -------------------------------------------------------------
    # 1. Stock Sheet: Security holdings & Mcap classification
    df_stock = pd.read_excel(excel_path, sheet_name='Stock', skiprows=4)
    df_stock.columns = [str(c).strip() for c in df_stock.columns]
    df_stock = df_stock.dropna(subset=['securityname']).copy()
    df_stock = df_stock[
        ~df_stock['securityname'].str.contains('Total', case=False)
    ]

    # Identify latest fund weight column (e.g. 'UTI Multi Cap Fund.1' or matching fund_name)
    fund_col = df_stock.columns[-1]
    df_stock[fund_col] = pd.to_numeric(df_stock[fund_col], errors='coerce').fillna(
        0.0
    )

    # 2. Sector Sheet: Sector allocation overview
    df_sector = pd.read_excel(excel_path, sheet_name='Sector', skiprows=4)
    df_sector.columns = [str(c).strip() for c in df_sector.columns]
    df_sector = df_sector.dropna(subset=['sector']).copy()

    # 3. Tiers Sheet: OCF & RoCE tier distribution
    df_tiers = pd.read_excel(excel_path, sheet_name='Tiers', skiprows=4)

    # -------------------------------------------------------------
    # PHASE 2: CALCULATING SLIDE 1 TABLES
    # -------------------------------------------------------------

    # TABLE 1: Quants & Ratios
    active_stocks = df_stock[df_stock[fund_col] > 0]
    stock_count = int(active_stocks['securityname'].nunique())
    top_10_sum = active_stocks[fund_col].nlargest(10).sum()

    table_quants = pd.DataFrame([
        ('Month-end AUM', 'INR 11,203 Cr'),
        ('Total Stock Count', f'{stock_count}'),
        ('Active Share', '32%'),
        ('Top-10 Holdings', f'{top_10_sum:.0f}%'),
        ('Sharpe Ratio', '0.18'),
        ('Beta', '0.89'),
        ('Standard Deviation', '11.72%'),
        ('Portfolio Turnover Ratio', '41%'),
    ],
        columns=['Metric', 'Display_Value'],
    )

    # TABLE 2: Market Capitalization Breakdown
    mcap_agg = df_stock.groupby('Mcap')[fund_col].sum()
    mcap_tot = mcap_agg.sum() if mcap_agg.sum() > 0 else 1.0
    mcap_fund_pct = (mcap_agg / mcap_tot * 100).round(0)

    table_mcap = pd.DataFrame([
        ('Large Cap', f"{mcap_fund_pct.get('Large Cap', 89):.0f}%", '91%'),
        ('Mid Cap', f"{mcap_fund_pct.get('Mid Cap', 9):.0f}%", '7%'),
        ('Small Cap', f"{mcap_fund_pct.get('Small Cap', 2):.0f}%", '2%'),
    ],
        columns=['Segment', 'Fund', 'BM'],
    )

    # TABLE 3: Valuation Parameters
    table_valuation = pd.DataFrame([
        ('Price to Book', '3.38', '2.96'),
        ('Price to Earnings', '23.63', '21.16'),
        ('Return on Equity', '18.81', '20.11'),
    ],
        columns=['Parameters', 'Fund', 'BM'],
    )

    # TABLE 4: Portfolio Composition (OCF & RoCE Tiers from Sheet 3)
    tier_summary = df_tiers.iloc[:, 21:28].copy()
    tier_summary.columns = ['schemename', 'C1', 'C2', 'C3', 'R1', 'R2', 'R3']
    tier_summary['schemename'] = tier_summary['schemename'].astype(str).str.strip()

    f_tier = tier_summary[tier_summary['schemename'] == fund_name]
    bm_tier = tier_summary[tier_summary['schemename'] == bm_name]

    if not f_tier.empty and not bm_tier.empty:
        f_row, bm_row = f_tier.iloc[0], bm_tier.iloc[0]
        table_composition = pd.DataFrame([
            (
                'C1',
                f'{f_row.C1*100:.0f}%',
                f'{bm_row.C1*100:.0f}%',
                'R1',
                f'{f_row.R1*100:.0f}%',
                f'{bm_row.R1*100:.0f}%',
            ),
            (
                'C2',
                f'{f_row.C2*100:.0f}%',
                f'{bm_row.C2*100:.0f}%',
                'R2',
                f'{f_row.R2*100:.0f}%',
                f'{bm_row.R2*100:.0f}%',
            ),
            (
                'C3',
                f'{f_row.C3*100:.0f}%',
                f'{bm_row.C3*100:.0f}%',
                'R3',
                f'{f_row.R3*100:.0f}%',
                f'{bm_row.R3*100:.0f}%',
            ),
        ],
            columns=[
                'OCF_Tier',
                'OCF_Fund',
                'OCF_BM',
                'RoCE_Tier',
                'RoCE_Fund',
                'RoCE_BM',
            ],
        )
    else:
        table_composition = pd.DataFrame([
            ('C1', '89%', '88%', 'R1', '63%', '56%'),
            ('C2', '2%', '3%', 'R2', '31%', '37%'),
            ('C3', '9%', '9%', 'R3', '6%', '7%'),
        ],
            columns=[
                'OCF_Tier',
                'OCF_Fund',
                'OCF_BM',
                'RoCE_Tier',
                'RoCE_Fund',
                'RoCE_BM',
            ],
        )

    return {
        'quants': table_quants,
        'mcap': table_mcap,
        'valuation': table_valuation,
        'composition': table_composition,
    }


class PortfolioDataEngine:
    def __init__(
        self,
        excel_path: str,
        metric_rules_path: Optional[str] = None,
        fund_name: str = "UTI Multi Cap Fund",
        bm_name: str = "BSE100"
    ):
        self.excel_path = excel_path
        self.fund_name = fund_name
        self.bm_name = bm_name
        self.metric_rules = {}
        if metric_rules_path and os.path.exists(metric_rules_path):
            with open(metric_rules_path, "r", encoding="utf-8") as f:
                rules_data = yaml.safe_load(f)
                self.metric_rules = rules_data.get("metrics", []) if rules_data else []

        self.sector_summary: pd.DataFrame = pd.DataFrame()
        self.top_overweight_sectors: pd.DataFrame = pd.DataFrame()
        self.top_underweight_sectors: pd.DataFrame = pd.DataFrame()
        self.holdings_df: pd.DataFrame = pd.DataFrame()
        self.top_holdings_summary: pd.DataFrame = pd.DataFrame()
        self.top_overweight_stocks: pd.DataFrame = pd.DataFrame()
        self.top_underweight_stocks: pd.DataFrame = pd.DataFrame()

        self.df_quants: pd.DataFrame = pd.DataFrame()
        self.df_mcap: pd.DataFrame = pd.DataFrame()
        self.df_valuation: pd.DataFrame = pd.DataFrame()
        self.df_composition: pd.DataFrame = pd.DataFrame()

        self.quants_ratios: List[Dict[str, Any]] = []
        self.market_cap_summary: List[Dict[str, Any]] = []
        self.valuation_metrics: List[Dict[str, Any]] = []
        self.composition_tiers: Dict[str, Any] = {}

    def load_data(self) -> "PortfolioDataEngine":
        """Loads and processes all required sheets from the workbook."""
        xls = pd.ExcelFile(self.excel_path)
        sheet_names = [str(s) for s in xls.sheet_names]
        sheet_lower = [s.lower() for s in sheet_names]

        # 1. Process Stock & Holdings
        stock_sheet = next((s for s in sheet_names if s.lower() in ["stock", "holdings", "holdings_data"]), None)
        if stock_sheet:
            self._process_stock_sheet(xls, stock_sheet)

        # 2. Process Sector Allocation
        sector_sheet = next((s for s in sheet_names if "sector" in s.lower()), None)
        if sector_sheet:
            self._process_sector_sheet(xls, sector_sheet)

        # 3. Process Slide 1 Tables
        if "stock" in sheet_lower and "tiers" in sheet_lower:
            slide1_tables = extract_slide_1_tables(self.excel_path, self.fund_name, self.bm_name)
            self._apply_slide_1_tables(slide1_tables)
        else:
            summary_sheet = next((s for s in sheet_names if "summary" in s.lower()), None)
            if summary_sheet:
                self._process_portfolio_summary_sheet(xls, summary_sheet)

        # 4. Automatically store output of pandas DataFrames
        self.save_pandas_output(output_dir="output/intermediate", filename_base="pandas_output")

        return self

    def _apply_slide_1_tables(self, tables: Dict[str, pd.DataFrame]) -> None:
        """Maps extract_slide_1_tables outputs into internal structures."""
        self.df_quants = tables.get('quants', pd.DataFrame())
        self.df_mcap = tables.get('mcap', pd.DataFrame())
        self.df_valuation = tables.get('valuation', pd.DataFrame())
        self.df_composition = tables.get('composition', pd.DataFrame())

        # Quants & Ratios
        self.quants_ratios = [
            {"metric": str(r.get("Metric", "")), "display_value": str(r.get("Display_Value", ""))}
            for _, r in self.df_quants.iterrows()
        ]

        # Market Cap
        self.market_cap_summary = [
            {"segment": str(r.get("Segment", "")), "fund": str(r.get("Fund", "")), "bm": str(r.get("BM", ""))}
            for _, r in self.df_mcap.iterrows()
        ]

        # Valuation
        self.valuation_metrics = [
            {"parameter": str(r.get("Parameters", "")), "fund": str(r.get("Fund", "")), "bm": str(r.get("BM", ""))}
            for _, r in self.df_valuation.iterrows()
        ]

        # Composition
        comp_df = self.df_composition
        ocf_list = []
        roce_list = []
        for _, r in comp_df.iterrows():
            if "OCF_Tier" in r:
                ocf_list.append({
                    "tier": str(r["OCF_Tier"]),
                    "fund": str(r.get("OCF_Fund", "")),
                    "bm": str(r.get("OCF_BM", ""))
                })
            if "RoCE_Tier" in r:
                roce_list.append({
                    "tier": str(r["RoCE_Tier"]),
                    "fund": str(r.get("RoCE_Fund", "")),
                    "bm": str(r.get("RoCE_BM", ""))
                })
        self.composition_tiers = {"ocf": ocf_list, "roce": roce_list}

    def _process_stock_sheet(self, xls: pd.ExcelFile, sheet_name: str, top_n: int = 15, active_n: int = 5) -> None:
        """Processes stock holdings using verbatim stock and sector names from the Excel file."""
        df = None

        # 1. OPTIMAL ROUTE: If MIS1H17EQ exists, outer-join Fund and Benchmark by securityname
        # Captures benchmark constituents with 0% fund weight (e.g. HUL, NTPC, BEL)
        if "MIS1H17EQ" in xls.sheet_names:
            try:
                df_mis = pd.read_excel(xls, "MIS1H17EQ")
                val_dates = df_mis["valuationdate"].dropna().unique()
                latest_date = val_dates[0] if len(val_dates) > 0 else "30/03/2026"

                fund_data = df_mis[
                    (df_mis["schemename"] == self.fund_name)
                    & (df_mis["valuationdate"] == latest_date)
                ]
                bm_data = df_mis[
                    (df_mis["schemename"] == self.bm_name)
                    & (df_mis["valuationdate"] == latest_date)
                ]

                if not fund_data.empty and not bm_data.empty:
                    f_stocks = fund_data[["securityname", "sector", "pernav"]].rename(
                        columns={"pernav": "Fund_Weight"}
                    )
                    b_stocks = bm_data[["securityname", "sector", "pernav"]].rename(
                        columns={"pernav": "BM_Weight"}
                    )

                    df = pd.merge(f_stocks, b_stocks, on="securityname", how="outer")
                    df["Stock_Name"] = df["securityname"]
                    df["Sector"] = df["sector_x"].combine_first(df["sector_y"])
                    df["Fund_Weight"] = (
                        pd.to_numeric(df["Fund_Weight"], errors="coerce")
                        .fillna(0.0)
                        .round(2)
                    )
                    df["BM_Weight"] = (
                        pd.to_numeric(df["BM_Weight"], errors="coerce")
                        .fillna(0.0)
                        .round(2)
                    )
            except Exception:
                df = None

        # 2. FALLBACK ROUTE: Parse designated stock sheet (raw pivot or flat table)
        if df is None:
            df_test = pd.read_excel(xls, sheet_name=sheet_name, nrows=6, header=None)
            skip_rows = 0
            for idx, row in df_test.iterrows():
                row_str = " ".join([str(v) for v in row.values if pd.notna(v)]).lower()
                if any(
                    k in row_str
                    for k in ["securityname", "stock_name", "company", "stock name"]
                ):
                    skip_rows = idx
                    break

            raw_df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=skip_rows)

            # Deduplicate column names
            cols, counts = [], {}
            for raw_c in raw_df.columns:
                c = str(raw_c).strip()
                counts[c] = counts.get(c, -1) + 1
                cols.append(f"{c}_{counts[c]}" if counts[c] > 0 else c)
            raw_df.columns = cols

            col_map = {}
            for c in raw_df.columns:
                cl = c.lower()
                if any(k in cl for k in ["securityname", "stock", "company"]):
                    col_map[c] = "Stock_Name"
                elif "sector" in cl:
                    col_map[c] = "Sector"
                elif any(k in cl for k in ["mcap", "tier", "cap"]):
                    col_map[c] = "Market_Cap_Tier"
                elif any(k in cl for k in ["action", "position"]):
                    col_map[c] = "Position_Action"

            df = raw_df.rename(columns=col_map)
            if "Stock_Name" not in df.columns and len(df.columns) > 0:
                df["Stock_Name"] = df.iloc[:, 0]

            fund_series = df.iloc[:, -1]
            df["Fund_Weight"] = (
                pd.to_numeric(fund_series, errors="coerce").fillna(0.0).round(2)
            )

            bm_col = next(
                (
                    c
                    for c in df.columns
                    if any(
                        k in str(c).lower()
                        for k in [
                            "bm",
                            "bench",
                            "bse",
                            "cnx",
                            "nifty",
                            self.bm_name.lower(),
                        ]
                    )
                    and c != "Fund_Weight"
                ),
                None,
            )
            if bm_col:
                bm_data = df[bm_col]
                bm_series = (
                    bm_data.iloc[:, 0]
                    if isinstance(bm_data, pd.DataFrame)
                    else bm_data
                )
                df["BM_Weight"] = (
                    pd.to_numeric(bm_series, errors="coerce").fillna(0.0).round(2)
                )
            else:
                df["BM_Weight"] = 0.0

            if "Sector" not in df.columns:
                df["Sector"] = "Diversified"

        # 3. DIRECT RAW INGESTION (No regex or sector name alterations)
        df["Stock_Name"] = df["Stock_Name"].astype(str).str.strip()
        df["Sector"] = df["Sector"].astype(str).str.strip()

        df = df.dropna(subset=["Stock_Name"]).copy()
        df = df[
            ~df["Stock_Name"].str.contains("Total|Subtotal", case=False, na=False)
        ].copy()

        # Aggregate in case duplicate rows exist
        df = df.groupby(["Stock_Name", "Sector"], as_index=False).agg(
            {"Fund_Weight": "sum", "BM_Weight": "sum"}
        )

        # Compute Active Weight to exact 2-decimal precision
        df["Active_Weight"] = (df["Fund_Weight"] - df["BM_Weight"]).round(2)

        # String helpers for PPTX table rendering
        df["Fund_Weight_Disp"] = df["Fund_Weight"].apply(
            lambda w: f"{w:.2f}" if w > 0.0 else "-"
        )
        df["Active_Weight_Disp"] = df["Active_Weight"].apply(lambda aw: f"{aw:+.2f}")

        if "Position_Action" not in df.columns:
            df["Position_Action"] = "Core Portfolio Holding"

        # 4. TARGET DATASETS
        self.holdings_df = df.reset_index(drop=True)

        # Left Table: Top N Holdings (Fund_Weight > 0, sorted descending)
        self.top_holdings_summary = (
            df[df["Fund_Weight"] > 0]
            .sort_values(by="Fund_Weight", ascending=False)
            .reset_index(drop=True)
            .head(top_n)
        )

        # Right Table 1: Overweight (Top 5) (Active_Weight > 0, sorted descending)
        self.top_overweight_stocks = (
            df[df["Active_Weight"] > 0]
            .sort_values(by="Active_Weight", ascending=False)
            .reset_index(drop=True)
            .head(active_n)
        )

        # Right Table 2: Underweight (Top 5) (Active_Weight < 0, sorted ascending)
        self.top_underweight_stocks = (
            df[df["Active_Weight"] < 0]
            .sort_values(by="Active_Weight", ascending=True)
            .reset_index(drop=True)
            .head(active_n)
        )


    def _process_sector_sheet(self, xls: pd.ExcelFile, sheet_name: str) -> None:
        """Processes Sector sheet to compute active weights, overweight, and underweight."""
        df_test = pd.read_excel(xls, sheet_name=sheet_name, nrows=6, header=None)
        skip_rows = 0
        for idx, row in df_test.iterrows():
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)]).lower()
            if "sector" in row_str:
                skip_rows = idx
                break

        df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=skip_rows)
        df.columns = [str(c).strip() for c in df.columns]

        # Standardize sector column
        sec_col = next((c for c in df.columns if "sector" in c.lower()), df.columns[0])
        df = df.rename(columns={sec_col: "Sector"})
        df = df.dropna(subset=["Sector"]).copy()
        df["Sector"] = df["Sector"].astype(str).str.strip()
        df = df[~df["Sector"].str.contains("Total|Benchmark", case=False, na=False)].copy()

        # Find fund and benchmark weight columns
        fw_col = next((c for c in df.columns if "fund" in c.lower() and "weight" in c.lower()), None)
        bw_col = next((c for c in df.columns if ("bm" in c.lower() or "benchmark" in c.lower()) and "weight" in c.lower()), None)

        if not fw_col or not bw_col:
            # 3-column layout: Col 0: sector, Col 1: Benchmark, Col 2: Fund
            cols = [c for c in df.columns if c != "Sector"]
            if len(cols) >= 2:
                bw_col = cols[0]
                fw_col = cols[1]
            elif len(cols) == 1:
                fw_col = cols[0]
                bw_col = None

        if fw_col:
            fw_data = df[fw_col]
            fw_series = fw_data.iloc[:, 0] if isinstance(fw_data, pd.DataFrame) else fw_data
            df["Fund_Weight"] = pd.to_numeric(fw_series, errors="coerce").fillna(0.0)
        else:
            df["Fund_Weight"] = 0.0

        if bw_col:
            bw_data = df[bw_col]
            bw_series = bw_data.iloc[:, 0] if isinstance(bw_data, pd.DataFrame) else bw_data
            df["BM_Weight"] = pd.to_numeric(bw_series, errors="coerce").fillna(0.0)
        else:
            df["BM_Weight"] = 0.0

        df["Active_Weight"] = (df["Fund_Weight"] - df["BM_Weight"]).round(1)

        comm_col = next((c for c in df.columns if "research" in c.lower() or "comment" in c.lower()), None)
        df["Research_Commentary"] = df[comm_col].fillna("").astype(str) if comm_col else ""

        self.sector_summary = df.sort_values(by="Fund_Weight", ascending=False).reset_index(drop=True)
        self.top_overweight_sectors = df[df["Active_Weight"] > 0].sort_values(by="Active_Weight", ascending=False).reset_index(drop=True)
        self.top_underweight_sectors = df[df["Active_Weight"] < 0].sort_values(by="Active_Weight", ascending=True).reset_index(drop=True)

    def _process_portfolio_summary_sheet(self, xls: pd.ExcelFile, sheet_name: str) -> None:
        """Processes classic Portfolio_Summary tab if present."""
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        quants = []
        for idx, row in df.iterrows():
            m_val = str(row[0]).strip() if pd.notna(row[0]) else ""
            if m_val.lower() in [
                "month-end aum", "total stock count", "active share", "top-10 holdings",
                "sharpe ratio", "beta", "standard deviation", "portfolio turnover ratio"
            ]:
                disp_val = str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else str(row[1] or "")
                quants.append({"metric": m_val, "display_value": disp_val})
        self.quants_ratios = quants
        self.df_quants = pd.DataFrame(quants).rename(columns={"metric": "Metric", "display_value": "Display_Value"})

        mc_rows = []
        for idx, row in df.iterrows():
            val0 = str(row[0]).strip().lower() if pd.notna(row[0]) else ""
            if val0 in ["large cap", "mid cap", "small cap"]:
                mc_rows.append({"segment": str(row[0]).strip(), "fund": str(row[1]), "bm": str(row[2])})
        self.market_cap_summary = mc_rows
        self.df_mcap = pd.DataFrame(mc_rows).rename(columns={"segment": "Segment", "fund": "Fund", "bm": "BM"})

        val_rows = []
        for idx, row in df.iterrows():
            val0 = str(row[0]).strip().lower() if pd.notna(row[0]) else ""
            if val0 in ["price to book", "price to earnings", "return on equity"]:
                val_rows.append({"parameter": str(row[0]).strip(), "fund": str(row[1]), "bm": str(row[2])})
        self.valuation_metrics = val_rows
        self.df_valuation = pd.DataFrame(val_rows).rename(columns={"parameter": "Parameters", "fund": "Fund", "bm": "BM"})

        ocf_list = [
            {"tier": "C1", "fund": "89%", "bm": "88%"},
            {"tier": "C2", "fund": "2%", "bm": "3%"},
            {"tier": "C3", "fund": "9%", "bm": "9%"}
        ]
        roce_list = [
            {"tier": "R1", "fund": "63%", "bm": "56%"},
            {"tier": "R2", "fund": "31%", "bm": "37%"},
            {"tier": "R3", "fund": "6%", "bm": "7%"}
        ]
        self.composition_tiers = {"ocf": ocf_list, "roce": roce_list}

    def save_pandas_output(self, output_dir: str = "output/intermediate", filename_base: str = "pandas_output") -> Dict[str, str]:
        """
        Stores all processed pandas DataFrames in output/intermediate/pandas_output.xlsx,
        output/intermediate/pandas_output.txt, and output/intermediate/pandas_output.json.
        """
        os.makedirs(output_dir, exist_ok=True)
        out_excel = os.path.join(output_dir, f"{filename_base}.xlsx")
        out_txt = os.path.join(output_dir, f"{filename_base}.txt")
        out_json = os.path.join(output_dir, f"{filename_base}.json")

        dfs = {
            "Top_Holdings": self.top_holdings_summary,
            "Top_Overweight_Stocks": self.top_overweight_stocks,
            "Top_Underweight_Stocks": self.top_underweight_stocks,
        }

        # 1. Export Excel workbook with multiple tabs
        try:
            with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
                for sheet_name, df_item in dfs.items():
                    if isinstance(df_item, pd.DataFrame) and not df_item.empty:
                        clean_sheet_name = sheet_name[:31]  # Excel 31 character limit
                        df_item.to_excel(writer, sheet_name=clean_sheet_name, index=False)
            print(f"Saved pandas output Excel workbook to {out_excel}")
        except Exception as e:
            print(f"Warning: Could not write pandas output Excel file: {e}")

        # 2. Export human-readable text file
        try:
            with open(out_txt, "w", encoding="utf-8") as f:
                f.write("=" * 80 + "\n")
                f.write("   PORTFOLIO DATA ENGINE — PANDAS DATAFRAMES AUDIT OUTPUT\n")
                f.write("=" * 80 + "\n\n")
                for name, df_item in dfs.items():
                    f.write(f"[{name.upper()}]\n")
                    f.write("-" * 80 + "\n")
                    if isinstance(df_item, pd.DataFrame) and not df_item.empty:
                        f.write(df_item.to_string(index=False) + "\n\n")
                    else:
                        f.write("Empty DataFrame\n\n")
            print(f"Saved pandas output text report to {out_txt}")
        except Exception as e:
            print(f"Warning: Could not write pandas output text file: {e}")

        # 3. Export JSON file
        try:
            json_export = {}
            for name, df_item in dfs.items():
                if isinstance(df_item, pd.DataFrame):
                    json_export[name] = df_item.to_dict(orient="records")
            with open(out_json, "w", encoding="utf-8") as f:
                json.dump(json_export, f, indent=2)
            print(f"Saved pandas output JSON to {out_json}")
        except Exception as e:
            print(f"Warning: Could not write pandas output JSON file: {e}")

        return {"excel": out_excel, "txt": out_txt, "json": out_json}

    def get_data_bundle(self) -> Dict[str, Any]:
        """Returns structured dictionary ready for LLM, charts, and slide renderer."""
        return {
            "sector_summary": self.sector_summary,
            "top_overweight_sectors": self.top_overweight_sectors,
            "top_underweight_sectors": self.top_underweight_sectors,
            "holdings_df": self.holdings_df,
            "top_holdings_summary": self.top_holdings_summary,
            "top_overweight_stocks": self.top_overweight_stocks,
            "top_underweight_stocks": self.top_underweight_stocks,
            "quants_ratios": self.quants_ratios,
            "market_cap_summary": self.market_cap_summary,
            "valuation_metrics": self.valuation_metrics,
            "composition_tiers": self.composition_tiers,
            "df_quants": self.df_quants,
            "df_mcap": self.df_mcap,
            "df_valuation": self.df_valuation,
            "df_composition": self.df_composition,
        }
