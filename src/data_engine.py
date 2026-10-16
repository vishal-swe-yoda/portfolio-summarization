"""
PortfolioDataEngine: Ingests Excel workbooks, cleans and computes portfolio metrics,
sector allocations, active weights, and stock holdings.
"""

from typing import Dict, Any, List, Optional
import os
import pandas as pd
import yaml


class PortfolioDataEngine:
    def __init__(self, excel_path: str, metric_rules_path: Optional[str] = None):
        self.excel_path = excel_path
        self.metric_rules = {}
        if metric_rules_path and os.path.exists(metric_rules_path):
            with open(metric_rules_path, "r", encoding="utf-8") as f:
                rules_data = yaml.safe_load(f)
                self.metric_rules = rules_data.get("metrics", []) if rules_data else []

        self.sheets: Dict[str, pd.DataFrame] = {}
        self.sector_summary: pd.DataFrame = pd.DataFrame()
        self.top_overweight_sectors: pd.DataFrame = pd.DataFrame()
        self.top_underweight_sectors: pd.DataFrame = pd.DataFrame()
        self.holdings_df: pd.DataFrame = pd.DataFrame()
        self.top_holdings_summary: pd.DataFrame = pd.DataFrame()
        self.active_stock_positions: pd.DataFrame = pd.DataFrame()
        self.quants_ratios: List[Dict[str, Any]] = []
        self.market_cap_summary: List[Dict[str, Any]] = []
        self.valuation_metrics: List[Dict[str, Any]] = []
        self.composition_tiers: Dict[str, Any] = {}

    def load_data(self) -> "PortfolioDataEngine":
        """Loads and processes all required sheets from the workbook."""
        xls = pd.ExcelFile(self.excel_path)
        sheet_names = xls.sheet_names

        # Load Sector Allocation
        sector_sheet = next((s for s in sheet_names if "sector" in s.lower()), None)
        if sector_sheet:
            self._process_sector_allocation(xls, sector_sheet)

        # Load Holdings Data
        holdings_sheet = next((s for s in sheet_names if "holding" in s.lower()), None)
        if holdings_sheet:
            self._process_holdings_data(xls, holdings_sheet)

        # Load Portfolio Summary
        summary_sheet = next((s for s in sheet_names if "summary" in s.lower()), None)
        if summary_sheet:
            self._process_portfolio_summary(xls, summary_sheet)

        return self

    def _process_sector_allocation(self, xls: pd.ExcelFile, sheet_name: str) -> None:
        """Processes Sector_Allocation tab to compute active weights and rankings."""
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        # Locate header row containing 'Sector' and 'Fund'
        header_row_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)]).lower()
            if "sector" in row_str and ("fund" in row_str or "weight" in row_str):
                header_row_idx = idx
                break

        df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_row_idx)
        df.columns = [str(c).strip() for c in df.columns]

        # Standardize column names
        col_map = {}
        for c in df.columns:
            cl = c.lower()
            if "sector" in cl:
                col_map[c] = "Sector"
            elif "fund" in cl and "weight" in cl:
                col_map[c] = "Fund_Weight"
            elif "bm" in cl or "benchmark" in cl:
                col_map[c] = "BM_Weight"
            elif "active" in cl:
                col_map[c] = "Active_Weight"
            elif "research" in cl or "comment" in cl or "overview" in cl:
                col_map[c] = "Research_Commentary"

        df = df.rename(columns=col_map)

        if "Sector" not in df.columns:
            df["Sector"] = df.iloc[:, 0]
        if "Fund_Weight" not in df.columns and len(df.columns) > 1:
            df["Fund_Weight"] = df.iloc[:, 1]
        if "BM_Weight" not in df.columns and len(df.columns) > 2:
            df["BM_Weight"] = df.iloc[:, 2]

        # Clean rows
        df = df[df["Sector"].notna()].copy()
        df["Sector"] = df["Sector"].astype(str).str.strip()
        df = df[~df["Sector"].str.contains("Total|Benchmark", case=False, na=False)].copy()

        # Coerce numeric weights
        df["Fund_Weight"] = pd.to_numeric(df["Fund_Weight"], errors="coerce").fillna(0.0)
        df["BM_Weight"] = pd.to_numeric(df["BM_Weight"], errors="coerce").fillna(0.0)
        
        # Calculate active weight = Fund_Weight - BM_Weight
        df["Active_Weight"] = (df["Fund_Weight"] - df["BM_Weight"]).round(1)

        if "Research_Commentary" not in df.columns:
            df["Research_Commentary"] = ""
        else:
            df["Research_Commentary"] = df["Research_Commentary"].fillna("").astype(str)

        # Sort sector summary by Fund_Weight descending
        self.sector_summary = df.sort_values(by="Fund_Weight", ascending=False).reset_index(drop=True)

        # Overweight & Underweight
        ow = df[df["Active_Weight"] > 0].sort_values(by="Active_Weight", ascending=False).reset_index(drop=True)
        self.top_overweight_sectors = ow

        uw = df[df["Active_Weight"] < 0].sort_values(by="Active_Weight", ascending=True).reset_index(drop=True)
        self.top_underweight_sectors = uw

    def _process_holdings_data(self, xls: pd.ExcelFile, sheet_name: str) -> None:
        """Processes Holdings_Data tab for top holdings and active positions."""
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        header_row_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)]).lower()
            if "stock" in row_str or "company" in row_str or "holding" in row_str:
                header_row_idx = idx
                break

        df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_row_idx)
        df.columns = [str(c).strip() for c in df.columns]

        col_map = {}
        for c in df.columns:
            cl = c.lower()
            if "stock" in cl or "company" in cl or "security" in cl:
                col_map[c] = "Stock_Name"
            elif "sector" in cl:
                col_map[c] = "Sector"
            elif "tier" in cl or "cap" in cl:
                col_map[c] = "Market_Cap_Tier"
            elif "fund" in cl and "weight" in cl:
                col_map[c] = "Fund_Weight"
            elif "bm" in cl or "benchmark" in cl:
                col_map[c] = "BM_Weight"
            elif "active" in cl:
                col_map[c] = "Active_Weight"
            elif "action" in cl or "note" in cl or "position" in cl:
                col_map[c] = "Position_Action"

        df = df.rename(columns=col_map)
        if "Stock_Name" not in df.columns:
            df["Stock_Name"] = df.iloc[:, 0]

        df = df[df["Stock_Name"].notna()].copy()
        df["Stock_Name"] = df["Stock_Name"].astype(str).str.strip()
        df = df[~df["Stock_Name"].str.contains("Total|Subtotal", case=False, na=False)].copy()

        df["Fund_Weight"] = pd.to_numeric(df.get("Fund_Weight", 0.0), errors="coerce").fillna(0.0)
        df["BM_Weight"] = pd.to_numeric(df.get("BM_Weight", 0.0), errors="coerce").fillna(0.0)
        df["Active_Weight"] = (df["Fund_Weight"] - df["BM_Weight"]).round(1)

        if "Position_Action" not in df.columns:
            df["Position_Action"] = ""
        else:
            df["Position_Action"] = df["Position_Action"].fillna("").astype(str)

        self.holdings_df = df.reset_index(drop=True)
        self.top_holdings_summary = df.sort_values(by="Fund_Weight", ascending=False).reset_index(drop=True)
        
        # Sort active stock positions by absolute active weight descending
        df_active = df.copy()
        df_active["abs_aw"] = df_active["Active_Weight"].abs()
        self.active_stock_positions = df_active.sort_values(by="abs_aw", ascending=False).drop(columns=["abs_aw"]).reset_index(drop=True)

    def _process_portfolio_summary(self, xls: pd.ExcelFile, sheet_name: str) -> None:
        """Processes summary metrics, market cap tiers, valuations, and composition tiers."""
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        # 1. Quants & Ratios (scan for 'Month-end AUM' or 'Metric')
        quants = []
        for idx, row in df.iterrows():
            m_val = str(row[0]).strip() if pd.notna(row[0]) else ""
            if m_val.lower() in [
                "month-end aum",
                "total stock count",
                "active share",
                "top-10 holdings",
                "sharpe ratio",
                "beta",
                "standard deviation",
                "portfolio turnover ratio",
            ]:
                raw_val = row[1] if len(row) > 1 and pd.notna(row[1]) else None
                disp_val = str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else str(raw_val or "")
                quants.append({
                    "metric": m_val,
                    "raw_value": raw_val,
                    "display_value": disp_val
                })

        # Apply metric rules if defined
        quants = self._apply_metric_rules(quants)
        self.quants_ratios = quants

        # 2. Market Cap Summary
        mc_rows = []
        for idx, row in df.iterrows():
            val0 = str(row[0]).strip().lower() if pd.notna(row[0]) else ""
            if val0 in ["large cap", "mid cap", "small cap"]:
                seg = str(row[0]).strip()
                fund_val = self._format_pct_or_str(row[1])
                bm_val = self._format_pct_or_str(row[2])
                mc_rows.append({"segment": seg, "fund": fund_val, "bm": bm_val})
        self.market_cap_summary = mc_rows

        # 3. Valuation Parameters
        val_rows = []
        for idx, row in df.iterrows():
            val0 = str(row[0]).strip().lower() if pd.notna(row[0]) else ""
            if val0 in ["price to book", "price to earnings", "return on equity"]:
                param = str(row[0]).strip()
                fund_val = self._format_decimal_or_pct(row[1])
                bm_val = self._format_decimal_or_pct(row[2])
                val_rows.append({"parameter": param, "fund": fund_val, "bm": bm_val})
        self.valuation_metrics = val_rows

        # 4. Composition Tiers (OCF & RoCE)
        # Scan for C1..C3 and R1..R3 in the sheet
        ocf_list = []
        roce_list = []
        for idx, row in df.iterrows():
            # Look across columns for C1, C2, C3 and R1, R2, R3
            for c_idx in range(len(row)):
                cell_val = str(row[c_idx]).strip() if pd.notna(row[c_idx]) else ""
                if cell_val in ["C1", "C2", "C3"] and c_idx + 2 < len(row):
                    f_val = self._format_pct_or_str(row[c_idx + 1])
                    b_val = self._format_pct_or_str(row[c_idx + 2])
                    if not any(x["tier"] == cell_val for x in ocf_list):
                        ocf_list.append({"tier": cell_val, "fund": f_val, "bm": b_val})
                elif cell_val in ["R1", "R2", "R3"] and c_idx + 2 < len(row):
                    f_val = self._format_pct_or_str(row[c_idx + 1])
                    b_val = self._format_pct_or_str(row[c_idx + 2])
                    if not any(x["tier"] == cell_val for x in roce_list):
                        roce_list.append({"tier": cell_val, "fund": f_val, "bm": b_val})

        # Ensure sorted C1..C3 and R1..R3
        ocf_list.sort(key=lambda x: x["tier"])
        roce_list.sort(key=lambda x: x["tier"])
        self.composition_tiers = {"ocf": ocf_list, "roce": roce_list}

    def _apply_metric_rules(self, quants: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Recalculates or formats metrics according to metric_rules.yaml."""
        if not self.metric_rules:
            return quants

        rule_dict = {r.get("source_column", "").lower(): r for r in self.metric_rules}
        updated_quants = []

        existing_metrics = {q["metric"].lower(): q for q in quants}

        for r in self.metric_rules:
            label = r.get("label", "")
            src = r.get("source_column", "").lower()
            fmt = r.get("format", "{}")
            agg = r.get("aggregation", "direct")

            existing = existing_metrics.get(src, {})
            raw_val = existing.get("raw_value")
            disp_val = existing.get("display_value", "")

            if agg == "count_holdings" and not self.holdings_df.empty:
                count_val = len(self.holdings_df)
                disp_val = fmt.format(count_val)
                raw_val = count_val
            elif agg == "sum_top_10_weights" and not self.top_holdings_summary.empty:
                top10_sum = self.top_holdings_summary.head(10)["Fund_Weight"].sum()
                disp_val = fmt.format(top10_sum) if "{" in fmt else f"{top10_sum:.0f}%"
                raw_val = top10_sum
            elif agg == "half_sum_abs_active_weight" and not self.holdings_df.empty:
                # Active share = 0.5 * sum(|fund - bm|)
                half_sum = 0.5 * (self.holdings_df["Fund_Weight"] - self.holdings_df["BM_Weight"]).abs().sum()
                disp_val = fmt.format(half_sum) if "{" in fmt else f"{half_sum:.0f}%"
                raw_val = half_sum
            elif disp_val:
                pass
            elif raw_val is not None:
                try:
                    disp_val = fmt.format(raw_val)
                except Exception:
                    disp_val = str(raw_val)

            updated_quants.append({
                "metric": label,
                "raw_value": raw_val,
                "display_value": disp_val
            })

        return updated_quants if updated_quants else quants

    @staticmethod
    def _format_pct_or_str(val: Any) -> str:
        if pd.isna(val) or val is None:
            return "0%"
        s = str(val).strip()
        if s.endswith("%"):
            return s
        try:
            num = float(s)
            if num <= 1.0 and num > 0:
                return f"{int(round(num * 100))}%"
            else:
                return f"{int(round(num))}%"
        except Exception:
            return s

    @staticmethod
    def _format_decimal_or_pct(val: Any) -> str:
        if pd.isna(val) or val is None:
            return "0.00"
        s = str(val).strip()
        if s.endswith("%"):
            return s
        try:
            num = float(s)
            return f"{num:.2f}"
        except Exception:
            return s

    def get_data_bundle(self) -> Dict[str, Any]:
        """Returns structured dictionary ready for LLM, charts, and slide renderer."""
        return {
            "sector_summary": self.sector_summary,
            "top_overweight_sectors": self.top_overweight_sectors,
            "top_underweight_sectors": self.top_underweight_sectors,
            "holdings_df": self.holdings_df,
            "top_holdings_summary": self.top_holdings_summary,
            "active_stock_positions": self.active_stock_positions,
            "quants_ratios": self.quants_ratios,
            "market_cap_summary": self.market_cap_summary,
            "valuation_metrics": self.valuation_metrics,
            "composition_tiers": self.composition_tiers,
        }
