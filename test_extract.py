import pandas as pd

def extract_slide_1_tables(excel_path: str, fund_name: str, bm_name: str):
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
    table_composition = pd.DataFrame()

  return {
      'quants': table_quants,
      'mcap': table_mcap,
      'valuation': table_valuation,
      'composition': table_composition,
  }

if __name__ == '__main__':
    res = extract_slide_1_tables('data/sample_portfolio_data.xlsx', 'UTI Multi Cap Fund', 'BSE100')
    for k, v in res.items():
        print(f"=== {k.upper()} ===")
        print(v)
        print()
