import pandas as pd

excel_path = 'data/sample_portfolio_data.xlsx'
df_sector = pd.read_excel(excel_path, sheet_name='Sector', skiprows=4)
df_sector.columns = [str(c).strip() for c in df_sector.columns]
df_sector = df_sector.dropna(subset=['sector']).copy()
df_sector = df_sector[~df_sector['sector'].str.contains('Total', case=False)]
print("Sector head:")
print(df_sector.head(10))
