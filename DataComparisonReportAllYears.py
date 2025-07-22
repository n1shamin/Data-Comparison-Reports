import sqlite3
import pandas as pd
import os
from xlsxwriter.utility import xl_col_to_name

#change the database paths to yours
new_db = r"C:\IFs\RUNFILES\IFsDataImport (23).db"
hist_db = r"C:\IFs\DATA\IFsHistSeries.db"

new_conn = sqlite3.connect(new_db)
hist_conn = sqlite3.connect(hist_db)

table_query = "SELECT name FROM sqlite_master WHERE type='table' AND name != 'DataDict';"
tables = pd.read_sql(table_query, new_conn)['name'].tolist()

#change folder name as needed
output_folder = "IMF changes"
os.makedirs(output_folder, exist_ok=True)

for table_name in tables:
    print(f"Processing table: {table_name}")

    try:
        new_df = pd.read_sql(f"SELECT * FROM [{table_name}]", new_conn)
        hist_df = pd.read_sql(f"SELECT * FROM [{table_name}]", hist_conn)
    except Exception as e:
        print(f"sskipping {table_name}: {e}")
        continue

    if 'Country' not in new_df.columns or 'Country' not in hist_df.columns:
        continue

    year_cols = sorted(
        [col for col in new_df.columns if col in hist_df.columns and col.isdigit()]
    )
    if not year_cols:
        continue

    hist_data = hist_df[['Country'] + year_cols].copy()
    new_data = new_df[['Country'] + year_cols].copy()

    hist_data.set_index('Country', inplace=True)
    new_data.set_index('Country', inplace=True)
    new_data = new_data.reindex(hist_data.index)

    abs_change = new_data - hist_data
    pct_change = (abs_change / hist_data.replace(0, pd.NA)) * 100
    pct_change = pct_change.fillna(pd.NA)

    hist_data.reset_index(inplace=True)
    new_data.reset_index(inplace=True)
    abs_change.reset_index(inplace=True)
    pct_change.reset_index(inplace=True)

    output_file = os.path.join(output_folder, f"{table_name[:31]}.xlsx")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        hist_data.to_excel(writer, sheet_name='Historical', index=False)
        new_data.to_excel(writer, sheet_name='New', index=False)
        abs_change.to_excel(writer, sheet_name='Absolute Change', index=False)
        pct_change.to_excel(writer, sheet_name='Percent Change', index=False)
