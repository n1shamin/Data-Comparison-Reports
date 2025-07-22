import sqlite3
import pandas as pd
import os
from xlsxwriter.utility import xl_col_to_name
#change the database paths to yours
new_db = r"C:\Users\Norah\Downloads\IFsDataImport (29).db"
hist_db = r"C:\IFs\DATA\IFsHistSeries.db"

new_conn = sqlite3.connect(new_db)
hist_conn = sqlite3.connect(hist_db)

table_query = "SELECT name FROM sqlite_master WHERE type='table' AND name != 'DataDict';"
tables = pd.read_sql(table_query, new_conn)['name'].tolist()

output_file = "DataComparisonReportBaseYearFAOLand.xlsx"
first_write = True

for table_name in tables:
    print(f"Processing table: {table_name}")

    try:
        new_df = pd.read_sql(f"SELECT * FROM [{table_name}]", new_conn)
        hist_df = pd.read_sql(f"SELECT * FROM [{table_name}]", hist_conn)
    except Exception as e:
        print(f"Skipping {table_name}: {e}")
        continue

    if 'Country' not in new_df.columns or 'Country' not in hist_df.columns:
        continue

    new_df = new_df.set_index('Country').sort_index()
    hist_df = hist_df.set_index('Country').sort_index()

#change year if needed
    base_year = '2020'
    if base_year not in new_df.columns or base_year not in hist_df.columns:
        continue

    hist_data = hist_df[[base_year]].copy()
    new_data = new_df[[base_year]].copy()

    abs_change = pd.DataFrame(index=new_df.index)
    pct_change = pd.DataFrame(index=new_df.index)
    abs_change[base_year] = new_df[base_year] - hist_df[base_year]
    pct_change[base_year] = (
        (new_df[base_year] - hist_df[base_year]) / hist_df[base_year].replace(0, pd.NA)
    ) * 100
  
    hist_renamed = hist_data.rename(columns={base_year: f"old{base_year}"})
    new_renamed = new_data.rename(columns={base_year: f"new{base_year}"})
    abs_renamed = abs_change.rename(columns={base_year: "Absolute Change"})
    pct_renamed = pct_change.rename(columns={base_year: "Percent Change"})
    
    merged_df = hist_renamed.merge(new_renamed, left_index=True, right_index=True)\
                            .merge(abs_renamed, left_index=True, right_index=True)\
                            .merge(pct_renamed, left_index=True, right_index=True)

    merged_df.reset_index(inplace=True)
    with pd.ExcelWriter(output_file, engine='openpyxl', mode="a" if os.path.exists(output_file) else "w") as writer:
        merged_df.to_excel(writer, sheet_name=table_name, index=False)