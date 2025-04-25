########################################
# we-serve/main.py                     #
# Dan Nelson:  dan.nelson1@gmail.com   #
########################################

import pandas as pd
from pathlib import Path
from helpers.style_cells import styling_wb

filepath = Path.cwd() / "data" / "final_report.xlsx"
resource_path = Path(Path.cwd(), "data", "accounts.xlsx")
transactions_path = Path(Path.cwd(), "data", "Ledger 25.xlsx")

# Load the two dfs
resource = pd.read_excel(resource_path, index_col=0)
transactions = pd.read_excel(transactions_path, index_col=0)

print(resource_path)
print(filepath)
print(transactions_path)

# Look at pivot tables in pandas
grouped = transactions.groupby(by="BUDGET_LINE")["AMOUNT"].sum()

# Put them together
result = pd.concat([resource, grouped], axis=1, join="outer")
result = result.fillna(0)  # insert 0 for all null cells

# Save to disk
result.to_excel(filepath, sheet_name="final")

styling_wb(filepath)
