import pandas as pd
from openpyxl import load_workbook

# === File paths and sheet ===
EXCEL_PATH = "/Users/srikanth/Library/CloudStorage/OneDrive-Personal/DataWorks/DataWorks Progress Plan.xlsx"
SHEET_NAME = "2025 - October"
README_PATH = "README.md"

# === Step 1: Load Excel and clean table ===
wb = load_workbook(EXCEL_PATH, data_only=True)
ws = wb[SHEET_NAME]

# Convert sheet to list of lists
data = list(ws.values)

# Find the first non-empty row (header)
header_row = None
for i, row in enumerate(data):
    if any(cell is not None for cell in row):
        header_row = i
        break

# Extract headers and data cleanly
columns = [c for c in data[header_row] if c is not None]
data_rows = [
    [cell for cell in row[:len(columns)]]
    for row in data[header_row + 1:]
    if any(cell is not None for cell in row)
]

df = pd.DataFrame(data_rows, columns=columns)

# === Step 2: Replace hyperlinks with Markdown links ===
# Loop through Excel cells directly to preserve hyperlinks
for r_idx, row in enumerate(ws.iter_rows(min_row=header_row + 2, max_col=len(columns))):
    for c_idx, cell in enumerate(row):
        if cell.hyperlink:
            display_text = str(cell.value).strip() if cell.value else cell.hyperlink.target
            df.iat[r_idx, c_idx] = f"[{display_text}]({cell.hyperlink.target})"

# === Step 3: Convert DataFrame to Markdown ===
markdown_table = df.to_markdown(index=False)

# === Step 4: Replace table in README ===
with open(README_PATH, "r", encoding="utf-8") as f:
    content = f.read()

start_marker = "<!-- START_TABLE -->"
end_marker = "<!-- END_TABLE -->"

if start_marker not in content or end_marker not in content:
    raise ValueError("README.md must contain <!-- START_TABLE --> and <!-- END_TABLE --> markers.")

before = content.split(start_marker)[0]
after = content.split(end_marker)[-1]

new_content = (
    before
    + start_marker
    + "\n"
    + markdown_table
    + "\n"
    + end_marker
    + after
)

# === Step 5: Write updated README ===
with open(README_PATH, "w", encoding="utf-8") as f:
    f.write(new_content)

print("✅ README.md successfully updated with the latest Excel data and hyperlinks!")
