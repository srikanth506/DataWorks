import pandas as pd
from openpyxl import load_workbook

# === Config ===
EXCEL_PATH = "/Users/srikanth/Library/CloudStorage/OneDrive-Personal/DataWorks/DataWorks Progress Plan.xlsx"
README_PATH = "README.md"

SHEETS_IN_ORDER = [
    "2025 - October",
    "2025 - November",
    "2025 - December"
]

START_TABLE = "<!-- START_TABLE -->"
END_TABLE = "<!-- END_TABLE -->"
START_SUMMARY = "<!-- START_SUMMARY -->"
END_SUMMARY = "<!-- END_SUMMARY -->"


# === Helper: Read one sheet with hyperlinks ===
def read_sheet(wb, sheet_name):
    ws = wb[sheet_name]
    data = list(ws.values)

    header_row = next(i for i, r in enumerate(data) if any(r))
    columns = [c for c in data[header_row] if c is not None]

    rows = [
        row[:len(columns)]
        for row in data[header_row + 1:]
        if any(row)
    ]

    df = pd.DataFrame(rows, columns=columns)

    for r_idx, row in enumerate(ws.iter_rows(min_row=header_row + 2, max_col=len(columns))):
        for c_idx, cell in enumerate(row):
            if cell.hyperlink:
                text = str(cell.value).strip() if cell.value else cell.hyperlink.target
                df.iat[r_idx, c_idx] = f"[{text}]({cell.hyperlink.target})"

    return df


# === Load workbook ===
wb = load_workbook(EXCEL_PATH, data_only=True)

final_rows = []
summary_df = []

for sheet in SHEETS_IN_ORDER:
    if sheet not in wb.sheetnames:
        continue

    year, month = sheet.split(" - ")
    df = read_sheet(wb, sheet)

    # Month separator row
    separator = {col: "" for col in df.columns}
    separator[df.columns[0]] = f"**{month} {year}**"
    final_rows.append(separator)

    final_rows.extend(df.to_dict(orient="records"))
    summary_df.append(df)

# === Final DataFrame ===
final_df = pd.DataFrame(final_rows)

# === Generate summary ===
summary_all = pd.concat(summary_df, ignore_index=True)

summary_md = f"""
## 📊 Progress Summary

- **Total Days Logged:** {len(summary_all)}
- **SQL Topics Covered:** {summary_all['SQL'].astype(bool).sum()} days
- **Big Data Activities:** {summary_all['Big Data'].astype(bool).sum()} days
- **Data Science Activities:** {summary_all['Data Science'].astype(bool).sum()} days
- **Job Search Activities:** {summary_all['Job Search'].astype(bool).sum()} days
""".strip()

# === Markdown table ===
table_md = final_df.to_markdown(index=False)

# === Update README ===
with open(README_PATH, "r", encoding="utf-8") as f:
    content = f.read()

# Summary replace
if START_SUMMARY in content and END_SUMMARY in content:
    before = content.split(START_SUMMARY)[0]
    after = content.split(END_SUMMARY)[-1]
    content = f"{before}{START_SUMMARY}\n{summary_md}\n{END_SUMMARY}{after}"
else:
    content = f"{summary_md}\n\n{content}"

# Table replace
before = content.split(START_TABLE)[0]
after = content.split(END_TABLE)[-1]

content = f"{before}{START_TABLE}\n{table_md}\n{END_TABLE}{after}"

with open(README_PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("✅ README updated with month separators and summary")
