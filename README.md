🧹 Excel Data Cleaner — Python Automation

Automatically clean and standardize messy Excel files using Python.

What It Does

- ✅ Removes duplicate rows
- ✅ Strips extra whitespace from all cells
- ✅ Standardizes text (Title Case for names, cities, departments)
- ✅ Removes completely empty rows
- ✅ Generates a **Cleaning Report** sheet showing what was fixed
- ✅ Outputs a professionally formatted, color-coded Excel file

Before vs After

|   Issue    |        Before        |   After  |
|------------|----------------------|----------|
| Duplicates | 213 rows             | 200 rows |
| Names      | "  ahmed  ", "AHMED" | "Ahmed"  |
| Cities     | "cairo", "CAIRO "    | "Cairo"  |
| Empty rows | Present              | Removed  |

Requirements

```bash
pip install pandas openpyxl
```
Usage

```python
from data_cleaner import clean_excel_file

clean_excel_file("your_dirty_file.xlsx", "clean_output.xlsx")
```

Sample Data

- `dirty_sales_data.xlsx` — 213 rows with duplicates, inconsistent formatting, missing values

Output

Each cleaned file contains:
1. **Clean Data** sheet — formatted, organized data
2. **Cleaning Report** sheet — summary of changes made
