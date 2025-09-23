# Universal Excel/CSV Filter

A **Python GUI tool** for loading Excel (`.xlsx`) and CSV (`.csv`) files, building advanced filtering rules, and exporting filtered data to new Excel/CSV files. This tool automatically detects headers (including multi-row and missing headers), supports a wide range of filtering operations, and provides a convenient export workflow.

---

## ✨ Features

* **Supported Files**: Excel (`.xlsx`) and CSV (`.csv`)
* **Header Detection**: Handles single-row, multi-row, or missing headers automatically
* **Flexible Filtering**:
  * String: `is`, `is_not`, `contains_any`, `not_contains_any`, `regex_any`, `not_regex`, `in_list_file`, `not_in_list_file`, `is_empty`, `not_empty`
  * Numeric: `gt`, `gte`, `lt`, `lte`, `between`
* **Rule Logic**: Combine multiple rules using **AND** or **OR**
* **Export Options**:
  * Save filtered data as Excel or CSV
  * Option to export each Excel sheet separately
  * Optionally add source sheet names to merged outputs
* **Column Normalisation**: Cleans and deduplicates column names automatically
* **Intuitive Tkinter GUI**: Lightweight, fast, and cross-platform

---

## 🚀 Installation

1. Install Python 3.8 or newer
2. Install required packages:

   ```bash
   pip install pandas openpyxl
   ```

   (Tkinter is included with most Python installations.)

---

## 🖥️ Usage

Start the GUI application:

```bash
python ExcelFilter_Tool.py
```

### Workflow:

1. **Load File**
   * Choose an Excel or CSV file to load
   * Optionally process all sheets (for Excel files)

2. **Configure Options**
   * Decide whether to normalise column names
   * Select rule logic (AND/OR)
   * Choose to keep or exclude matching rows
   * Select output format(s): Excel, CSV, or both

3. **Build Filtering Rules**
   * Pick a column
   * Select a filtering operator (e.g., `contains_any`, `is_empty`, `gte`)
   * Enter values, numeric thresholds, or attach a list file as needed

4. **Preview & Export**
   * Preview the number of rows that will be kept or excluded
   * Export filtered results with a timestamped filename

---

## 📂 Example

Suppose you want to:

* Keep rows where `Status` **contains "Active"**
* Exclude rows where `Age` is less than 18

You would:

* Add rule: `Status → contains_any → ["Active"]`
* Add rule: `Age → gte → 18`
* Set rule combination to **AND**
* Export the results to Excel or CSV

---

## ⚙️ Supported Operators

| Type    | Operator                 | Description                            |
| ------- | ------------------------ | -------------------------------------- |
| String  | `is` / `is_not`          | Exact match or not match               |
|         | `contains_any`           | Contains any of given substrings       |
|         | `not_contains_any`       | Does not contain substrings            |
|         | `regex_any`              | Matches regex patterns                 |
|         | `not_regex`              | Does not match regex                   |
|         | `in_list_file`           | Matches values from external list file |
|         | `not_in_list_file`       | Excludes values in external list file  |
|         | `is_empty`               | Cell is empty or blank                 |
|         | `not_empty`              | Cell is not empty                      |
| Numeric | `gt`, `gte`, `lt`, `lte` | Greater/less than (numeric)            |
|         | `between`                | Within a specified numeric range       |

---

## 🛠 Notes

* Header detection scans the top 50 rows to determine the most likely header row(s)
* Multi-row headers are merged automatically for clarity
* Column names are deduplicated to avoid ambiguity (e.g., `Name`, `Name__1`)
* Filtering supports both string and numeric logic, as well as matching from external lists

---

## 📜 License

MIT License – free to use, modify, and distribute.
