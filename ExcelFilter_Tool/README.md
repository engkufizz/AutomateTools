# Universal Excel/CSV Filter

A **Python GUI tool** that lets you load Excel (`.xlsx`) and CSV (`.csv`) files, define custom filtering rules, and export the results to new Excel/CSV files.
It automatically detects headers, handles multi-row headers, and supports both simple and advanced filtering operations.

---

## ✨ Features

* **File Support**: Excel (`.xlsx`) and CSV (`.csv`)
* **Automatic Header Detection**: Handles single-row, multi-row, or missing headers
* **Filtering Options**:

  * String operations: `is`, `is_not`, `contains_any`, `not_contains_any`, `regex_any`, `not_regex`
  * Numeric operations: `gt`, `gte`, `lt`, `lte`, `between`
  * External list file matching: `in_list_file`, `not_in_list_file`
* **Combine Rules**: Use **AND** or **OR** logic
* **Output Options**:

  * Save results as Excel or CSV
  * Separate sheets for multiple Excel tabs
  * Append source sheet name to merged outputs
* **Column Normalisation**: Automatically deduplicates/cleans column names
* **GUI Powered by Tkinter** – simple, lightweight, and cross-platform

---

## 🚀 Installation

1. Install Python 3.8+
2. Install dependencies:

   ```bash
   pip install pandas openpyxl
   ```

   (Tkinter comes bundled with most Python installations.)

---

## 🖥️ Usage

Run the program:

```bash
python your_script.py
```

### GUI Steps:

1. **Load File**

   * Select your Excel/CSV file
   * Optionally process all sheets (Excel only)

2. **Set Options**

   * Choose whether to normalise column names
   * Select how to combine multiple rules (AND/OR)
   * Decide if matching rows should be kept or excluded
   * Pick output format (Excel, CSV, or both)

3. **Define Rules**

   * Select column
   * Choose operator (`contains_any`, `is`, `gt`, etc.)
   * Enter values or attach a list file if required

4. **Apply Filters & Export**

   * Results are saved with a timestamped filename in your chosen directory

---

## 📂 Example

Suppose you have an Excel file with customer records and you want to:

* Keep rows where `Status` **contains "Active"**
* Exclude rows where `Age < 18`

You would:

* Add rule: `Status → contains_any → ["Active"]`
* Add rule: `Age → gte → 18`
* Combine rules with **AND**
* Export results to Excel

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
| Numeric | `gt`, `gte`, `lt`, `lte` | Greater/less than (with float support) |
|         | `between`                | Between min and max values             |

---

## 🛠 Development Notes

* Header detection scans first 50 rows and scores them based on "header-likeness"
* Multi-row headers may be merged (e.g., `Region | Sales` → `Region Sales`)
* Deduplication ensures unique column names (`Name`, `Name__1`, `Name__2`)

---

## 📜 License

MIT License – free to use, modify, and distribute.

