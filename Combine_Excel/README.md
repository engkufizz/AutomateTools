# Universal Excel Combiner

A Python tool (with both **CLI** and **GUI**) to **combine multiple Excel (`.xlsx`) and CSV files** into a single structured dataset.
It uses **robust header auto-detection** to handle messy files with inconsistent headers, empty rows, or multi-row titles.

---

## ✨ Features

* 🔍 **Automatic header detection**
  Detects the most likely header row using heuristics (header-ish tokens, uniqueness, empties, etc.).

* 🧩 **Multi-row header merging**
  Handles cases where headers span 2 rows and merges them intelligently.

* 📑 **Duplicate header handling**
  Automatically deduplicates column names (`Name`, `Name__1`, …).

* 📂 **Multiple file support**
  Supports both `.xlsx` (all sheets) and `.csv`.

* 🛠️ **Column normalisation**
  Optionally trims spaces and unifies formatting.

* 📊 **Canonical schema alignment**
  Optionally aligns headerless or weakly-detected files to the most common schema.

* 📎 **Metadata tracking**
  Adds `source_file` and `source_sheet` columns (optional) so you always know where each row came from.

* 💾 **Flexible output**

  * Excel (`.xlsx`) with single sheet or per-source sheets + “All”
  * CSV (`.csv`)
  * Or both

* 🖥️ **Dual interface**

  * **CLI** for automation
  * **Tkinter GUI** for point-and-click usage

---

## 📦 Installation

1. Clone or download this repository.
2. Install dependencies:

   ```bash
   pip install pandas openpyxl
   ```

   > Tkinter is part of the Python standard library (no extra install needed on most systems).

---

## 🚀 Usage

### CLI

Run directly from terminal:

```bash
python combiner.py --files data1.xlsx data2.csv --outdir ./output --basename merged
```

#### Options:

* `--files` *(required)*: Input files (`.xlsx`, `.csv`).
* `--outdir`: Output directory (default: `.`).
* `--basename`: Base name for output files (default: `combined`).
* `--no-metadata`: Disable `source_file` / `source_sheet` columns.
* `--no-normalise`: Keep raw column names.
* `--align-headerless`: Try aligning headerless/weak files to the most common schema.
* `--format`: `xlsx`, `csv`, or `both` (default: `xlsx`).
* `--separate-sheets`: For Excel, create one sheet per source sheet plus an "All" sheet.

Example:

```bash
python combiner.py --files file1.xlsx file2.csv --format both --separate-sheets --basename report
```

---

### GUI

Simply run:

```bash
python combiner.py
```

The Tkinter GUI will open with options to:

1. Select multiple input files
2. Choose output directory
3. Configure options (metadata, normalisation, output format, etc.)
4. Generate merged Excel/CSV file(s)

---

## 📝 Example Workflow

Input:

* `Sales_Q1.xlsx` (multiple sheets with messy headers)
* `Sales_Q2.csv`

Run:

```bash
python combiner.py --files Sales_Q1.xlsx Sales_Q2.csv --format both --separate-sheets
```

Output:

* `combined_20250918_153022.xlsx`

  * `All` sheet = all merged data
  * Separate sheets for each original source sheet
* `combined_20250918_153022.csv`

Both with unified, cleaned headers.

---

## 🛠 Development Notes

* Supported file types: `.xlsx`, `.csv` (extendable via `SUPPORTED_EXTS`).
* Encodings auto-tried for CSV (`utf-8`, `utf-8-sig`, `utf-16`, `latin-1`).
* Excel reading requires `openpyxl`.
* GUI defaults:

  * Output = Excel (`.xlsx`)
  * Basename = `Combined`
  * Metadata = enabled
  * Column normalisation = enabled

---

## 📜 License

---

Do you want me to also add a **"Quick Start with Screenshots" section** showing how the GUI looks (mock screenshots or diagrams), or keep it clean and text-only?
