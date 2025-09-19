# Universal Excel Combiner

A Python tool to **combine Excel (`.xlsx`) and CSV files** into a single clean dataset with **automatic header detection**, **schema alignment**, and flexible **output options**.

Supports both **command-line interface (CLI)** and an optional **GUI (Tkinter)**.

---

## ✨ Features

* 📂 **Supports Excel (`.xlsx`) and CSV** (auto-detects multiple sheets).
* 🧠 **Smart header detection**:
  * Finds the most likely header row.
  * Handles multi-row headers (merges if needed).
  * De-duplicates duplicate column names.
* 🔗 **Schema alignment** across files:
  * Aligns weak/no-header files to the most common header layout (optional).
* 🛠 **Options for column name normalisation** (trimming, whitespace collapse).
* 🏷 **Metadata tracking**:
  * Adds `source_file` and `source_sheet` columns to trace data origin.
* 📊 **Output formats**:
  * Excel (`.xlsx`) with either one combined sheet or **separate per-source sheets + All**.
  * CSV (`.csv`).
  * Both simultaneously.
* 🎛 **Two usage modes**:
  * **CLI** for scripting and automation.
  * **GUI** (if `tkinter` is available).

---

## 🚀 Setup / Installation

Make sure you have:

* Python 3.8 or newer.
* Required packages:
  * `pandas`
  * `openpyxl` (for Excel support)
* (Optional) `tkinter` for the GUI

Install dependencies, for example with pip:

```bash
pip install pandas openpyxl
```

---

## 🖥 Usage

### CLI Mode

Run with:

```bash
python combiner.py --files data1.xlsx data2.csv --outdir ./output --basename merged
```

#### CLI Options

| Option                     | Description                                             |
| -------------------------- | ------------------------------------------------------- |
| `--files FILES...`         | Input files (`.xlsx`, `.csv`)                           |
| `--outdir DIR`             | Output directory (default: `.`)                         |
| `--basename NAME`          | Base name for output files (default: `combined`)        |
| `--no-metadata`            | Do not add `source_file` / `source_sheet` columns       |
| `--no-normalise`           | Do not normalise column names                           |
| `--align-headerless`       | Try aligning weak/no-header files to most common schema |
| `--format {xlsx,csv,both}` | Output format(s), default: `xlsx`                       |
| `--separate-sheets`        | For Excel, create per-sheet outputs + All               |

#### Example 1 – Simple combine

```bash
python combiner.py --files data1.xlsx data2.xlsx
```

Produces a single `combined_YYYYMMDD_HHMMSS.xlsx`.

#### Example 2 – Separate sheets + CSV output

```bash
python combiner.py --files data1.xlsx data2.csv --separate-sheets --format both
```

---

### GUI Mode

Run without arguments:

```bash
python combiner.py
```

If `tkinter` is available, a simple GUI opens that walks you through:

1. Selecting input files (`.xlsx`, `.csv`).
2. Choosing output file name and location using "Save As..." (sets both directory and base file name).
3. Setting options (normalisation, metadata, output format, etc.).
4. Running the combine action for output.

**Note:**  
The chosen "Save As..." name is used as the base for all outputs. The tool appends a timestamp and the correct extension (`.xlsx` and/or `.csv`) automatically.

---

## 📂 Output

Generated files are timestamped for uniqueness, e.g.:

```
output/
  merged_20250918_153245.xlsx
  merged_20250918_153245.csv
```

* Excel sheets will contain:
  * `All` → combined dataset.
  * One sheet per input file/sheet (if "separate sheets" is enabled).
* CSV is encoded as UTF-8 with BOM (`utf-8-sig`) for better compatibility.

---

## 🔧 Development Notes

* Header detection uses heuristics:
  * Prefers rows with more “header-like” tokens (non-numeric, unique).
  * Can merge two consecutive rows into one header if they both look headerish.
  * Falls back to placeholder headers if nothing is convincing.
* Alignment across files can be enabled with `--align-headerless` to map weak/no-header files to the most common schema (by column count).
* GUI uses `tkinter` and `ttk` for cross-platform basic UI.
* In the GUI, output location and base file name are set in a single "Save As..." step.

---

## 📝 License

MIT License — free to use, modify, and distribute.
