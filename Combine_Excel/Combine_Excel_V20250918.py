import os
import sys
import traceback
from collections import defaultdict

import pandas as pd
from tkinter import Tk, messagebox, filedialog

# Optional: improve CSV inference
CSV_KWARGS = dict(sep=None, engine="python", low_memory=False)


def normalise_column_name(name: str) -> str:
    """Normalise a column name to improve matching across files."""
    if not isinstance(name, str):
        name = str(name)
    # Strip, collapse internal whitespace, flatten newlines, lower-case
    name = name.replace("\r", " ").replace("\n", " ")
    name = " ".join(name.strip().split())
    return name.lower()


def normalise_columns_inplace(df: pd.DataFrame) -> None:
    """Normalise column names in-place and deduplicate clashes."""
    seen = defaultdict(int)
    mapping = {}
    for col in df.columns:
        key = normalise_column_name(col)
        seen[key] += 1
        if seen[key] == 1:
            new_name = key
        else:
            new_name = f"{key}_{seen[key]}"  # ensure uniqueness if duplicates after normalisation
        mapping[col] = new_name
    df.rename(columns=mapping, inplace=True)


def read_csv_safely(path: str) -> pd.DataFrame:
    """Read a CSV with robust delimiter and encoding handling."""
    try:
        return pd.read_csv(path, encoding="utf-8-sig", **CSV_KWARGS)
    except Exception:
        # Fallback encodings commonly seen
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(path, encoding=enc, **CSV_KWARGS)
            except Exception:
                continue
        raise  # re-raise original


def read_excel_all_sheets(path: str) -> list:
    """Read all sheets from an Excel file, return list of DataFrames with Source_Sheet set."""
    dfs = []
    try:
        sheets = pd.read_excel(path, sheet_name=None)  # dict of sheet_name -> df
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file: {path}\n{e}")
    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue
        df = df.copy()
        df["Source_Sheet"] = sheet_name
        dfs.append(df)
    return dfs


def combine_files(paths: list) -> pd.DataFrame:
    """Combine heterogeneous CSV/Excel files with column union and provenance."""
    all_dfs = []
    column_order = []  # maintain order of first appearance across all files
    errors = []

    for path in paths:
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in (".csv",):
                df = read_csv_safely(path)
                df = df.copy()
                # Add provenance BEFORE normalisation so names are normalised consistently
                df["Source_Sheet"] = None
                df["Source_File"] = os.path.basename(path)
                normalise_columns_inplace(df)

                # Track column order
                for c in df.columns:
                    if c not in column_order:
                        column_order.append(c)

                all_dfs.append(df)

            elif ext in (".xlsx", ".xls"):
                dfs = read_excel_all_sheets(path)
                for df in dfs:
                    df = df.copy()
                    df["Source_File"] = os.path.basename(path)
                    normalise_columns_inplace(df)
                    for c in df.columns:
                        if c not in column_order:
                            column_order.append(c)
                    all_dfs.append(df)
            else:
                errors.append(f"Skipped unsupported file type: {path}")
        except Exception as e:
            tb = traceback.format_exc()
            errors.append(f"Error reading {path}: {e}\n{tb}")

    if not all_dfs:
        if errors:
            raise RuntimeError("No data was loaded.\n\n" + "\n".join(errors))
        raise RuntimeError("No data was loaded. Please select valid CSV or Excel files.")

    # Union of columns with outer concat
    combined = pd.concat(all_dfs, ignore_index=True, sort=False)

    # Reorder columns: provenance first (if present), then others in first-seen order
    ordered = []
    for prov in ("source_file", "source_sheet"):
        if prov in combined.columns:
            ordered.append(prov)
    for c in column_order:
        if c not in ordered:
            ordered.append(c)
    # Safety: any columns not tracked (shouldn't happen) appended at end
    for c in combined.columns:
        if c not in ordered:
            ordered.append(c)

    combined = combined.reindex(columns=ordered)

    # Store errors in pandas-supported attrs (avoids UserWarning)
    combined.attrs["load_errors"] = errors

    return combined


def pick_files() -> list:
    root = Tk()
    root.withdraw()
    root.update()
    filetypes = [
        ("Excel/CSV files", "*.xlsx *.xls *.csv"),
        ("Excel files", "*.xlsx *.xls"),
        ("CSV files", "*.csv"),
        ("All files", "*.*"),
    ]
    paths = filedialog.askopenfilenames(title="Select CSV/Excel files to combine", filetypes=filetypes)
    root.update()
    root.destroy()
    return list(paths)


def pick_save_path() -> str:
    root = Tk()
    root.withdraw()
    root.update()
    filetypes = [
        ("Excel Workbook", "*.xlsx"),
        ("CSV (Comma delimited)", "*.csv"),
    ]
    path = filedialog.asksaveasfilename(
        title="Save combined file as",
        defaultextension=".xlsx",
        filetypes=filetypes,
        initialfile="combined.xlsx",
    )
    root.update()
    root.destroy()
    return path


def save_output(df: pd.DataFrame, out_path: str) -> None:
    ext = os.path.splitext(out_path)[1].lower()
    if ext == ".xlsx":
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Combined")
    elif ext == ".csv":
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
    else:
        # Default to Excel if unknown extension
        out_path = out_path + ".xlsx"
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Combined")


def main():
    try:
        paths = pick_files()
        if not paths:
            messagebox.showinfo("Cancelled", "No files selected.")
            return

        combined = combine_files(paths)

        out_path = pick_save_path()
        if not out_path:
            messagebox.showinfo("Cancelled", "Save location not selected.")
            return

        save_output(combined, out_path)

        msg = f"Successfully combined {len(paths)} file(s).\nRows: {len(combined):,}\nSaved to:\n{out_path}"
        load_errors = combined.attrs.get("load_errors")
        if load_errors:
            msg += "\n\nSome files/sheets had issues:\n- " + "\n- ".join(load_errors)

        messagebox.showinfo("Done", msg)

    except Exception as e:
        messagebox.showerror("Error", f"{e}\n\n{traceback.format_exc()}")


if __name__ == "__main__":
    # Ensure needed packages are present (non-fatal hint)
    try:
        import pandas  # noqa: F401
    except ImportError:
        messagebox.showerror(
            "Missing dependency",
            "This script requires the 'pandas' package.\nInstall it with:\n\npip install pandas openpyxl"
        )
        sys.exit(1)

    main()
