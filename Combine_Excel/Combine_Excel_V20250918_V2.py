import os
import sys
import traceback
from collections import defaultdict

import pandas as pd
from tkinter import (
    Tk, Toplevel, Label, Button, Listbox, Scrollbar, SINGLE, MULTIPLE, END,
    VERTICAL, RIGHT, Y, LEFT, BOTH, X, BOTTOM, messagebox, filedialog
)

# CSV reading defaults: auto-detect delimiter, robust parsing
CSV_KWARGS = dict(sep=None, engine="python", low_memory=False)

# ---------- Column normalisation utilities ----------

def normalise_column_name(name: str) -> str:
    """Normalise a column name to improve matching across files."""
    if not isinstance(name, str):
        name = str(name)
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
        new_name = key if seen[key] == 1 else f"{key}_{seen[key]}"
        mapping[col] = new_name
    df.rename(columns=mapping, inplace=True)

def flatten_multiindex_columns(df: pd.DataFrame) -> None:
    """Flatten a MultiIndex header into a single row of strings."""
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join([str(x) for x in tup if str(x) != "None"]).strip() for tup in df.columns.tolist()]

# ---------- CSV/Excel readers ----------

def read_csv_safely(path: str, nrows=None, header='infer'):
    """Read a CSV with robust delimiter and encoding handling."""
    try:
        return pd.read_csv(path, encoding="utf-8-sig", nrows=nrows, header=header, **CSV_KWARGS)
    except Exception:
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(path, encoding=enc, nrows=nrows, header=header, **CSV_KWARGS)
            except Exception:
                continue
        raise

def list_excel_sheets(path: str):
    """Return list of sheet names in an Excel file."""
    try:
        xls = pd.ExcelFile(path)
        return xls.sheet_names
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file: {path}\n{e}")

def read_excel_preview(path: str, sheet_name: str, nrows: int = 10) -> pd.DataFrame:
    """Read top nrows of a sheet without headers for preview/header detection."""
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=nrows)
        return df
    except Exception as e:
        raise RuntimeError(f"Failed to preview Excel file: {path} (sheet: {sheet_name})\n{e}")

# ---------- Header detection ----------

def _is_string_like(val) -> bool:
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return False
        # Treat as string-like if it contains any alphabetic character
        return any(ch.isalpha() for ch in s)
    return False

def _normalise_cell_for_unique(val) -> str:
    try:
        s = str(val).strip().lower()
    except Exception:
        s = ""
    # collapse whitespace
    s = " ".join(s.split())
    return s

def detect_header_row_from_preview(preview_df: pd.DataFrame):
    """
    Return (best_index, confidence, per_row_scores).
    Confidence is a float 0..1. We'll consider 'confident' if >= 0.85 or a strong lead over second best.
    """
    if preview_df is None or preview_df.empty:
        return 0, 0.0, []

    scores = []
    ncols = preview_df.shape[1]
    for ridx in range(preview_df.shape[0]):
        row = preview_df.iloc[ridx].tolist()
        # Non-empty
        non_empty_vals = [c for c in row if pd.notna(c) and str(c).strip() != ""]
        non_empty_ratio = len(non_empty_vals) / max(ncols, 1)

        # String-like
        str_like_vals = [c for c in non_empty_vals if _is_string_like(c)]
        string_like_ratio = len(str_like_vals) / max(len(non_empty_vals), 1)

        # Uniqueness among string-like cells
        norm = [_normalise_cell_for_unique(c) for c in str_like_vals]
        unique_ratio = len(set([v for v in norm if v != ""])) / max(len([v for v in norm if v != ""]), 1)

        # Penalise rows where many cells look like pure numbers/dates
        numericish_penalty = 0.0
        numericish_count = 0
        for c in non_empty_vals:
            if isinstance(c, (int, float, pd.Timestamp)):
                numericish_count += 1
            else:
                s = str(c).strip()
                if s and all(ch.isdigit() or ch in " .,-/:+" for ch in s) and not any(ch.isalpha() for ch in s):
                    numericish_count += 1
        if non_empty_vals:
            numericish_penalty = min(0.3, numericish_count / len(non_empty_vals) * 0.5)

        score = (0.4 * non_empty_ratio + 0.4 * string_like_ratio + 0.2 * unique_ratio) - numericish_penalty
        # Clamp to [0,1]
        score = max(0.0, min(1.0, score))
        scores.append(score)

    best_idx = max(range(len(scores)), key=lambda i: scores[i])
    best_score = scores[best_idx]
    # Second best
    s2 = sorted(scores, reverse=True)[1] if len(scores) > 1 else 0.0

    # Confidence heuristic
    lead = best_score - s2
    confidence = max(0.0, min(1.0, best_score))
    # Boost confidence if clear lead
    if lead >= 0.25 and best_score >= 0.7:
        confidence = max(confidence, 0.9)
    return best_idx, confidence, scores

# ---------- UI dialogs (Tkinter) ----------

def _trunc(s: str, max_len: int) -> str:
    return s if len(s) <= max_len else s[: max_len - 1] + "…"

def select_sheets_dialog(root: Tk, file_path: str, sheet_names: list[str]) -> list[str] | None:
    """Return list of selected sheets (or None if cancelled)."""
    top = Toplevel(root)
    top.title(f"Select sheets - {os.path.basename(file_path)}")
    top.grab_set()
    top.geometry("520x420")

    Label(top, text=f"Select sheet(s) to include from:\n{file_path}", justify="left").pack(anchor="w", padx=10, pady=(10, 6))

    lb = Listbox(top, selectmode=MULTIPLE)
    sb = Scrollbar(top, orient=VERTICAL, command=lb.yview)
    lb.config(yscrollcommand=sb.set)
    for name in sheet_names:
        lb.insert(END, name)
    lb.pack(side=LEFT, fill=BOTH, expand=True, padx=(10,0), pady=(0,10))
    sb.pack(side=RIGHT, fill=Y, padx=(0,10), pady=(0,10))

    btn_frame = Toplevel(top)  # Use a separate Toplevel for proper layout? We'll keep simple: use Buttons below.
    # Simpler: pack buttons in the same top using a frame-like approach:
    def select_all():
        lb.select_set(0, END)

    def clear_all():
        lb.selection_clear(0, END)

    selected = {"value": None}

    def ok():
        idxs = lb.curselection()
        if not idxs:
            messagebox.showwarning("No selection", "Please select at least one sheet.", parent=top)
            return
        selected["value"] = [sheet_names[i] for i in idxs]
        top.destroy()

    def cancel():
        selected["value"] = None
        top.destroy()

    # Dummy frame using another window is clunky; instead, create a lightweight container area:
    Button(top, text="Select all", command=select_all).pack(fill=X, padx=10, pady=(0,4))
    Button(top, text="Clear", command=clear_all).pack(fill=X, padx=10, pady=(0,10))

    Button(top, text="OK", command=ok).pack(side=LEFT, padx=(10,5), pady=(0,12))
    Button(top, text="Cancel", command=cancel).pack(side=LEFT, padx=(0,10), pady=(0,12))

    top.protocol("WM_DELETE_WINDOW", cancel)
    root.wait_window(top)
    return selected["value"]

def select_header_row_dialog(root: Tk, title: str, preview_df: pd.DataFrame, suggested_index: int) -> int | None:
    """
    Show a list of top rows and let the user choose which row is the header.
    Returns the selected row index (0-based), or None if cancelled.
    """
    top = Toplevel(root)
    top.title(title)
    top.grab_set()
    top.geometry("840x420")

    Label(
        top,
        text="Select the row that contains the column headers.\n"
             "Tip: Pick the row where values look like field names, not data.",
        justify="left"
    ).pack(anchor="w", padx=10, pady=(10, 6))

    lb = Listbox(top, selectmode=SINGLE, height=16)
    sb = Scrollbar(top, orient=VERTICAL, command=lb.yview)
    lb.config(yscrollcommand=sb.set)

    # Build preview lines
    max_cell_len = 24
    max_row_len = 160
    for ridx in range(preview_df.shape[0]):
        row = preview_df.iloc[ridx].tolist()
        cells = []
        for v in row:
            s = "" if pd.isna(v) else str(v)
            s = _trunc(s.replace("\n", " ").strip(), max_cell_len)
            cells.append(s)
        line = f"Row {ridx}: " + " | ".join(cells)
        line = _trunc(line, max_row_len)
        lb.insert(END, line)

    lb.pack(side=LEFT, fill=BOTH, expand=True, padx=(10,0), pady=(0,10))
    sb.pack(side=RIGHT, fill=Y, padx=(0,10), pady=(0,10))

    # Preselect suggestion
    try:
        lb.selection_set(suggested_index)
        lb.see(suggested_index)
    except Exception:
        pass

    selected = {"value": None}

    def ok():
        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a header row.", parent=top)
            return
        selected["value"] = sel[0]
        top.destroy()

    def cancel():
        selected["value"] = None
        top.destroy()

    Button(top, text="OK", command=ok).pack(side=LEFT, padx=(10,5), pady=(0,12))
    Button(top, text="Cancel", command=cancel).pack(side=LEFT, padx=(0,10), pady=(0,12))

    top.protocol("WM_DELETE_WINDOW", cancel)
    root.wait_window(top)
    return selected["value"]

# ---------- Main combination logic with interactive sheet/header selection ----------

def combine_files_with_interaction(root: Tk, paths: list) -> pd.DataFrame:
    """
    Combine CSV/Excel files. For Excel, ask which sheet(s) to include first.
    For each file/sheet (and CSV), intelligently suggest a header row; if not confident, ask the user to pick.
    """
    all_dfs = []
    column_order = []  # first-seen order across all files
    errors = []

    for path in paths:
        ext = os.path.splitext(path)[1].lower()

        # Handle Excel: sheet selection first
        if ext in (".xlsx", ".xls"):
            try:
                sheet_names = list_excel_sheets(path)
            except Exception as e:
                tb = traceback.format_exc()
                errors.append(f"Error listing sheets {path}: {e}\n{tb}")
                continue

            selected_sheets = sheet_names
            if len(sheet_names) > 1:
                chosen = select_sheets_dialog(root, path, sheet_names)
                if chosen is None:
                    # User cancelled this file; skip it
                    continue
                selected_sheets = chosen
                if not selected_sheets:
                    # None selected, skip file
                    continue

            for sheet in selected_sheets:
                try:
                    # Preview and detect header
                    preview = read_excel_preview(path, sheet, nrows=10)
                    hdr_idx, confidence, _scores = detect_header_row_from_preview(preview)

                    if confidence < 0.85:
                        title = f"Select header row - {os.path.basename(path)} [{sheet}]"
                        user_choice = select_header_row_dialog(root, title, preview, hdr_idx)
                        if user_choice is not None:
                            hdr_idx = user_choice

                    # Read sheet with chosen header
                    df = pd.read_excel(path, sheet_name=sheet, header=hdr_idx)
                    flatten_multiindex_columns(df)

                    # Provenance + normalise
                    df = df.copy()
                    df["Source_File"] = os.path.basename(path)
                    df["Source_Sheet"] = sheet
                    normalise_columns_inplace(df)

                    # Track column order
                    for c in df.columns:
                        if c not in column_order:
                            column_order.append(c)

                    all_dfs.append(df)

                except Exception as e:
                    tb = traceback.format_exc()
                    errors.append(f"Error reading {path} sheet '{sheet}': {e}\n{tb}")

        elif ext == ".csv":
            try:
                # Preview CSV without header to detect header row
                preview = read_csv_safely(path, nrows=10, header=None)
                hdr_idx, confidence, _scores = detect_header_row_from_preview(preview)

                if confidence < 0.85:
                    title = f"Select header row - {os.path.basename(path)}"
                    user_choice = select_header_row_dialog(root, title, preview, hdr_idx)
                    if user_choice is not None:
                        hdr_idx = user_choice

                # Read full CSV with chosen header
                df = read_csv_safely(path, header=hdr_idx)
                df = df.copy()
                df["Source_File"] = os.path.basename(path)
                df["Source_Sheet"] = None
                normalise_columns_inplace(df)

                # Track column order
                for c in df.columns:
                    if c not in column_order:
                        column_order.append(c)

                all_dfs.append(df)

            except Exception as e:
                tb = traceback.format_exc()
                errors.append(f"Error reading CSV {path}: {e}\n{tb}")

        else:
            errors.append(f"Skipped unsupported file type: {path}")

    if not all_dfs:
        msg = "No data was loaded."
        if errors:
            msg += "\n\n" + "\n".join(errors)
        raise RuntimeError(msg)

    combined = pd.concat(all_dfs, ignore_index=True, sort=False)

    # Order columns: provenance first if present (normalised to lower-case), then the rest
    ordered = []
    for prov in ("source_file", "source_sheet"):
        if prov in combined.columns:
            ordered.append(prov)
    for c in column_order:
        if c not in ordered:
            ordered.append(c)
    for c in combined.columns:
        if c not in ordered:
            ordered.append(c)
    combined = combined.reindex(columns=ordered)

    # Attach errors using DataFrame attrs to avoid pandas UserWarning
    combined.attrs["load_errors"] = errors
    return combined

# ---------- Save / pick helpers ----------

def pick_files(root: Tk) -> list:
    filetypes = [
        ("Excel/CSV files", "*.xlsx *.xls *.csv"),
        ("Excel files", "*.xlsx *.xls"),
        ("CSV files", "*.csv"),
        ("All files", "*.*"),
    ]
    paths = filedialog.askopenfilenames(parent=root, title="Select CSV/Excel files to combine", filetypes=filetypes)
    return list(paths)

def pick_save_path(root: Tk) -> str:
    filetypes = [
        ("Excel Workbook", "*.xlsx"),
        ("CSV (Comma delimited)", "*.csv"),
    ]
    path = filedialog.asksaveasfilename(
        parent=root,
        title="Save combined file as",
        defaultextension=".xlsx",
        filetypes=filetypes,
        initialfile="combined.xlsx",
    )
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
        out_path2 = out_path + ".xlsx"
        with pd.ExcelWriter(out_path2, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Combined")

# ---------- Main ----------

def main():
    try:
        root = Tk()
        root.withdraw()

        paths = pick_files(root)
        if not paths:
            messagebox.showinfo("Cancelled", "No files selected.", parent=root)
            root.destroy()
            return

        combined = combine_files_with_interaction(root, paths)

        out_path = pick_save_path(root)
        if not out_path:
            messagebox.showinfo("Cancelled", "Save location not selected.", parent=root)
            root.destroy()
            return

        save_output(combined, out_path)

        msg = f"Successfully combined {len(paths)} file(s).\nRows: {len(combined):,}\nSaved to:\n{out_path}"
        load_errors = combined.attrs.get("load_errors")
        if load_errors:
            msg += "\n\nSome files/sheets had issues:\n- " + "\n- ".join(load_errors)

        messagebox.showinfo("Done", msg, parent=root)

        root.destroy()

    except Exception as e:
        try:
            messagebox.showerror("Error", f"{e}\n\n{traceback.format_exc()}")
        except Exception:
            print(f"Error: {e}\n\n{traceback.format_exc()}", file=sys.stderr)
        finally:
            try:
                root.destroy()
            except Exception:
                pass
        sys.exit(1)

if __name__ == "__main__":
    # Ensure needed packages are present (non-fatal hint)
    try:
        import pandas  # noqa: F401
    except ImportError:
        # Create a minimal root to show the messagebox if possible
        try:
            _r = Tk(); _r.withdraw()
            messagebox.showerror(
                "Missing dependency",
                "This script requires the 'pandas' package.\nInstall it with:\n\npip install pandas openpyxl",
                parent=_r
            )
            _r.destroy()
        except Exception:
            print("This script requires the 'pandas' package.\nInstall it with:\n\npip install pandas openpyxl", file=sys.stderr)
        sys.exit(1)

    main()
