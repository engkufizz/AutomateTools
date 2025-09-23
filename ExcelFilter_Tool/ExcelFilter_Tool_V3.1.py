import os
import sys
import re
import math
import datetime
from typing import List, Tuple, Optional, Dict, Any

import pandas as pd

# Tkinter GUI
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except Exception as e:
    print("Tkinter is required to run this GUI. Please ensure tkinter is installed.", file=sys.stderr)
    sys.exit(2)

# ===================== Utilities and header detection =====================

SUPPORTED_EXTS = {".xlsx", ".csv"}

def timestamp() -> str:
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def _safe_str(v: object) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()

def _is_date_like(s: str) -> bool:
    s = s.strip()
    if not s:
        return False
    return bool(re.match(r"^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$", s))

def _is_numeric_like(s: str) -> bool:
    s = s.strip()
    if not s:
        return False
    if s.endswith("%"):
        s = s[:-1]
    return bool(re.match(r"^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$", s))

def _token_is_headerish(v: object) -> bool:
    s = _safe_str(v)
    if not s:
        return False
    if len(s) > 120:
        return False
    if _is_numeric_like(s) or _is_date_like(s):
        return False
    return bool(re.search(r"[A-Za-z_]", s))

def _row_score(cells: list) -> Tuple[float, float, float, int, int]:
    tokens = [_safe_str(c) for c in cells]
    nonempty = [c for c in tokens if c != ""]
    total = len(tokens)
    if not nonempty:
        return -1.0, 0.0, 0.0, 0, total
    headerish = sum(1 for c in nonempty if _token_is_headerish(c))
    headerish_ratio = headerish / max(1, len(nonempty))
    unique_ratio = len(set(map(lambda x: x.lower(), nonempty))) / max(1, len(nonempty))
    empty_penalty = (total - len(nonempty)) / max(1, total)
    score = headerish + 0.8 * unique_ratio - 0.5 * empty_penalty
    return score, headerish_ratio, unique_ratio, len(nonempty), total

def _dedupe_headers(cols: list) -> list:
    seen: Dict[str, int] = {}
    out: List[str] = []
    for c in cols:
        name = _safe_str(c) or "column"
        if name in seen:
            seen[name] += 1
            out.append(f"{name}__{seen[name]}")
        else:
            seen[name] = 0
            out.append(name)
    return out

def _maybe_merge_two_row_header(df: pd.DataFrame, h: int) -> Tuple[List[str], int]:
    row1 = list(df.iloc[h].tolist())
    if h + 1 >= len(df):
        return _dedupe_headers([_safe_str(x) for x in row1]), 1
    row2 = list(df.iloc[h + 1].tolist())
    r1_tokens = [_safe_str(x) for x in row1]
    r2_tokens = [_safe_str(x) for x in row2]

    r1_nonempty = sum(1 for x in r1_tokens if x != "")
    r2_nonempty = sum(1 for x in r2_tokens if x != "")
    r1_headerish = sum(1 for x in r1_tokens if _token_is_headerish(x))
    r2_headerish = sum(1 for x in r2_tokens if _token_is_headerish(x))
    r1_dupes = len(r1_tokens) - len(set(map(lambda x: x.lower(), r1_tokens)))

    should_merge = (
        r2_headerish >= max(2, int(0.5 * max(r1_nonempty, 1))) and
        (r1_nonempty < len(r1_tokens) or r1_dupes > 0)
    )
    if not should_merge:
        return _dedupe_headers([_safe_str(x) for x in row1]), 1

    merged: List[str] = []
    for a, b in zip(r1_tokens, r2_tokens):
        sa = a
        sb = b
        if sa and sb:
            merged.append(f"{sa} {sb}")
        elif sa:
            merged.append(sa)
        else:
            merged.append(sb)
    return _dedupe_headers(merged), 2

def analyse_header(df_raw: pd.DataFrame, max_scan_rows: int = 50) -> Dict[str, Any]:
    meta: Dict[str, Any] = {
        "best_idx": 0,
        "best_score": float("-inf"),
        "best_headerish_ratio": 0.0,
        "best_unique_ratio": 0.0,
        "rows_used": 0,
        "headers": [],
        "classification": "none",
        "data_start": 0,
        "ncols": int(df_raw.shape[1]) if df_raw is not None else 0,
    }
    if df_raw is None or df_raw.empty:
        return meta
    scan_upto = min(max_scan_rows, len(df_raw))
    best_idx, best_score, best_hdr_ratio, best_uniq = 0, float("-inf"), 0.0, 0.0
    for i in range(scan_upto):
        score, hdr_ratio, uniq_ratio, _nonempty, _total = _row_score(list(df_raw.iloc[i].tolist()))
        if score > best_score:
            best_score = score
            best_idx = i
            best_hdr_ratio = hdr_ratio
            best_uniq = uniq_ratio

    headers, rows_used = _maybe_merge_two_row_header(df_raw, best_idx)
    data_start = best_idx + rows_used

    data_rows = df_raw.iloc[data_start : min(len(df_raw), data_start + 5)]
    if len(data_rows) == 0:
        data_like_ratio = 0.0
    else:
        tokens = [_safe_str(v) for v in data_rows.values.flatten().tolist()]
        nonempty = [t for t in tokens if t != ""]
        if not nonempty:
            data_like_ratio = 0.0
        else:
            dataish = sum(1 for t in nonempty if (_is_numeric_like(t) or _is_date_like(t) or not _token_is_headerish(t)))
            data_like_ratio = dataish / max(1, len(nonempty))

    if best_score >= 0.8 and best_hdr_ratio >= 0.5 and data_like_ratio >= 0.5:
        classification = "strong"
    elif best_score >= 0.3 and best_hdr_ratio >= 0.3:
        classification = "weak"
    else:
        classification = "none"

    meta.update(
        best_idx=best_idx,
        best_score=best_score,
        best_headerish_ratio=best_hdr_ratio,
        best_unique_ratio=best_uniq,
        rows_used=rows_used,
        headers=_dedupe_headers(headers),
        classification=classification,
        data_start=data_start,
        ncols=int(df_raw.shape[1]),
    )
    return meta

def build_dataframe(df_raw: pd.DataFrame, headers: List[str], data_start: int) -> pd.DataFrame:
    df = df_raw.iloc[data_start:].copy()
    if len(df.columns) > len(headers):
        headers = headers + [f"column_{i+1}" for i in range(len(df.columns) - len(headers))]
    elif len(df.columns) < len(headers):
        headers = headers[:len(df.columns)]
    df.columns = headers
    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(axis=1, how="all")
    df = df.reset_index(drop=True).convert_dtypes()
    return df

def build_placeholder_headers(ncols: int) -> List[str]:
    return [f"column_{i+1}" for i in range(ncols)]

def apply_column_normalisation(df: pd.DataFrame, do_normalise: bool) -> pd.DataFrame:
    if not do_normalise:
        return df
    new_cols = []
    counts: Dict[str, int] = {}
    for c in df.columns:
        s = str(c).strip()
        s = re.sub(r"\s+", " ", s)
        if s in counts:
            counts[s] += 1
            new_cols.append(f"{s}__{counts[s]}")
        else:
            counts[s] = 0
            new_cols.append(s)
    df = df.copy()
    df.columns = new_cols
    return df

# ===================== IO helpers =====================

def read_csv_raw(path: str) -> pd.DataFrame:
    encodings = [None, "utf-8", "utf-8-sig", "utf-16", "latin-1"]
    last_err = None
    for enc in encodings:
        try:
            return pd.read_csv(path, sep=None, engine="python", encoding=enc, header=None, dtype=str)
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Failed to read CSV: {path} ({last_err})")

def read_excel_raw_all_sheets(path: str) -> List[Tuple[str, pd.DataFrame]]:
    try:
        sheets_raw: Dict[str, pd.DataFrame] = pd.read_excel(path, sheet_name=None, header=None, dtype=str)
    except ImportError as e:
        raise RuntimeError("openpyxl is required for .xlsx. Install with: pip install openpyxl") from e
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel: {path} ({e})")
    items: List[Tuple[str, pd.DataFrame]] = []
    for sheet, df_raw in sheets_raw.items():
        items.append((str(sheet), df_raw if df_raw is not None else pd.DataFrame()))
    return items

def load_main_source(path: str, all_sheets: bool = False) -> List[Tuple[str, pd.DataFrame, Dict[str, Any]]]:
    ext = os.path.splitext(path)[1].lower()
    items: List[Tuple[str, pd.DataFrame, Dict[str, Any]]] = []
    if ext == ".csv":
        df_raw = read_csv_raw(path)
        meta = analyse_header(df_raw)
        items.append(("(CSV)", df_raw, meta))
    elif ext == ".xlsx":
        sheets = read_excel_raw_all_sheets(path)
        if not sheets:
            raise RuntimeError("Workbook has no sheets or cannot be read.")
        if all_sheets:
            for name, df_raw in sheets:
                meta = analyse_header(df_raw)
                items.append((name, df_raw, meta))
        else:
            # choose best sheet using classification, score, and data length
            best = None
            best_meta = None
            for name, df_raw in sheets:
                meta = analyse_header(df_raw)
                strength_rank = {"strong": 2, "weak": 1, "none": 0}.get(meta["classification"], 0)
                data_len = max(0, len(df_raw) - int(meta["data_start"]))
                key = (strength_rank, float(meta["best_score"]), data_len)
                if best is None or key > best_meta[0]:
                    best = (name, df_raw)
                    best_meta = (key, meta)
            if best is None:
                raise RuntimeError("No readable sheets found.")
            items.append((best[0], best[1], best_meta[1]))
    else:
        raise RuntimeError(f"Unsupported file type: {ext}")
    return items

def build_df_from_unit(df_raw: pd.DataFrame, meta: Dict[str, Any]) -> pd.DataFrame:
    if meta["classification"] in ("strong", "weak"):
        headers = list(meta["headers"])
        start = int(meta["data_start"])
    else:
        headers = build_placeholder_headers(meta["ncols"])
        start = 0
    return build_dataframe(df_raw, headers, start)

def read_list_values(path: str) -> List[str]:
    """
    Read values from the first non-empty column (robust to commas and headers).
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"List file not found: {path}")
    ext = os.path.splitext(path)[1].lower()

    def extract_first_nonempty_col(df_raw: pd.DataFrame) -> List[str]:
        if df_raw is None or df_raw.shape[1] == 0:
            return []
        best_idx = 0
        best_count = -1
        for i in range(df_raw.shape[1]):
            col_vals = df_raw.iloc[:, i].tolist()
            cnt = sum(1 for v in col_vals if _safe_str(v) != "")
            if cnt > best_count:
                best_idx, best_count = i, cnt
        if best_count <= 0:
            return []
        vals = [_safe_str(v) for v in df_raw.iloc[:, best_idx].tolist()]
        vals = [v for v in vals if v != ""]
        return vals

    if ext == ".csv":
        df_raw = read_csv_raw(path)
        return extract_first_nonempty_col(df_raw)
    elif ext == ".xlsx":
        sheets = read_excel_raw_all_sheets(path)
        # choose sheet with most non-empty tokens
        best = None
        best_score = -1
        for name, df_raw in sheets:
            tokens = [_safe_str(v) for v in df_raw.values.flatten().tolist()]
            nonempty = sum(1 for t in tokens if t != "")
            if nonempty > best_score:
                best = df_raw
                best_score = nonempty
        if best is None:
            return []
        return extract_first_nonempty_col(best)
    else:
        raise RuntimeError(f"Unsupported list file type: {ext}")

# ===================== Filtering primitives =====================

def series_as_str(s: pd.Series, case_sensitive: bool, trim: bool = True) -> pd.Series:
    z = s.astype("string")
    if trim:
        z = z.str.strip()
    if not case_sensitive:
        z = z.str.lower()
    return z

def apply_rule(df: pd.DataFrame, rule: Dict[str, Any]) -> pd.Series:
    """
    Supported operators:
      String:
        - is (exact equals to any), is_not (exact not in any)
        - contains_any, not_contains_any
        - regex_any, not_regex
        - in_list_file, not_in_list_file
        - is_empty, not_empty
      Numeric:
        - gt, gte, lt, lte, between
    """
    col = rule["column"]
    if col not in df.columns:
        raise KeyError(f"Column '{col}' not in data.")
    op_raw = rule["operator"].lower().strip()
    # synonyms support (accepts legacy names too)
    synonyms = {"equals_any": "is", "not_equals_any": "is_not"}
    op = synonyms.get(op_raw, op_raw)
    case_sensitive = bool(rule.get("case_sensitive", False))

    # Empty checks
    if op in ("is_empty", "not_empty"):
        s = df[col]
        # NA values are empty
        mask = s.isna()
        # Also treat trimmed empty strings as empty
        try:
            s_str = series_as_str(s, case_sensitive=True)
            mask = mask | s_str.eq("")
        except Exception:
            pass
        if op == "not_empty":
            mask = ~mask
        return mask.fillna(False)

    # String and list-file ops
    if op in ("is", "is_not", "contains_any", "not_contains_any",
              "regex_any", "not_regex", "in_list_file", "not_in_list_file"):
        if op in ("in_list_file", "not_in_list_file"):
            list_path = rule.get("file")
            if not list_path:
                raise ValueError(f"operator={op} requires 'file'")
            values = read_list_values(list_path)
            if not values:
                return pd.Series(False, index=df.index)
        else:
            values = rule.get("values")
            if values is None or not isinstance(values, list):
                raise ValueError(f"operator={op} requires 'values' list")
            if len(values) == 0:
                return pd.Series(False, index=df.index)

        col_s = series_as_str(df[col], case_sensitive=case_sensitive)

        if op in ("is", "is_not", "in_list_file", "not_in_list_file"):
            vals = [str(v).strip() if v is not None else "" for v in values]
            if not case_sensitive:
                vals = [v.lower() for v in vals]
            mask = col_s.isin(vals)
            if op in ("is_not", "not_in_list_file"):
                mask = ~mask
            return mask.fillna(False)

        if op in ("contains_any", "not_contains_any"):
            toks = [str(v) for v in values if str(v) != ""]
            if not toks:
                return pd.Series(False, index=df.index)
            pattern = "|".join(re.escape(t if case_sensitive else t.lower()) for t in toks)
            mask = col_s.str.contains(pattern, regex=True, na=False)
            if op == "not_contains_any":
                mask = ~mask
            return mask

        if op in ("regex_any", "not_regex"):
            patterns = [str(v) for v in values if str(v) != ""]
            if not patterns:
                return pd.Series(False, index=df.index)
            joined = "|".join(f"(?:{p})" for p in patterns)
            flags = 0 if case_sensitive else re.IGNORECASE
            mask = df[col].astype("string").str.contains(joined, regex=True, flags=flags, na=False)
            if op == "not_regex":
                mask = ~mask
            return mask

    # Numeric ops
    if op in ("gt", "gte", "lt", "lte"):
        val = rule.get("value", None)
        if val is None:
            raise ValueError(f"operator={op} requires 'value'")
        s = pd.to_numeric(df[col], errors="coerce")
        if op == "gt":
            mask = s > float(val)
        elif op == "gte":
            mask = s >= float(val)
        elif op == "lt":
            mask = s < float(val)
        else:
            mask = s <= float(val)
        return mask.fillna(False)

    if op == "between":
        vmin = rule.get("value_min", None)
        vmax = rule.get("value_max", None)
        if vmin is None or vmax is None:
            raise ValueError("operator=between requires value_min and value_max")
        s = pd.to_numeric(df[col], errors="coerce")
        mask = (s >= float(vmin)) & (s <= float(vmax))
        return mask.fillna(False)

    raise ValueError(f"Unsupported operator: {op_raw}")

def combine_masks(masks: List[pd.Series], mode: str) -> pd.Series:
    if not masks:
        return pd.Series(True)  # no rules => keep all
    mode = mode.upper()
    out = masks[0].copy()
    for m in masks[1:]:
        if mode == "AND":
            out = out & m
        elif mode == "OR":
            out = out | m
        else:
            raise ValueError("combine mode must be AND or OR")
    return out

# ===================== Output writing =====================

def write_outputs(df_by_sheet: List[Tuple[str, pd.DataFrame]],
                  out_dir: str,
                  base_name: str,
                  write_excel: bool,
                  write_csv: bool,
                  separate_sheets: bool) -> List[str]:
    written: List[str] = []
    safe_base = re.sub(r"[^\w\-.]+", "_", base_name).strip("_") or "filtered"
    ts = timestamp()

    if write_excel:
        xlsx_path = os.path.join(out_dir, f"{safe_base}_{ts}.xlsx")
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            if separate_sheets and len(df_by_sheet) > 1:
                df_all = pd.concat([df.assign(source_sheet=nm) for nm, df in df_by_sheet], ignore_index=True)
                df_all.to_excel(writer, sheet_name="All", index=False)
                used = {"All"}
                for name, df in df_by_sheet:
                    sheet_name = str(name).strip() or "Sheet"
                    sheet_name = sheet_name[:31]
                    base = sheet_name
                    i = 1
                    while sheet_name in used:
                        suffix = f"_{i}"
                        sheet_name = (base[: max(0, 31 - len(suffix))]) + suffix
                        i += 1
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    used.add(sheet_name)
            else:
                if len(df_by_sheet) == 1:
                    df_by_sheet[0][1].to_excel(writer, sheet_name="Filtered", index=False)
                else:
                    df_all = pd.concat([df.assign(source_sheet=nm) for nm, df in df_by_sheet], ignore_index=True)
                    df_all.to_excel(writer, sheet_name="Filtered", index=False)
        written.append(xlsx_path)

    if write_csv:
        csv_path = os.path.join(out_dir, f"{safe_base}_{ts}.csv")
        if len(df_by_sheet) == 1:
            df_by_sheet[0][1].to_csv(csv_path, index=False, encoding="utf-8-sig")
        else:
            df_all = pd.concat([df.assign(source_sheet=nm) for nm, df in df_by_sheet], ignore_index=True)
            df_all.to_csv(csv_path, index=False, encoding="utf-8-sig")
        written.append(csv_path)

    return written

# ===================== GUI App =====================

class FilterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal Excel/CSV Filter")

        # State
        self.main_path: Optional[str] = None
        self.units: List[Tuple[str, pd.DataFrame, Dict[str, Any]]] = []
        self.df_by_sheet_ready: Dict[str, pd.DataFrame] = {}  # post-header-build, post-normalisation
        self.columns_current: List[str] = []  # for rule dropdown

        # Options
        self.var_all_sheets = tk.BooleanVar(value=False)
        self.var_normalise = tk.BooleanVar(value=True)
        self.var_combine = tk.StringVar(value="AND")
        self.var_keep_matches = tk.BooleanVar(value=True)
        self.var_case_sensitive = tk.BooleanVar(value=False)  # default for new rules
        self.var_out_excel = tk.BooleanVar(value=True)
        self.var_out_csv = tk.BooleanVar(value=False)
        self.var_separate_sheets = tk.BooleanVar(value=True)

        # Build UI
        self._build_ui()

    # --------------- UI construction ---------------

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)

        # File frame
        frm_file = ttk.LabelFrame(self.root, text="1) Main file")
        frm_file.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        frm_file.columnconfigure(1, weight=1)

        ttk.Button(frm_file, text="Load Excel/CSV...", command=self.load_main).grid(row=0, column=0, padx=(6, 8), pady=6, sticky="w")
        self.lbl_file = ttk.Label(frm_file, text="No file loaded")
        self.lbl_file.grid(row=0, column=1, sticky="w")

        ttk.Checkbutton(frm_file, text="Apply to all sheets (Excel only)", variable=self.var_all_sheets, command=self._on_all_sheets_toggle).grid(row=1, column=0, columnspan=2, sticky="w", padx=6, pady=(0,6))
        ttk.Checkbutton(frm_file, text="Normalise column names", variable=self.var_normalise, command=self._rebuild_ready_frames).grid(row=2, column=0, columnspan=2, sticky="w", padx=6)

        # Combine frame
        frm_opts = ttk.LabelFrame(self.root, text="2) Combine and output options")
        frm_opts.grid(row=1, column=0, sticky="ew", padx=10, pady=8)
        for i in range(4):
            frm_opts.columnconfigure(i, weight=1)

        ttk.Label(frm_opts, text="Combine rules with:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Radiobutton(frm_opts, text="AND", value="AND", variable=self.var_combine).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(frm_opts, text="OR", value="OR", variable=self.var_combine).grid(row=0, column=2, sticky="w")

        ttk.Checkbutton(frm_opts, text="Keep rows that match (untick to exclude matches)", variable=self.var_keep_matches).grid(row=1, column=0, columnspan=3, sticky="w", padx=6)

        ttk.Label(frm_opts, text="Output formats:").grid(row=2, column=0, sticky="w", padx=6)
        ttk.Checkbutton(frm_opts, text="Excel (.xlsx)", variable=self.var_out_excel).grid(row=2, column=1, sticky="w")
        ttk.Checkbutton(frm_opts, text="CSV (.csv)", variable=self.var_out_csv).grid(row=2, column=2, sticky="w")
        ttk.Checkbutton(frm_opts, text="Excel: separate sheets (if multiple)", variable=self.var_separate_sheets).grid(row=2, column=3, sticky="w")

        # Rules frame
        frm_rules = ttk.LabelFrame(self.root, text="3) Rules")
        frm_rules.grid(row=2, column=0, sticky="ew", padx=10, pady=8)
        frm_rules.columnconfigure(0, weight=1)
        frm_rules.columnconfigure(1, weight=1)
        frm_rules.columnconfigure(2, weight=1)
        frm_rules.columnconfigure(3, weight=1)

        # Rule creator controls
        ttk.Label(frm_rules, text="Column:").grid(row=0, column=0, sticky="w", padx=6)
        self.cmb_column = ttk.Combobox(frm_rules, values=[], state="readonly")
        self.cmb_column.grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(frm_rules, text="Operator:").grid(row=0, column=2, sticky="w", padx=6)
        self.cmb_operator = ttk.Combobox(frm_rules, values=[
            "is", "is_not",
            "contains_any", "not_contains_any",
            "regex_any", "not_regex",
            "gt", "gte", "lt", "lte",
            "between",
            "in_list_file", "not_in_list_file",
            "is_empty", "not_empty"
        ], state="readonly")
        self.cmb_operator.grid(row=0, column=3, sticky="ew", padx=6)
        self.cmb_operator.bind("<<ComboboxSelected>>", lambda e: self._update_rule_inputs())
        self.cmb_operator.set("contains_any")

        # Dynamic inputs area
        self.rule_inputs = ttk.Frame(frm_rules)
        self.rule_inputs.grid(row=1, column=0, columnspan=4, sticky="ew", padx=6, pady=(4, 6))
        self.rule_inputs.columnconfigure(1, weight=1)
        self._build_dynamic_rule_inputs()

        # Buttons for rule list
        btns = ttk.Frame(frm_rules)
        btns.grid(row=2, column=0, columnspan=4, sticky="ew", padx=6)
        ttk.Button(btns, text="Add rule", command=self.add_rule).grid(row=0, column=0, padx=(0,6))
        ttk.Button(btns, text="Remove selected", command=self.remove_selected_rule).grid(row=0, column=1, padx=(0,6))
        ttk.Button(btns, text="Clear rules", command=self.clear_rules).grid(row=0, column=2, padx=(0,6))

        # Rules list
        self.rules_tree = ttk.Treeview(frm_rules, columns=("column","operator","details"), show="headings", height=6)
        self.rules_tree.heading("column", text="Column")
        self.rules_tree.heading("operator", text="Operator")
        self.rules_tree.heading("details", text="Details")
        self.rules_tree.column("column", width=180, anchor="w")
        self.rules_tree.column("operator", width=140, anchor="w")
        self.rules_tree.column("details", width=500, anchor="w")
        self.rules_tree.grid(row=3, column=0, columnspan=4, sticky="ew", padx=6, pady=(6,0))

        # Action frame
        frm_act = ttk.LabelFrame(self.root, text="4) Run")
        frm_act.grid(row=3, column=0, sticky="nsew", padx=10, pady=8)
        frm_act.columnconfigure(0, weight=1)
        frm_act.rowconfigure(1, weight=1)
        # Keep the first row tall enough so buttons remain visible even with scaling
        frm_act.grid_rowconfigure(0, minsize=56)

        act_btns = ttk.Frame(frm_act)
        act_btns.grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Button(act_btns, text="Preview (F5)", command=self.preview).grid(row=0, column=0, padx=(0,8))
        ttk.Button(act_btns, text="Save... (Ctrl+S)", command=self.save_outputs).grid(row=0, column=1, padx=(0,8))

        # Preview label
        self.lbl_preview = ttk.Label(frm_act, text="No preview yet.")
        self.lbl_preview.grid(row=0, column=1, sticky="e", padx=6)

        # Log box
        self.txt_log = tk.Text(frm_act, height=10)
        self.txt_log.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=6, pady=(4,6))
        self._log("Ready. Load a main file to begin.")

    def _build_dynamic_rule_inputs(self):
        # Destroy previous widgets
        for w in self.rule_inputs.winfo_children():
            w.destroy()

        # String/regex values: a multi-line textbox (one value per line)
        ttk.Label(self.rule_inputs, text="Values (one per line):").grid(row=0, column=0, sticky="nw")
        self.txt_values = tk.Text(self.rule_inputs, height=5)
        self.txt_values.grid(row=0, column=1, sticky="ew", pady=4, padx=(8,0))
        # Buttons for loading/clearing values
        btns = ttk.Frame(self.rule_inputs)
        btns.grid(row=0, column=2, sticky="nw", padx=6)
        ttk.Button(btns, text="Load values from file...", command=self.load_values_from_file_into_text).grid(row=0, column=0, pady=(0,6), sticky="w")
        ttk.Button(btns, text="Clear", command=lambda: self.txt_values.delete("1.0", "end")).grid(row=1, column=0, sticky="w")

        # Case sensitive
        self.chk_case = ttk.Checkbutton(self.rule_inputs, text="Case sensitive", variable=self.var_case_sensitive)
        self.chk_case.grid(row=1, column=0, columnspan=2, sticky="w", pady=(0,6))

        # Numeric inputs
        self.frm_numeric = ttk.Frame(self.rule_inputs)
        self.frm_numeric.grid(row=2, column=0, columnspan=3, sticky="ew")
        self.frm_numeric.columnconfigure(1, weight=1)
        ttk.Label(self.frm_numeric, text="Value:").grid(row=0, column=0, sticky="w")
        self.ent_value = ttk.Entry(self.frm_numeric)
        self.ent_value.grid(row=0, column=1, sticky="ew", padx=6)

        # Between inputs
        self.frm_between = ttk.Frame(self.rule_inputs)
        self.frm_between.grid(row=3, column=0, columnspan=3, sticky="ew")
        for i in range(4):
            self.frm_between.columnconfigure(i, weight=1)
        ttk.Label(self.frm_between, text="Min:").grid(row=0, column=0, sticky="w")
        self.ent_min = ttk.Entry(self.frm_between)
        self.ent_min.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Label(self.frm_between, text="Max:").grid(row=0, column=2, sticky="w")
        self.ent_max = ttk.Entry(self.frm_between)
        self.ent_max.grid(row=0, column=3, sticky="ew", padx=6)

        # In-list file selector
        self.frm_listfile = ttk.Frame(self.rule_inputs)
        self.frm_listfile.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(6,0))
        self.frm_listfile.columnconfigure(1, weight=1)
        ttk.Label(self.frm_listfile, text="List file:").grid(row=0, column=0, sticky="w")
        self.ent_listfile = ttk.Entry(self.frm_listfile)
        self.ent_listfile.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(self.frm_listfile, text="Browse...", command=self._browse_listfile).grid(row=0, column=2, sticky="w")

        self._update_rule_inputs()  # set initial visibility based on operator

    def _update_rule_inputs(self):
        op = (self.cmb_operator.get() or "").strip()

        # Show/hide blocks based on operator
        def hide_all():
            self.txt_values.grid()
            self.txt_values.grid_remove()
            self.chk_case.grid()
            self.chk_case.grid_remove()
            self.frm_numeric.grid()
            self.frm_numeric.grid_remove()
            self.frm_between.grid()
            self.frm_between.grid_remove()
            self.frm_listfile.grid()
            self.frm_listfile.grid_remove()

        hide_all()

        if op in ("is", "is_not", "contains_any", "not_contains_any", "regex_any", "not_regex"):
            self.txt_values.grid()
            self.chk_case.grid()
        elif op in ("in_list_file", "not_in_list_file"):
            self.frm_listfile.grid()
            self.chk_case.grid()
        elif op in ("gt", "gte", "lt", "lte"):
            self.frm_numeric.grid()
        elif op == "between":
            self.frm_between.grid()
        elif op in ("is_empty", "not_empty"):
            # no inputs needed
            pass

    # --------------- Log helper ---------------

    def _log(self, msg: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")
        self.root.update_idletasks()

    # --------------- Actions ---------------

    def _on_all_sheets_toggle(self):
        # Reload if a file is already set
        if self.main_path:
            self._load_main_internal(self.main_path)

    def _rebuild_ready_frames(self):
        # Rebuild ready frames and update column list when normalisation toggled
        if not self.units:
            return
        self._prepare_ready_dataframes()
        self._update_columns_dropdown()

    def load_main(self):
        path = filedialog.askopenfilename(
            title="Select main Excel/CSV file",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        self._load_main_internal(path)

    def _load_main_internal(self, path: str):
        ext = os.path.splitext(path)[1].lower()
        if ext not in SUPPORTED_EXTS:
            messagebox.showwarning("Unsupported", "Please select a .xlsx or .csv file.")
            return
        try:
            self._log(f"Loading main file: {path}")
            units = load_main_source(path, all_sheets=self.var_all_sheets.get())
            self.main_path = path
            self.units = units
            self.lbl_file.config(text=f"{os.path.basename(path)} ({len(units)} sheet(s))")
            self._prepare_ready_dataframes()
            self._update_columns_dropdown()
            self._log(f"Loaded {len(units)} sheet(s). Columns detected for rules: {self.columns_current}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load main file:\n{e}")
            self._log(f"Error: {e}")

    def _prepare_ready_dataframes(self):
        # Build per-sheet df (post header build), and apply normalisation if selected
        self.df_by_sheet_ready.clear()
        normalise = self.var_normalise.get()
        for sheet_name, df_raw, meta in self.units:
            df = build_df_from_unit(df_raw, meta)
            if normalise:
                df = apply_column_normalisation(df, True)
            self.df_by_sheet_ready[sheet_name] = df

    def _update_columns_dropdown(self):
        # Use columns from the first sheet for rule selection
        if not self.df_by_sheet_ready:
            self.columns_current = []
        else:
            first_sheet = list(self.df_by_sheet_ready.keys())[0]
            self.columns_current = list(self.df_by_sheet_ready[first_sheet].columns)
        self.cmb_column["values"] = self.columns_current
        if self.columns_current:
            self.cmb_column.set(self.columns_current[0])
        else:
            self.cmb_column.set("")

    def load_values_from_file_into_text(self):
        path = filedialog.askopenfilename(
            title="Select list file (.xlsx/.csv)",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            vals = read_list_values(path)
            self.txt_values.delete("1.0", "end")
            if vals:
                self.txt_values.insert("1.0", "\n".join(vals))
                self._log(f"Loaded {len(vals)} values from {os.path.basename(path)}.")
            else:
                self._log(f"No values found in {os.path.basename(path)}.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read list file:\n{e}")
            self._log(f"Error reading list file: {e}")

    def _browse_listfile(self):
        path = filedialog.askopenfilename(
            title="Select list file (.xlsx/.csv)",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        self.ent_listfile.delete(0, "end")
        self.ent_listfile.insert(0, path)

    # --------------- Rules handling ---------------

    def _collect_rule_from_inputs(self) -> Optional[Dict[str, Any]]:
        col = (self.cmb_column.get() or "").strip()
        if not col:
            messagebox.showwarning("Missing", "Please choose a column.")
            return None
        if col not in self.columns_current:
            messagebox.showwarning("Column not found", "Selected column is not in the loaded data.")
            return None
        op = (self.cmb_operator.get() or "").strip()
        if not op:
            messagebox.showwarning("Missing", "Please choose an operator.")
            return None

        rule: Dict[str, Any] = {"column": col, "operator": op}

        if op in ("is", "is_not", "contains_any", "not_contains_any", "regex_any", "not_regex"):
            raw = self.txt_values.get("1.0", "end").strip()
            vals = [line.strip() for line in raw.splitlines() if line.strip() != ""]
            if not vals:
                messagebox.showwarning("Values required", "Please provide at least one value (one per line) or choose a different operator.")
                return None
            rule["values"] = vals
            rule["case_sensitive"] = bool(self.var_case_sensitive.get())

        elif op in ("in_list_file", "not_in_list_file"):
            path = self.ent_listfile.get().strip()
            if not path:
                messagebox.showwarning("List file required", "Please choose a list file.")
                return None
            rule["file"] = path
            rule["case_sensitive"] = bool(self.var_case_sensitive.get())

        elif op in ("gt", "gte", "lt", "lte"):
            val = self.ent_value.get().strip()
            try:
                rule["value"] = float(val)
            except Exception:
                messagebox.showwarning("Invalid number", "Please enter a valid numeric value.")
                return None

        elif op == "between":
            vmin = self.ent_min.get().strip()
            vmax = self.ent_max.get().strip()
            try:
                vmin_f = float(vmin); vmax_f = float(vmax)
                if vmin_f > vmax_f:
                    messagebox.showwarning("Range error", "Min must be less than or equal to Max.")
                    return None
                rule["value_min"] = vmin_f
                rule["value_max"] = vmax_f
            except Exception:
                messagebox.showwarning("Invalid numbers", "Please enter valid numeric Min/Max.")
                return None

        elif op in ("is_empty", "not_empty"):
            # no additional inputs required
            pass

        else:
            messagebox.showwarning("Unsupported", f"Unsupported operator: {op}")
            return None

        return rule

    def clear_rules(self):
        self.rules_tree.delete(*self.rules_tree.get_children())
        setattr(self, "_rules_store", {})  # reset store
        self._log("Cleared all rules.")

    def remove_selected_rule(self):
        sel = self.rules_tree.selection()
        if not sel:
            return
        for iid in sel:
            self.rules_tree.delete(iid)
            if hasattr(self, "_rules_store"):
                self._rules_store.pop(iid, None)
        self._log("Removed selected rule(s).")

    def add_rule(self):
        rule = self._collect_rule_from_inputs()
        if rule is None:
            return
        details = self._rule_details_text(rule)
        iid = self.rules_tree.insert("", "end", values=(rule["column"], rule["operator"], details))
        if not hasattr(self, "_rules_store"):
            self._rules_store: Dict[str, Dict[str, Any]] = {}
        self._rules_store[iid] = rule
        # Clear inputs
        if "values" in rule:
            self.txt_values.delete("1.0", "end")
        self.ent_value.delete(0, "end")
        self.ent_min.delete(0, "end")
        self.ent_max.delete(0, "end")
        self._log(f"Added rule: {rule}")

    def _rule_details_text(self, rule: Dict[str, Any]) -> str:
        op = rule["operator"]
        if op in ("is", "is_not", "contains_any", "not_contains_any", "regex_any", "not_regex"):
            cs = "case-sensitive" if rule.get("case_sensitive") else "case-insensitive"
            preview = ", ".join(rule.get("values", [])[:5])
            more = "" if len(rule.get("values", [])) <= 5 else f" (+{len(rule['values'])-5} more)"
            return f"{cs}; values: {preview}{more}"
        elif op in ("in_list_file", "not_in_list_file"):
            cs = "case-sensitive" if rule.get("case_sensitive") else "case-insensitive"
            return f"{cs}; file: {rule.get('file','')}"
        elif op in ("gt","gte","lt","lte"):
            return f"value: {rule.get('value')}"
        elif op == "between":
            return f"{rule.get('value_min')} to {rule.get('value_max')}"
        elif op == "is_empty":
            return "empty cells"
        elif op == "not_empty":
            return "non-empty cells"
        return ""

    def _get_rules_list(self) -> List[Dict[str, Any]]:
        rules: List[Dict[str, Any]] = []
        if hasattr(self, "_rules_store"):
            for iid in self.rules_tree.get_children(""):
                r = self._rules_store.get(iid)
                if r:
                    rules.append(r)
        return rules

    # --------------- Filtering and preview ---------------

    def _apply_to_all_sheets(self) -> List[Tuple[str, pd.DataFrame]]:
        rules = self._get_rules_list()
        if not rules:
            raise RuntimeError("No rules defined.")
        if not self.df_by_sheet_ready:
            raise RuntimeError("No data loaded.")
        combine_mode = self.var_combine.get()
        keep_matches = self.var_keep_matches.get()

        outputs: List[Tuple[str, pd.DataFrame]] = []
        for name, df in self.df_by_sheet_ready.items():
            # Validate columns exist in this sheet as well
            for r in rules:
                if r["column"] not in df.columns:
                    raise KeyError(f"Column '{r['column']}' not found in sheet '{name}'.")
            masks = [apply_rule(df, r) for r in rules]
            combined = combine_masks(masks, combine_mode)
            if keep_matches:
                out = df[combined]
            else:
                out = df[~combined]
            outputs.append((name, out.reset_index(drop=True)))
        return outputs

    def preview(self):
        try:
            outputs = self._apply_to_all_sheets()
            msg_lines = []
            for name, out in outputs:
                total = len(self.df_by_sheet_ready[name])
                msg_lines.append(f"{name}: {total} -> {len(out)} rows")
            msg = " | ".join(msg_lines)
            self.lbl_preview.config(text=msg)
            self._log("Preview: " + msg)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log(f"Error during preview: {e}")

    def save_outputs(self):
        if not (self.var_out_excel.get() or self.var_out_csv.get()):
            messagebox.showwarning("Output format", "Please select at least one output format (Excel or CSV).")
            return
        try:
            outputs = self._apply_to_all_sheets()
        except Exception as e:
            messagebox.showerror("Error", f"Error before save: {e}")
            self._log(f"Error before save: {e}")
            return

        # Choose base save path (we will append timestamp and extension(s))
        save_path = filedialog.asksaveasfilename(
            title="Save filtered output as...",
            initialfile="Filtered.xlsx" if self.var_out_excel.get() and not self.var_out_csv.get() else "Filtered",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("CSV", "*.csv"), ("All files", "*.*")]
        )
        if not save_path:
            return
        out_dir = os.path.dirname(save_path) or "."
        base = os.path.splitext(os.path.basename(save_path))[0]

        try:
            written = write_outputs(
                df_by_sheet=outputs,
                out_dir=out_dir,
                base_name=base,
                write_excel=self.var_out_excel.get(),
                write_csv=self.var_out_csv.get(),
                separate_sheets=self.var_separate_sheets.get()
            )
            for w in written:
                self._log(f"Wrote: {w}")
            messagebox.showinfo("Done", "Outputs saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write outputs:\n{e}")
            self._log(f"Error writing outputs: {e}")

# ===================== Main =====================

def main():
    root = tk.Tk()
    # Apply a nice theme if available
    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    app = FilterGUI(root)
    # Keep window comfortably tall so the Run buttons stay visible
    root.minsize(1000, 720)

    # Keyboard shortcuts
    root.bind("<F5>", lambda e: app.preview())
    root.bind("<Control-s>", lambda e: app.save_outputs())

    root.mainloop()

if __name__ == "__main__":
    main()
