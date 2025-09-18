import os
import sys
import re
import math
import argparse
import datetime
from typing import List, Tuple, Optional, Dict, Any

import pandas as pd

# Tkinter is part of the standard library for GUI dialogs
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except Exception:
    tk = None  # CLI mode can still work without tkinter

SUPPORTED_EXTS = {".xlsx", ".csv"}  # extend if you wish (e.g., ".xls")


def timestamp() -> str:
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def is_supported_file(path: str) -> bool:
    ext = os.path.splitext(path)[1].lower()
    return ext in SUPPORTED_EXTS


def normalise_column_name(col: object, do_normalise: bool) -> str:
    s = str(col)
    if not do_normalise:
        return s
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s


def apply_column_normalisation(df: pd.DataFrame, do_normalise: bool) -> pd.DataFrame:
    if not do_normalise:
        return df
    # Handle duplicate column names after normalisation by de-duplicating
    new_cols = [normalise_column_name(c, True) for c in df.columns]
    counts: Dict[str, int] = {}
    final_cols: List[str] = []
    for c in new_cols:
        if c in counts:
            counts[c] += 1
            final_cols.append(f"{c}__{counts[c]}")
        else:
            counts[c] = 0
            final_cols.append(c)
    df = df.copy()
    df.columns = final_cols
    return df


# --------------------- Header auto-detection helpers ---------------------

def _safe_str(v: object) -> str:
    # Return empty string for None/NaN, else stripped string
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    try:
        if pd.isna(v):  # covers pandas NA/NaT
            return ""
    except Exception:
        pass
    return str(v).strip()


def _is_date_like(s: str) -> bool:
    s = s.strip()
    if not s:
        return False
    # Simple date-like patterns: 2024-09-15, 15/09/2024, 15-09-24, 2024/09/15
    return bool(re.match(r"^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$", s))


def _is_numeric_like(s: str) -> bool:
    s = s.strip()
    if not s:
        return False
    # Numbers incl. decimals, thousands, negatives, percents
    if s.endswith("%"):
        s = s[:-1]
    return bool(re.match(r"^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$", s))


def _token_is_headerish(v: object) -> bool:
    s = _safe_str(v)
    if not s:
        return False
    if len(s) > 120:
        return False
    # Header tokens usually contain letters/underscores; filter out pure numbers/dates
    if _is_numeric_like(s) or _is_date_like(s):
        return False
    return bool(re.search(r"[A-Za-z_]", s))


def _row_score(cells: list) -> Tuple[float, float, float, int, int]:
    """
    Returns:
      score (float),
      headerish_ratio (0..1),
      unique_ratio (0..1),
      nonempty_count,
      total_count
    """
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
        name = _safe_str(c)
        if not name:
            name = "column"
        if name in seen:
            seen[name] += 1
            out.append(f"{name}__{seen[name]}")
        else:
            seen[name] = 0
            out.append(name)
    return out


def _maybe_merge_two_row_header(df: pd.DataFrame, h: int) -> Tuple[List[str], int]:
    """
    Returns (headers, rows_consumed_for_header).
    If a second row looks like header continuation, merge row h and h+1.
    """
    row1 = list(df.iloc[h].tolist())
    # If there is no next row, return row1
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

    # Heuristic: merge if row2 also looks header-ish and row1 has empties or duplicates
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
    """
    Analyse a raw DataFrame (header=None) to locate the most likely header row.
    Returns a dict with detection metadata.
    """
    meta: Dict[str, Any] = {
        "best_idx": 0,
        "best_score": float("-inf"),
        "best_headerish_ratio": 0.0,
        "best_unique_ratio": 0.0,
        "rows_used": 0,
        "headers": [],
        "classification": "none",  # 'strong' | 'weak' | 'none'
        "data_start": 0,
        "ncols": int(df_raw.shape[1]),
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

    # Inspect first few data rows after data_start to see if they look like data
    data_rows = df_raw.iloc[data_start : min(len(df_raw), data_start + 5)]
    if len(data_rows) == 0:
        data_like_ratio = 0.0
    else:
        # proportion of tokens that are numeric/date-like or empty
        tokens = [_safe_str(v) for v in data_rows.values.flatten().tolist()]
        nonempty = [t for t in tokens if t != ""]
        if not nonempty:
            data_like_ratio = 0.0
        else:
            dataish = sum(1 for t in nonempty if (_is_numeric_like(t) or _is_date_like(t) or not _token_is_headerish(t)))
            data_like_ratio = dataish / max(1, len(nonempty))

    # Classification thresholds (heuristic but pragmatic):
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
    # Adjust header length vs number of columns
    if len(df.columns) > len(headers):
        headers = headers + [f"column_{i+1}" for i in range(len(df.columns) - len(headers))]
    elif len(df.columns) < len(headers):
        headers = headers[:len(df.columns)]
    df.columns = headers
    # Clean up
    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(axis=1, how="all")
    df = df.reset_index(drop=True).convert_dtypes()
    return df


def build_placeholder_headers(ncols: int) -> List[str]:
    return [f"column_{i+1}" for i in range(ncols)]


# --------------------- Readers (CSV/Excel) into raw frames ---------------------

def read_csv_raw(path: str) -> pd.DataFrame:
    encodings = [None, "utf-8", "utf-8-sig", "utf-16", "latin-1"]
    last_err = None
    for enc in encodings:
        try:
            # Read as raw (no header) so we can detect where the header actually is
            return pd.read_csv(path, sep=None, engine="python", encoding=enc, header=None, dtype=str)
        except UnicodeDecodeError as e:
            last_err = e
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Failed to read CSV: {path} ({last_err})")


def read_excel_raw_all_sheets(path: str) -> List[Tuple[str, pd.DataFrame]]:
    try:
        sheets_raw: Dict[str, pd.DataFrame] = pd.read_excel(path, sheet_name=None, header=None, dtype=str)
    except ImportError as e:
        raise RuntimeError("openpyxl is required for Excel support. pip install openpyxl") from e
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel: {path} ({e})")
    items: List[Tuple[str, pd.DataFrame]] = []
    for sheet, df_raw in sheets_raw.items():
        items.append((str(sheet), df_raw if df_raw is not None else pd.DataFrame()))
    return items


# --------------------- Core load/combine with canonical header alignment ---------------------

class Unit:
    def __init__(self, source_file: str, source_sheet: str, df_raw: pd.DataFrame):
        self.source_file = source_file
        self.source_sheet = source_sheet
        self.df_raw = df_raw if df_raw is not None else pd.DataFrame()
        self.meta = analyse_header(self.df_raw)
        self.classification: str = self.meta["classification"]  # 'strong' | 'weak' | 'none'
        self.ncols: int = self.meta["ncols"]
        # Initial build: if we have some header, use it; else placeholder
        if self.classification in ("strong", "weak"):
            self.headers_detected: List[str] = list(self.meta["headers"])
            self.data_start: int = int(self.meta["data_start"])
            self.df_initial: pd.DataFrame = build_dataframe(self.df_raw, self.headers_detected, self.data_start)
        else:
            self.headers_detected = build_placeholder_headers(self.ncols)
            self.data_start = 0
            self.df_initial = build_dataframe(self.df_raw, self.headers_detected, self.data_start)

        # Will be set later if aligned to canonical
        self.headers_final: List[str] = list(self.df_initial.columns)
        self.df_final: pd.DataFrame = self.df_initial


def normalise_headers_for_key(headers: List[str]) -> Tuple[str, ...]:
    # Normalise headers for grouping/comparison (trim spaces, collapse spaces, lower)
    normed = [re.sub(r"\s+", " ", str(h)).strip().lower() for h in headers]
    return tuple(normed)


def build_canonical_schemas(units: List[Unit]) -> Tuple[Optional[List[str]], Dict[int, List[str]]]:
    """
    Returns:
      - primary canonical header list (most frequent among strong units; fallback to weak)
      - per-width canonical headers: dict col_count -> header list (most common for that width)
    """
    # Collect candidates from strong, then weak
    buckets: Dict[Tuple[str, ...], Dict[str, Any]] = {}
    width_buckets: Dict[int, Dict[Tuple[str, ...], int]] = {}

    def add_candidate(headers: List[str], strength: str, score: float):
        key = normalise_headers_for_key(headers)
        rec = buckets.get(key)
        if rec is None:
            rec = {"count": 0, "score_sum": 0.0, "headers_sample": headers, "strength_counts": {"strong": 0, "weak": 0}}
            buckets[key] = rec
        rec["count"] += 1
        rec["score_sum"] += score
        rec["strength_counts"][strength] = rec["strength_counts"].get(strength, 0) + 1
        width = len(headers)
        if width not in width_buckets:
            width_buckets[width] = {}
        width_buckets[width][key] = width_buckets[width].get(key, 0) + 1

    for u in units:
        if u.classification in ("strong", "weak"):
            add_candidate(u.headers_detected, u.classification, float(u.meta["best_score"]))

    # Determine primary canonical: prefer strong dominance by count, then score
    primary_headers: Optional[List[str]] = None
    if buckets:
        ranked = sorted(
            buckets.items(),
            key=lambda kv: (
                kv[1]["strength_counts"].get("strong", 0),
                kv[1]["count"],
                kv[1]["score_sum"]
            ),
            reverse=True
        )
        primary_headers = list(ranked[0][1]["headers_sample"])

    # Per-width canonicals
    per_width: Dict[int, List[str]] = {}
    for width, d in width_buckets.items():
        best_key, best_count = None, -1
        for k, c in d.items():
            if c > best_count:
                best_key, best_count = k, c
            elif c == best_count and primary_headers is not None and len(primary_headers) == width:
                if k == normalise_headers_for_key(primary_headers):
                    best_key, best_count = k, c
        if best_key is not None:
            headers_sample = None
            for key, rec in buckets.items():
                if key == best_key:
                    headers_sample = rec["headers_sample"]
                    break
            if headers_sample is None and primary_headers is not None and len(primary_headers) == width:
                headers_sample = primary_headers
            if headers_sample is None:
                headers_sample = list(best_key)
            per_width[width] = list(headers_sample)

    return primary_headers, per_width


def align_units_to_canonical(units: List[Unit], align_headerless: bool = False) -> None:
    """
    For units classified as 'none' or 'weak', if align_headerless is True,
    align headers to the canonical schema by matching column count.
    """
    if not align_headerless:
        return

    primary, per_width = build_canonical_schemas(units)

    for u in units:
        width = u.ncols
        if width <= 0:
            continue

        should_align = (u.classification in ("none", "weak"))
        if not should_align:
            continue

        target_headers = per_width.get(width)
        if target_headers is None and primary is not None and len(primary) == width:
            target_headers = primary

        if target_headers is None:
            continue

        u.headers_final = list(target_headers)
        # When aligning headerless/weak, treat entire frame as data (start at 0)
        u.df_final = build_dataframe(u.df_raw, u.headers_final, 0)


def load_sources(
    files: List[str],
    normalise_columns: bool = True,
    include_metadata: bool = True,
    csv_sheet_label: str = "(CSV)",
    align_headerless: bool = False
) -> pd.DataFrame:
    """
    Read all files, detect/align headers across sources, union columns, add metadata columns.
    Returns a single concatenated DataFrame. Column order is preserved in the order columns
    are first seen across inputs (no alphabetical sorting).
    """
    units: List[Unit] = []
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext == ".csv":
            df_raw = read_csv_raw(f)
            units.append(Unit(source_file=os.path.basename(f), source_sheet=csv_sheet_label, df_raw=df_raw))
        elif ext == ".xlsx":
            for sheet_name, df_raw in read_excel_raw_all_sheets(f):
                units.append(Unit(source_file=os.path.basename(f), source_sheet=str(sheet_name), df_raw=df_raw))
        else:
            continue

    if not units:
        return pd.DataFrame()

    # Align questionable units (headerless/weak) to canonical header by width (OFF by default)
    align_units_to_canonical(units, align_headerless=align_headerless)

    # Build final frames with metadata and optional column normalisation
    frames: List[pd.DataFrame] = []
    for u in units:
        df = u.df_final.copy()
        df = apply_column_normalisation(df, normalise_columns)
        if include_metadata:
            # Insert metadata at the front so they remain leading columns
            df.insert(0, "source_sheet", u.source_sheet)
            df.insert(0, "source_file", u.source_file)
        frames.append(df)

    # Concatenate without sorting columns
    combined = pd.concat(frames, ignore_index=True, sort=False)

    # Preserve first-seen column order across all frames
    column_order: List[str] = []
    for df in frames:
        for c in df.columns:
            if c not in column_order:
                column_order.append(c)

    combined = combined.reindex(columns=column_order)

    return combined


def write_outputs(
    df: pd.DataFrame,
    out_dir: str,
    base_name: str,
    write_excel: bool,
    write_csv: bool,
    separate_sheets: bool
) -> List[str]:
    """
    Save DataFrame to chosen formats. If separate_sheets is True and Excel output selected,
    writes one sheet per source_sheet (plus an 'All' sheet). Returns list of written paths.
    """
    written: List[str] = []
    safe_base = re.sub(r"[^\w\-.]+", "_", base_name).strip("_") or "combined"
    ts = timestamp()

    if write_excel:
        xlsx_path = os.path.join(out_dir, f"{safe_base}_{ts}.xlsx")
        used_sheet_names = set()
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            if separate_sheets and "source_sheet" in df.columns:
                def sheetify(name: object) -> str:
                    s = _safe_str(name) or "Sheet"
                    s = s.replace(":", " ").replace("/", " ").replace("\\", " ").replace("?", " ").replace("*", " ").replace("[", " ").replace("]", " ")
                    s = s.strip() or "Sheet"
                    if len(s) > 31:
                        s = s[:31]
                    return s

                # All combined
                writer.book  # ensure workbook initialised
                df.to_excel(writer, sheet_name="All", index=False)
                used_sheet_names.add("All")

                for key, group in df.groupby("source_sheet", dropna=False):
                    base = sheetify(key)
                    sheet_name = base
                    i = 1
                    while sheet_name in used_sheet_names:
                        suffix = f"_{i}"
                        if len(base) + len(suffix) > 31:
                            sheet_name = (base[: max(0, 31 - len(suffix))]) + suffix
                        else:
                            sheet_name = base + suffix
                        i += 1
                    group.to_excel(writer, sheet_name=sheet_name, index=False)
                    used_sheet_names.add(sheet_name)
            else:
                df.to_excel(writer, sheet_name="Combined", index=False)
        written.append(xlsx_path)

    if write_csv:
        csv_path = os.path.join(out_dir, f"{safe_base}_{ts}.csv")
        # Use utf-8-sig for Excel compatibility, keep index off
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        written.append(csv_path)

    return written


# --------------------- CLI ---------------------

def run_cli(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Universal Excel Combiner: combine Excel (.xlsx) and CSV files with robust header detection.")
    parser.add_argument("--files", nargs="+", help="Input files (.xlsx, .csv).", required=True)
    parser.add_argument("--outdir", help="Output directory.", default=".")
    parser.add_argument("--basename", help="Base name for output files.", default="combined")
    parser.add_argument("--no-metadata", action="store_true", help="Do not add source_file/source_sheet.")
    parser.add_argument("--no-normalise", action="store_true", help="Do not normalise column names.")
    parser.add_argument("--align-headerless", action="store_true",
                        help="Enable aligning headerless/weak files to the most common header by width (OFF by default).")
    parser.add_argument("--format", nargs="+", choices=["xlsx", "csv", "both"], default=["xlsx"],
                        help="Output format(s). Use 'both' or list both xlsx csv.")
    parser.add_argument("--separate-sheets", action="store_true",
                        help="For Excel output, create one sheet per source_sheet plus an 'All' sheet.")
    args = parser.parse_args(argv)

    files = [f for f in args.files if is_supported_file(f)]
    if not files:
        print("No supported files provided (.xlsx, .csv).", file=sys.stderr)
        return 2

    outdir = args.outdir
    os.makedirs(outdir, exist_ok=True)

    df = load_sources(
        files=files,
        normalise_columns=(not args.no_normalise),
        include_metadata=(not args.no_metadata),
        align_headerless=args.align_headerless,  # default False unless flag is set
    )

    if df.empty:
        print("No data loaded from the provided files.", file=sys.stderr)
        return 1

    fmts = set()
    for f in args.format:
        if f == "both":
            fmts.update(["xlsx", "csv"])
        else:
            fmts.add(f)
    write_excel = "xlsx" in fmts
    write_csv = "csv" in fmts

    written = write_outputs(
        df=df,
        out_dir=outdir,
        base_name=args.basename,
        write_excel=write_excel,
        write_csv=write_csv,
        separate_sheets=args.separate_sheets
    )
    for w in written:
        print(f"Wrote: {w}")
    return 0


# --------------------- GUI ---------------------

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal Excel Combiner")

        self.files: List[str] = []
        self.outdir: str = os.getcwd()

        # Options
        self.include_metadata = tk.BooleanVar(value=True)
        self.normalise_columns = tk.BooleanVar(value=True)
        self.separate_sheets = tk.BooleanVar(value=False)
        self.out_xlsx = tk.BooleanVar(value=True)
        self.out_csv = tk.BooleanVar(value=False)
        self.basename = tk.StringVar(value="Combined")
        self.align_headerless = tk.BooleanVar(value=False)  # disabled by default

        # Layout
        frm = ttk.Frame(root, padding=10)
        frm.grid(sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # File selection
        ttk.Label(frm, text="1) Select source files (.xlsx, .csv):").grid(row=0, column=0, sticky="w")
        ttk.Button(frm, text="Choose files...", command=self.choose_files).grid(row=0, column=1, sticky="e")
        self.files_lbl = ttk.Label(frm, text="No files selected")
        self.files_lbl.grid(row=1, column=0, columnspan=2, sticky="w")

        # Output directory
        ttk.Label(frm, text="2) Select output directory:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Button(frm, text="Choose folder...", command=self.choose_outdir).grid(row=2, column=1, sticky="e", pady=(10, 0))
        self.outdir_lbl = ttk.Label(frm, text=f"Output: {self.outdir}")
        self.outdir_lbl.grid(row=3, column=0, columnspan=2, sticky="w")

        # Options
        opts = ttk.LabelFrame(frm, text="Options")
        opts.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        opts.columnconfigure(0, weight=1)
        ttk.Checkbutton(opts, text="Include metadata columns (source_file, source_sheet)", variable=self.include_metadata).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(opts, text="Normalise column names (trim/collapse spaces)", variable=self.normalise_columns).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(opts, text="Excel: separate output sheets by source_sheet", variable=self.separate_sheets).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(opts, text="Align headerless/weak files to common header by width (optional)", variable=self.align_headerless).grid(row=3, column=0, sticky="w")

        fmt = ttk.LabelFrame(frm, text="Output format")
        fmt.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Checkbutton(fmt, text="Excel (.xlsx)", variable=self.out_xlsx).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(fmt, text="CSV (.csv)", variable=self.out_csv).grid(row=0, column=1, sticky="w")
        ttk.Label(fmt, text="Base file name:").grid(row=1, column=0, sticky="w", pady=(5, 5))
        ttk.Entry(fmt, textvariable=self.basename, width=30).grid(row=1, column=1, sticky="w", pady=(5, 5))

        # Actions
        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btns, text="Combine", command=self.combine).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(btns, text="Quit", command=root.destroy).grid(row=0, column=1)

        # Log
        self.log = tk.Text(frm, height=12, width=90, state="disabled")
        self.log.grid(row=7, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        frm.rowconfigure(7, weight=1)

        # Make UI responsive
        for i in range(2):
            frm.columnconfigure(i, weight=1)

        # Sync checkbox state with options used in combine()
        self.include_metadata.trace_add("write", self._on_option_change)
        self.normalise_columns.trace_add("write", self._on_option_change)
        self.separate_sheets.trace_add("write", self._on_option_change)
        self.out_xlsx.trace_add("write", self._on_option_change)
        self.out_csv.trace_add("write", self._on_option_change)
        self.align_headerless.trace_add("write", self._on_option_change)

    def _on_option_change(self, *args):
        # Placeholder for future dynamic UI changes if needed
        pass

    def log_msg(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Select source files",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All files", "*.*")]
        )
        if not paths:
            return
        self.files = [p for p in paths if is_supported_file(p)]
        self.files_lbl.config(text=f"{len(self.files)} files selected")
        if len(self.files) == 0:
            messagebox.showwarning("No supported files", "Please select .xlsx or .csv files.")

    def choose_outdir(self):
        d = filedialog.askdirectory(title="Select output directory", mustexist=True)
        if d:
            self.outdir = d
            self.outdir_lbl.config(text=f"Output: {self.outdir}")

    def combine(self):
        if not self.files:
            messagebox.showwarning("No files", "Please select at least one .xlsx or .csv file.")
            return
        if not (self.out_xlsx.get() or self.out_csv.get()):
            messagebox.showwarning("No output format", "Please select at least one output format (Excel or CSV).")
            return
        try:
            self.log_msg("Loading and analysing sources...")
            df = load_sources(
                files=self.files,
                normalise_columns=self.normalise_columns.get(),
                include_metadata=self.include_metadata.get(),
                align_headerless=self.align_headerless.get(),  # default False
            )
            if df.empty:
                self.log_msg("No data found in the selected files.")
                messagebox.showinfo("Done", "No data found in the selected files.")
                return
            self.log_msg(f"Loaded {len(df)} rows, {len(df.columns)} columns.")

            written = write_outputs(
                df=df,
                out_dir=self.outdir,
                base_name=self.basename.get() or "combined",
                write_excel=self.out_xlsx.get(),
                write_csv=self.out_csv.get(),
                separate_sheets=self.separate_sheets.get()
            )
            for w in written:
                self.log_msg(f"Wrote: {w}")
            messagebox.showinfo("Success", "Combine completed successfully.")
        except Exception as e:
            self.log_msg(f"Error: {e}")
            messagebox.showerror("Error", str(e))


def run_gui():
    if tk is None:
        print("Tkinter is not available. Please run in CLI mode or install tkinter.", file=sys.stderr)
        sys.exit(2)
    root = tk.Tk()
    # Use a themed style if available
    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    App(root)
    root.mainloop()


def main():
    # If arguments provided, run CLI; otherwise GUI
    if len(sys.argv) > 1:
        rc = run_cli(sys.argv[1:])
        sys.exit(rc)
    else:
        run_gui()


if __name__ == "__main__":
    main()
