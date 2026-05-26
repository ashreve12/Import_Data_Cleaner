"""Wafer import data cleaning script."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter


DEFAULT_SOURCE_FILE_NAME = "Wafer_Raw_AI.xlsx"
DEFAULT_OUTPUT_DIR_NAME = "AI Output"
EXCEL_FILE_PATTERNS = "*.xlsx *.xlsm *.xltx *.xltm"
FINAL_OUTPUT_FILE_NAME = "Wafer_Import_Final.xlsx"
FINAL_OUTPUT_SHEET_NAME = "DATA_Final"
STEP_FILE_NAMES = {
    "step0": "Step 0 - Reordered.xlsx",
    "step1": "Step 1.xlsx",
    "step2": "Step 2.xlsx",
    "step3": "Step 3.xlsx",
}

STEP0_COLUMN_ORDER = [
    "date", "estimated date", "hs code", "hs description", "country of origin",
    "quantity", "quantity unit", "metric tons", "kilograms", "month",
    "consignee declared", "shipper declared", "master consignee (unified)",
    "master shipper", "notify name", "notify address", "master notify name",
    "master notify address", "master shipper address", "consignee (unified)",
    "consignee (consolidated)", "shipper (unified)", "world region by country of origin",
    "world region by place of receipt", "consignee type", "state of port of arrival",
    "port of arrival", "us region", "consignee city", "consignee county", "in transit",
    "bill of lading nbr.", "short container description", "master short container description",
    "consignee declared address", "consignee telephone", "consignee email",
    "shipper declared address", "carrier", "master consignee declared address",
    "consignee duns", "consignee dom. ult. duns", "consignee state", "consignee zip code",
    "master/house", "mode of transport", "in bond entry type", "imo code", "bill master",
    "hs code (2)", "hs code (4)", "foreign destination", "world region by port of departure",
    "country by port of departure", "port of departure", "vessel", "vessel country",
    "final destination", "place of receipt", "country by place of receipt", "weight",
    "weight unit", "measure", "measure unit", "container quantity", "fcl/lcl",
    "calculated value by hs (hs)", "calculated value by hs (teus)",
    "calculated value by hs (container quantity)", "calculated value by hs (metric tons)",
]
STEP0_COLUMN_ORDER_SET = set(STEP0_COLUMN_ORDER)
STEP0_COLUMNS_TO_DELETE = {"carrier code", "bill master carrier", "imo code declared", "high cube"}

def _get_keywords(list_name: str, builtin: list) -> list:
    path = Path(__file__).parent.parent / "keywords.json"
    if not path.exists():
        return builtin
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        entry = data.get("wafer", {}).get(list_name, {})
        if isinstance(entry, list):
            return builtin + entry
        custom = entry.get("custom", [])
        removed = set(entry.get("removed", []))
        return [kw for kw in builtin if kw not in removed] + custom
    except Exception:
        return builtin


SUPPLIER_PATTERNS = _get_keywords("supplier_patterns", [
    "ESWIN",
    "Shin-Etsu",
    "SEH AMERICA",
    "S.E.H. AMERICA",
    "Zing Semiconductor",
    "SK SILTRON",
    "silicon valley",
    "Helitek Company",
    "WAFER WORKS CORPORATION",
    "A1 SILICON",
    "A-1 SILCON INC.",
    "SILTRONIC",
    "WAFERNET",
    "ECHEM SOLUTIONS",
    "XXX semiconductor",
    "SOUTHWEST SILICON",
    "SemiStar Corp",
    "PURE WAFER",
    "SAMSUNG AUSTIN SEMICONDUCTOR",
    "LUMENTUM OPERATIONS LLC",
    "Ramco Technology",
    "SCIENTECH CORPORATION",
    "Formosa Sumco Technology Corporation",
    "West European Silicon Technologies",
    "TEKNICAL MATERIAL RECYCLING",
    "INC. RAMCO TECHNOLOGY INC",
    "SOITEC",
    "LUXFER MEL TECHNOLOGIES",
    "WEST EUROPEAN SILICON TECH",
    "MEMC",
    "FOODS",
    "WOLFSPEED",
    "KINIK COMPANY",
    "FORMOSA SUMCO",
])

DIMENSION_TERMS = ["12 inch", "300mm", "8 inch", "200mm"]

EXCLUDE_TERMS_STEP3 = _get_keywords("exclude_terms", [
    "Semi",
    "SEMI CONDUCTOR",
    "CLEANING",
    "ADDITIVE",
    "TEXTURING",
    "Coating",
    "agent",
    "OVEN",
    "GALLIUM ARSENIDE",
    "Mixed",
    "CUTTING",
    "COOLANT",
    "MIXED",
    "RECYCLE",
    "RECLAIM",
    "SEED",
    "TAPE",
    "WAFFY",
    "COOKIES",
    "cakes",
    "peeler",
])

HS_CODE_REGEX = re.compile(r"(?<!\d)\d{6}(?!\d)")
_THIN = Side(style="thin")
_ALL_BORDERS = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def sanitize_file_component(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._")
    return cleaned or "wafer_output"


def normalize_for_match(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Z0-9]+", "", str(value).upper())


def fuzzy_contains(value: object, patterns: list[str]) -> bool:
    norm = normalize_for_match(value)
    if not norm:
        return False
    for pattern in patterns:
        np = normalize_for_match(pattern)
        if np and np in norm:
            return True
    return False


def header_key(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def create_zip_bundle(zip_path: Path, files_to_zip: list[Path]) -> Path:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for file_path in files_to_zip:
            if file_path.exists():
                zip_file.write(file_path, arcname=file_path.name)
    return zip_path


def save_workbook(workbook: Workbook, preferred_path: Path) -> Path:
    preferred_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(preferred_path)
        return preferred_path
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = preferred_path.with_name(f"{preferred_path.stem}_{timestamp}{preferred_path.suffix}")
        workbook.save(fallback_path)
        return fallback_path


def create_paths_for_run(source_path: Path, output_dir: Path, zip_name: str | None = None) -> dict[str, Path]:
    safe_stem = sanitize_file_component(source_path.stem)
    zip_stem = sanitize_file_component(Path(zip_name).stem) if zip_name else f"{safe_stem}_step_outputs"
    return {
        "output_dir": output_dir,
        "output_file": output_dir / FINAL_OUTPUT_FILE_NAME,
        "zip_file": output_dir / f"{zip_stem}.zip",
        "run_log": output_dir / f"{safe_stem}_wafer_processing_run_log.md",
    }


def open_file_dialog(initial_dir: Path) -> Path | None:
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="Select the wafer workbook",
            initialdir=str(initial_dir),
            filetypes=[("Excel workbooks", EXCEL_FILE_PATTERNS), ("All files", "*.*")],
        )
        root.destroy()
        return Path(selected) if selected else None
    except Exception:
        return None


def open_directory_dialog(initial_dir: Path) -> Path | None:
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askdirectory(
            title="Select the output folder",
            initialdir=str(initial_dir),
            mustexist=False,
        )
        root.destroy()
        return Path(selected) if selected else None
    except Exception:
        return None


def prompt_for_existing_file(prompt_text: str, default_path: Path | None) -> Path:
    while True:
        suffix = f" [{default_path}]" if default_path else ""
        raw_value = input(f"{prompt_text}{suffix}: ").strip().strip('"')
        chosen_path = Path(raw_value).expanduser() if raw_value else default_path
        if chosen_path is None:
            print("A source workbook is required.")
            continue
        if chosen_path.is_file():
            return chosen_path
        print(f"File not found: {chosen_path}")


def prompt_for_directory(prompt_text: str, default_path: Path | None) -> Path:
    while True:
        suffix = f" [{default_path}]" if default_path else ""
        raw_value = input(f"{prompt_text}{suffix}: ").strip().strip('"')
        chosen_path = Path(raw_value).expanduser() if raw_value else default_path
        if chosen_path is None:
            print("An output folder is required.")
            continue
        if chosen_path.exists() and not chosen_path.is_dir():
            print(f"Path is not a folder: {chosen_path}")
            continue
        try:
            chosen_path.mkdir(parents=True, exist_ok=True)
            return chosen_path
        except OSError as exc:
            print(f"Could not create folder {chosen_path}: {exc}")


def resolve_source_file(source_arg: str | None, workspace_root: Path, use_dialogs: bool) -> Path:
    if source_arg:
        source_path = Path(source_arg).expanduser()
        if not source_path.is_file():
            raise FileNotFoundError(f"Source file not found: {source_path}")
        return source_path
    default_source = workspace_root / "AI Source" / DEFAULT_SOURCE_FILE_NAME
    initial_dir = default_source.parent if default_source.parent.exists() else workspace_root
    if use_dialogs:
        selected_path = open_file_dialog(initial_dir)
        if selected_path:
            return selected_path
    fallback_default = default_source if default_source.exists() else None
    return prompt_for_existing_file("Enter the source workbook path", fallback_default)


def resolve_output_directory(output_dir_arg: str | None, workspace_root: Path, use_dialogs: bool) -> Path:
    if output_dir_arg:
        output_dir = Path(output_dir_arg).expanduser()
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir
    default_output_dir = workspace_root / DEFAULT_OUTPUT_DIR_NAME
    initial_dir = default_output_dir if default_output_dir.exists() else workspace_root
    if use_dialogs:
        selected_path = open_directory_dialog(initial_dir)
        if selected_path:
            selected_path.mkdir(parents=True, exist_ok=True)
            return selected_path
    return prompt_for_directory("Enter the output folder path", default_output_dir)


def get_source_columns(ws) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        key = header_key(ws.cell(row=1, column=col_idx).value)
        if key and key not in headers:
            headers[key] = col_idx
    for name in ("hs code", "consignee declared", "shipper declared"):
        if name not in headers:
            raise ValueError(f"Missing required column: {name}")
    short_desc_name = next((n for n in ("short container description", "description") if n in headers), None)
    if short_desc_name is None:
        raise ValueError("Missing required column: short container description (or description)")
    master_desc_name = next(
        (n for n in ("master short container description", "short container description", "description") if n in headers),
        short_desc_name,
    )
    return {
        "hs_code": headers["hs code"],
        "consignee_declared": headers["consignee declared"],
        "shipper_declared": headers["shipper declared"],
        "short_container_description": headers[short_desc_name],
        "master_short_container_description": headers[master_desc_name],
    }


def compute_column_order(ws) -> list[int]:
    header_to_col: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        h = header_key(ws.cell(row=1, column=col_idx).value)
        if h and h not in header_to_col:
            header_to_col[h] = col_idx

    ordered: list[int] = []
    seen: set[int] = set()

    for col_name in STEP0_COLUMN_ORDER:
        col_idx = header_to_col.get(col_name)
        if col_idx is not None and col_idx not in seen:
            ordered.append(col_idx)
            seen.add(col_idx)

    for h, col_idx in header_to_col.items():
        if h not in STEP0_COLUMN_ORDER_SET and h not in STEP0_COLUMNS_TO_DELETE and col_idx not in seen:
            ordered.append(col_idx)
            seen.add(col_idx)

    return ordered


def build_snapshot_workbook(
    source_ws,
    rows_to_keep: list[int],
    column_order: list[int],
    apply_formatting: bool = False,
) -> Workbook:
    workbook = Workbook()
    target_ws = workbook.active
    target_ws.title = source_ws.title
    target_ws.freeze_panes = "A2"
    target_ws.sheet_view.showGridLines = source_ws.sheet_view.showGridLines
    target_ws.auto_filter.ref = source_ws.auto_filter.ref

    for target_col_idx, source_col_idx in enumerate(column_order, start=1):
        source_letter = get_column_letter(source_col_idx)
        target_letter = get_column_letter(target_col_idx)
        source_dim = source_ws.column_dimensions[source_letter]
        target_dim = target_ws.column_dimensions[target_letter]
        target_dim.width = source_dim.width
        target_dim.hidden = source_dim.hidden
        target_dim.bestFit = source_dim.bestFit

    all_rows = [1] + rows_to_keep
    for target_row_index, source_row_index in enumerate(all_rows, start=1):
        source_row_dim = source_ws.row_dimensions[source_row_index]
        target_row_dim = target_ws.row_dimensions[target_row_index]
        target_row_dim.height = source_row_dim.height
        target_row_dim.hidden = source_row_dim.hidden

        for target_col_idx, source_col_idx in enumerate(column_order, start=1):
            source_cell = source_ws.cell(row=source_row_index, column=source_col_idx)
            target_cell = target_ws.cell(row=target_row_index, column=target_col_idx, value=source_cell.value)

            if apply_formatting:
                src_font = source_cell.font if source_cell.has_style else None
                target_cell.font = Font(
                    name="Calibri",
                    size=9,
                    bold=src_font.bold if src_font else False,
                    italic=src_font.italic if src_font else False,
                    underline=src_font.underline if src_font else None,
                    strike=src_font.strike if src_font else False,
                    color=copy(src_font.color) if src_font and src_font.color else None,
                )
                target_cell.border = _ALL_BORDERS
                if source_cell.has_style:
                    target_cell.fill = copy(source_cell.fill)
                    target_cell.alignment = copy(source_cell.alignment)
                    target_cell.number_format = source_cell.number_format
            elif source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.protection = copy(source_cell.protection)
                target_cell.number_format = source_cell.number_format

            if source_cell.hyperlink:
                target_cell._hyperlink = copy(source_cell.hyperlink)
            if source_cell.comment:
                target_cell.comment = copy(source_cell.comment)

    return workbook


def process_wafer_file(source_path: Path, output_path: Path, zip_output_path: Path) -> dict[str, int | str]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    zip_output_path.parent.mkdir(parents=True, exist_ok=True)

    source_workbook = load_workbook(source_path)
    source_ws = source_workbook["DATA"] if "DATA" in source_workbook.sheetnames else source_workbook.active
    required_columns = get_source_columns(source_ws)
    column_order = compute_column_order(source_ws)

    initial_rows = list(range(2, source_ws.max_row + 1))

    # Step 0: all rows, columns reordered + Calibri 9 + borders
    step0_wb = build_snapshot_workbook(source_ws, initial_rows, column_order, apply_formatting=True)
    step0_ws = step0_wb.active
    saved_step0_path = save_workbook(step0_wb, output_path.parent / STEP_FILE_NAMES["step0"])

    step0_col_range = list(range(1, len(column_order) + 1))

    # Rebuild required column indices in step0_ws space
    step0_headers: dict[str, int] = {}
    for col_idx in range(1, step0_ws.max_column + 1):
        key = header_key(step0_ws.cell(row=1, column=col_idx).value)
        if key and key not in step0_headers:
            step0_headers[key] = col_idx
    hs_col = step0_headers.get("hs code", required_columns["hs_code"])
    consignee_col = step0_headers.get("consignee declared", required_columns["consignee_declared"])
    shipper_col = step0_headers.get("shipper declared", required_columns["shipper_declared"])
    short_desc_col = (step0_headers.get("short container description")
                      or step0_headers.get("description")
                      or required_columns["short_container_description"])
    master_desc_col = (step0_headers.get("master short container description")
                       or step0_headers.get("short container description")
                       or step0_headers.get("description")
                       or required_columns["master_short_container_description"])

    # Step 1: remove rows with 4+ distinct HS codes
    step1_rows: list[int] = []
    for row_idx in range(2, step0_ws.max_row + 1):
        hs_val = str(step0_ws.cell(row=row_idx, column=hs_col).value or "")
        if len(set(HS_CODE_REGEX.findall(hs_val))) <= 4:
            step1_rows.append(row_idx)

    step1_wb = build_snapshot_workbook(step0_ws, step1_rows, step0_col_range)
    saved_step1_path = save_workbook(step1_wb, output_path.parent / STEP_FILE_NAMES["step1"])
    step1_wb.close()

    # Step 2: remove known semiconductor/non-PV-wafer suppliers
    step2_rows: list[int] = []
    for row_idx in step1_rows:
        consignee = step0_ws.cell(row=row_idx, column=consignee_col).value
        shipper = step0_ws.cell(row=row_idx, column=shipper_col).value
        if not (fuzzy_contains(consignee, SUPPLIER_PATTERNS) or fuzzy_contains(shipper, SUPPLIER_PATTERNS)):
            step2_rows.append(row_idx)

    step2_wb = build_snapshot_workbook(step0_ws, step2_rows, step0_col_range)
    saved_step2_path = save_workbook(step2_wb, output_path.parent / STEP_FILE_NAMES["step2"])
    step2_wb.close()

    # Step 3: keep only rows mentioning WAFER, then remove dimension + exclusion terms
    step3_rows: list[int] = []
    for row_idx in step2_rows:
        short_desc = step0_ws.cell(row=row_idx, column=short_desc_col).value
        master_desc = step0_ws.cell(row=row_idx, column=master_desc_col).value

        has_wafer = fuzzy_contains(short_desc, ["WAFER"]) or fuzzy_contains(master_desc, ["WAFER"])
        if not has_wafer:
            continue

        has_exclusion = fuzzy_contains(short_desc, DIMENSION_TERMS + EXCLUDE_TERMS_STEP3) or \
                        fuzzy_contains(master_desc, DIMENSION_TERMS + EXCLUDE_TERMS_STEP3)
        if has_exclusion:
            continue

        step3_rows.append(row_idx)

    step3_wb = build_snapshot_workbook(step0_ws, step3_rows, step0_col_range)
    saved_step3_path = save_workbook(step3_wb, output_path.parent / STEP_FILE_NAMES["step3"])
    step3_wb.close()

    # Final output: same data as step 3, sheet named DATA_Final
    final_wb = build_snapshot_workbook(step0_ws, step3_rows, step0_col_range)
    final_wb.active.title = FINAL_OUTPUT_SHEET_NAME
    saved_output_path = save_workbook(final_wb, output_path)
    final_wb.close()

    step0_wb.close()

    saved_zip_path = create_zip_bundle(
        zip_output_path,
        [saved_step0_path, saved_step1_path, saved_step2_path, saved_step3_path, saved_output_path],
    )

    summary: dict[str, int | str] = {
        "source_file": str(source_path),
        "worksheet": source_ws.title,
        "output_directory": str(output_path.parent),
        "initial_data_rows": len(initial_rows),
        "step1_removed_hs_code_rows": len(initial_rows) - len(step1_rows),
        "rows_after_step1": len(step1_rows),
        "step2_removed_supplier_rows": len(step1_rows) - len(step2_rows),
        "rows_after_step2": len(step2_rows),
        "step3_removed_rows": len(step2_rows) - len(step3_rows),
        "final_rows_left": len(step3_rows),
        "output_file": str(saved_output_path),
        "zip_file": str(saved_zip_path),
        "step_files": ", ".join(str(p) for p in [saved_step0_path, saved_step1_path, saved_step2_path, saved_step3_path]),
    }

    source_workbook.close()
    return summary


def write_run_log(log_path: Path, summary: dict[str, int | str]) -> None:
    lines = [
        f"\n## {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"- Source: {summary['source_file']}",
        f"- Worksheet: {summary['worksheet']}",
        f"- Output directory: {summary['output_directory']}",
        f"- Output: {summary['output_file']}",
        f"- Zip bundle: {summary['zip_file']}",
        f"- Step snapshots: {summary['step_files']}",
        f"- Initial data rows: {summary['initial_data_rows']}",
        f"- Step 1 removed (HS 4+): {summary['step1_removed_hs_code_rows']}",
        f"- Rows after step 1: {summary['rows_after_step1']}",
        f"- Step 2 removed (supplier filter): {summary['step2_removed_supplier_rows']}",
        f"- Rows after step 2: {summary['rows_after_step2']}",
        f"- Step 3 removed (non-wafer / exclusion terms): {summary['step3_removed_rows']}",
        f"- Final rows left: {summary['final_rows_left']}",
    ]
    with log_path.open("a", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Filter a wafer workbook and save the processed file, step snapshots, and zip bundle."
    )
    parser.add_argument("--source-file", help="Path to the input Excel workbook.")
    parser.add_argument("--output-dir", help="Folder for outputs.")
    parser.add_argument("--zip-name", help="Optional zip file name.")
    parser.add_argument("--no-dialogs", action="store_true", help="Skip graphical pickers.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    workspace_root = Path(__file__).resolve().parents[1]
    use_dialogs = not args.no_dialogs

    source_path = resolve_source_file(args.source_file, workspace_root, use_dialogs)
    output_dir = resolve_output_directory(args.output_dir, workspace_root, use_dialogs)
    run_paths = create_paths_for_run(source_path, output_dir, args.zip_name)

    summary = process_wafer_file(source_path, run_paths["output_file"], run_paths["zip_file"])
    write_run_log(run_paths["run_log"], summary)

    print("Wafer processing complete.")
    for key, value in summary.items():
        print(f"{key}: {value}")
    print(f"Run log updated: {run_paths['run_log']}")


if __name__ == "__main__":
    main()
