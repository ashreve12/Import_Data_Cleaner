from __future__ import annotations

import argparse
import re
import zipfile
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


DEFAULT_SOURCE_FILE_NAME = "Module_Raw_AI.xlsx"
DEFAULT_OUTPUT_DIR_NAME = "AI Output"
EXCEL_FILE_PATTERNS = "*.xlsx *.xlsm *.xltx *.xltm"
FINAL_OUTPUT_FILE_NAME = "Module_Import_Final.xlsx"
STEP_FILE_NAMES = {
    "step1": "Step 1.xlsx",
    "step2": "Step 2.xlsx",
}

STEP1_COLUMN_ORDER = [
    "date",
    "estimated date",
    "hs code",
    "hs description",
    "country of origin",
    "quantity",
    "quantity unit",
    "metric tons",
    "kilograms",
    "month",
    "consignee declared",
    "shipper declared",
    "master consignee (unified)",
    "master shipper",
    "notify name",
    "notify address",
    "master notify name",
    "master notify address",
    "master shipper address",
    "consignee (unified)",
    "consignee (consolidated)",
    "shipper (unified)",
    "world region by country of origin",
    "world region by place of receipt",
    "consignee type",
    "state of port of arrival",
    "port of arrival",
    "us region",
    "consignee city",
    "consignee county",
    "in transit",
    "bill of lading nbr.",
    "short container description",
    "master short container description",
]
STEP1_COLUMN_ORDER_SET = set(STEP1_COLUMN_ORDER)
STEP1_COLUMNS_TO_DELETE = {
    "carrier code",
    "bill master carrier",
    "imo code declared",
    "high cube",
}

# Step 2 Rule 1 — remove rows whose Consignee Declared or Shipper Declared
# contains any of these names (normalized substring match).
STEP2_COMPANY_EXCLUSIONS = [
    "extrusion",
    "SEOUL",
    "VON ARDENNE",
    "FHR ANLAGENBAU",
]

# Step 2 Rule 2 — keep only rows whose Short Container Description OR
# Master Short Container Description contains at least one of these terms
# (normalized substring match).
STEP2_POSITIVE_TERMS = [
    "module",
    "solar module",
    "photovoltaic module",
    "panel",
    "solar panel",
]

# Step 2 Rule 3 — remove rows whose Short Container Description contains any
# of these keywords (normalized substring match).
# LED, EVA, and USB are handled separately with boundary-aware regex.
STEP2_EXCLUSION_KEYWORDS = [
    "accessory",
    "alum",
    "amorphous",
    "amplifier",
    "apparatus",
    "baby",
    "battery",
    "boxes",
    "branket",
    "bug trap",
    "camera",
    "carrier",
    "cell",
    "charger",
    "chip",
    "circuit",
    "clay",
    "conditioner",
    "conduct",
    "connector",
    "construction",
    "controller",
    "convertor",
    "cooking",
    "coupler",
    "children",
    "crystal",
    "current",
    "daughter card",
    "decoration",
    "doors",
    "electronic products",
    "emitting",
    "equipment",
    "flexible",
    "foil",
    "fold",
    "fountain",
    "frame",
    "furniture",
    "glass",
    "hand wash",
    "heat",
    "heater",
    "igniter",
    "infrared",
    "insulation",
    "inverter",
    "jewelry",
    "junction",
    "kit",
    "laminator",
    "lamping",
    "laser",
    "light",
    "lithium",
    "machine",
    "material",
    "men's",
    "mount",
    "mounting",
    "neutral bar",
    "off-grid",
    "outdoor",
    "packaging",
    "part",
    "photo cell",
    "photocell",
    "piezo",
    "plastic",
    "portable",
    "power supply",
    "power",
    "rectifier",
    "rubber",
    "scissor",
    "seal",
    "semiconductor",
    "sensor",
    "shaver",
    "simulator",
    "smart",
    "solar light",
    "steel",
    "supply",
    "system",
    "thyristor",
    "tool",
    "tracking",
    "transducer",
    "transformer",
    "transitor",
    "tshirt",
    "tube",
    "underwear",
    "wire",
]

# These three are matched with word-boundary regex on the original text to avoid
# false positives (e.g. "FLED" for "LED", "EVALUATE" for "EVA").
STEP2_BOUNDARY_EXCLUSIONS = ["LED", "EVA", "USB"]

HS_CODE_REGEX = re.compile(r"(?<!\d)\d{6}(?!\d)")


def sanitize_file_component(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._")
    return cleaned or "module_output"


def normalize(value: object) -> str:
    """Uppercase and strip every non-alphanumeric character.

    Allows matching across punctuation, hyphens, and mid-word spaces:
      "PAN EL"      → "PANEL"
      "VON-ARDENNE" → "VONARDENNE"  (matches normalized "VON ARDENNE")
      "tracki ng"   → "TRACKING"
    """
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


def _make_boundary_pattern(keyword: str) -> re.Pattern[str]:
    return re.compile(rf"(?i)(?<![A-Z0-9]){re.escape(keyword)}(?![A-Z0-9])")


_BOUNDARY_PATTERNS = {kw: _make_boundary_pattern(kw) for kw in STEP2_BOUNDARY_EXCLUSIONS}


def header_key(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def norm_contains(cell_value: object, keyword: str) -> bool:
    """Return True if normalize(cell_value) contains normalize(keyword)."""
    nk = normalize(keyword)
    return bool(nk) and nk in normalize(cell_value)


# ---------------------------------------------------------------------------
# Step 2 — three deterministic rules applied in order
# ---------------------------------------------------------------------------

def _excluded_by_company(consignee: object, shipper: object) -> bool:
    """Rule 1: remove if Consignee Declared or Shipper Declared matches."""
    for pattern in STEP2_COMPANY_EXCLUSIONS:
        if norm_contains(consignee, pattern) or norm_contains(shipper, pattern):
            return True
    return False


def _has_positive_term(short_desc: object, master_desc: object) -> bool:
    """Rule 2: keep only if either description column contains a positive term."""
    for term in STEP2_POSITIVE_TERMS:
        if norm_contains(short_desc, term) or norm_contains(master_desc, term):
            return True
    return False


def _excluded_by_keyword(short_desc: object) -> bool:
    """Rule 3: remove if Short Container Description contains an exclusion keyword."""
    for keyword in STEP2_EXCLUSION_KEYWORDS:
        if norm_contains(short_desc, keyword):
            return True
    raw = str(short_desc or "")
    for keyword in STEP2_BOUNDARY_EXCLUSIONS:
        if _BOUNDARY_PATTERNS[keyword].search(raw):
            return True
    return False


def row_passes_step2(
    short_desc: object,
    master_desc: object,
    shipper: object,
    consignee: object,
) -> bool:
    if _excluded_by_company(consignee, shipper):
        return False
    if not _has_positive_term(short_desc, master_desc):
        return False
    if _excluded_by_keyword(short_desc):
        return False
    return True


# ---------------------------------------------------------------------------
# Workbook helpers
# ---------------------------------------------------------------------------

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
        "run_log": output_dir / f"{safe_stem}_module_processing_run_log.md",
    }


def open_file_dialog(initial_dir: Path) -> Path | None:
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="Select the module workbook",
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


def get_required_columns(ws) -> dict[str, int]:
    headers: dict[str, int] = {}
    for column_index in range(1, ws.max_column + 1):
        normalized = header_key(ws.cell(row=1, column=column_index).value)
        if normalized and normalized not in headers:
            headers[normalized] = column_index

    required = {
        "hs_code": "hs code",
        "consignee_declared": "consignee declared",
        "shipper_declared": "shipper declared",
        "short_container_description": "short container description",
        "master_short_container_description": "master short container description",
    }

    missing = [display_name for display_name in required.values() if display_name not in headers]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    return {key: headers[value] for key, value in required.items()}


def compute_column_order(ws) -> list[int]:
    header_to_col: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        h = header_key(ws.cell(row=1, column=col_idx).value)
        if h and h not in header_to_col:
            header_to_col[h] = col_idx

    ordered: list[int] = []
    seen: set[int] = set()

    for col_name in STEP1_COLUMN_ORDER:
        col_idx = header_to_col.get(col_name)
        if col_idx is not None and col_idx not in seen:
            ordered.append(col_idx)
            seen.add(col_idx)

    for h, col_idx in header_to_col.items():
        if h not in STEP1_COLUMN_ORDER_SET and h not in STEP1_COLUMNS_TO_DELETE and col_idx not in seen:
            ordered.append(col_idx)
            seen.add(col_idx)

    return ordered


def build_snapshot_workbook(
    source_ws,
    rows_to_keep: list[int],
    column_order: list[int],
    apply_calibri: bool = False,
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
        source_row_dimension = source_ws.row_dimensions[source_row_index]
        target_row_dimension = target_ws.row_dimensions[target_row_index]
        target_row_dimension.height = source_row_dimension.height
        target_row_dimension.hidden = source_row_dimension.hidden

        for target_col_idx, source_col_idx in enumerate(column_order, start=1):
            source_cell = source_ws.cell(row=source_row_index, column=source_col_idx)
            target_cell = target_ws.cell(row=target_row_index, column=target_col_idx, value=source_cell.value)
            if source_cell.has_style:
                if apply_calibri:
                    src_font = source_cell.font
                    target_cell.font = Font(
                        name="Calibri",
                        bold=src_font.bold,
                        italic=src_font.italic,
                        vertAlign=src_font.vertAlign,
                        underline=src_font.underline,
                        strike=src_font.strike,
                        color=src_font.color,
                        size=src_font.size,
                        scheme=src_font.scheme,
                    )
                else:
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


def process_module_file(source_path: Path, output_path: Path, zip_output_path: Path) -> dict[str, int | str]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    zip_output_path.parent.mkdir(parents=True, exist_ok=True)

    source_workbook = load_workbook(source_path)
    source_ws = source_workbook["DATA"] if "DATA" in source_workbook.sheetnames else source_workbook.active
    required_columns = get_required_columns(source_ws)
    column_order = compute_column_order(source_ws)

    initial_rows = list(range(2, source_ws.max_row + 1))

    # Step 1 filter: drop rows with more than 4 distinct HS codes.
    step1_rows: list[int] = []
    for row_index in initial_rows:
        hs_code_value = str(source_ws.cell(row=row_index, column=required_columns["hs_code"]).value or "")
        if len(set(HS_CODE_REGEX.findall(hs_code_value))) <= 4:
            step1_rows.append(row_index)

    # Step 2 filter: apply module rules (reads from original source positions).
    step2_rows: list[int] = []
    for row_index in step1_rows:
        consignee = source_ws.cell(row=row_index, column=required_columns["consignee_declared"]).value
        shipper = source_ws.cell(row=row_index, column=required_columns["shipper_declared"]).value
        short_desc = source_ws.cell(row=row_index, column=required_columns["short_container_description"]).value
        master_desc = source_ws.cell(row=row_index, column=required_columns["master_short_container_description"]).value
        if row_passes_step2(short_desc, master_desc, shipper, consignee):
            step2_rows.append(row_index)

    summary: dict[str, int | str] = {
        "source_file": str(source_path),
        "worksheet": source_ws.title,
        "output_directory": str(output_path.parent),
        "initial_data_rows": len(initial_rows),
        "step1_removed_hs_code_rows": len(initial_rows) - len(step1_rows),
        "rows_after_step1": len(step1_rows),
        "step2_removed_keyword_rows": len(step1_rows) - len(step2_rows),
        "final_rows_left": len(step2_rows),
    }

    # Build Step 1 workbook from source: applies column reorder, deleted headers,
    # freeze row, and Calibri font.
    step1_wb = build_snapshot_workbook(source_ws, step1_rows, column_order, apply_calibri=True)
    step1_ws = step1_wb.active
    saved_step1_path = save_workbook(step1_wb, output_path.parent / STEP_FILE_NAMES["step1"])

    # Translate step2 row indices from source-row space into step1_ws row space.
    # step1_ws row 1 = header; row i+2 = step1_rows[i].
    src_to_step1: dict[int, int] = {src: i + 2 for i, src in enumerate(step1_rows)}
    step2_rows_in_step1 = [src_to_step1[r] for r in step2_rows]

    # Columns in step1_ws are already sequentially ordered 1..N.
    step1_col_range = list(range(1, len(column_order) + 1))

    # Step 2 and Final are built from step1_ws so they inherit Calibri and the
    # reordered/deleted column layout without re-applying the transformation.
    step2_wb = build_snapshot_workbook(step1_ws, step2_rows_in_step1, step1_col_range)
    saved_step2_path = save_workbook(step2_wb, output_path.parent / STEP_FILE_NAMES["step2"])
    step2_wb.close()

    final_wb = build_snapshot_workbook(step1_ws, step2_rows_in_step1, step1_col_range)
    saved_output_path = save_workbook(final_wb, output_path)
    final_wb.close()

    step1_wb.close()

    saved_zip_path = create_zip_bundle(zip_output_path, [saved_step1_path, saved_step2_path, saved_output_path])

    summary["output_file"] = str(saved_output_path)
    summary["zip_file"] = str(saved_zip_path)
    summary["step_files"] = f"{saved_step1_path}, {saved_step2_path}"

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
        f"- Step 2 removed (module filter): {summary['step2_removed_keyword_rows']}",
        f"- Final rows left: {summary['final_rows_left']}",
    ]
    with log_path.open("a", encoding="utf-8") as file_handle:
        file_handle.write("\n".join(lines) + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Filter a module workbook and save the processed file, step snapshots, and zip bundle."
    )
    parser.add_argument(
        "--source-file",
        help="Path to the input Excel workbook. If omitted, the script opens a file picker or prompts for a path.",
    )
    parser.add_argument(
        "--output-dir",
        help="Folder for the processed workbook, step snapshots, run log, and zip bundle.",
    )
    parser.add_argument(
        "--zip-name",
        help="Optional zip file name. Defaults to <source_file_name>_step_outputs.zip.",
    )
    parser.add_argument(
        "--no-dialogs",
        action="store_true",
        help="Skip graphical file/folder pickers and use terminal prompts instead.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    workspace_root = Path(__file__).resolve().parents[1]
    use_dialogs = not args.no_dialogs

    source_path = resolve_source_file(args.source_file, workspace_root, use_dialogs)
    output_dir = resolve_output_directory(args.output_dir, workspace_root, use_dialogs)
    run_paths = create_paths_for_run(source_path, output_dir, args.zip_name)

    summary = process_module_file(source_path, run_paths["output_file"], run_paths["zip_file"])
    write_run_log(run_paths["run_log"], summary)

    print("Module processing complete.")
    for key, value in summary.items():
        print(f"{key}: {value}")
    print(f"Run log updated: {run_paths['run_log']}")


if __name__ == "__main__":
    main()
