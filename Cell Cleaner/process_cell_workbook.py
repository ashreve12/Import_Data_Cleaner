from __future__ import annotations

import argparse
import re
import zipfile
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


DEFAULT_SOURCE_FILE_NAME = "Cell_Raw_AI.xlsx"
DEFAULT_OUTPUT_DIR_NAME = "AI Output"
EXCEL_FILE_PATTERNS = "*.xlsx *.xlsm *.xltx *.xltm"
FINAL_OUTPUT_FILE_NAME = "Cell_Import_Final.xlsx"
STEP_FILE_NAMES = {
    "step1": "Step 1.xlsx",
    "step2": "Step 2.xlsx",
    "step3": "Step 3.xlsx",
}

STEP2_NAME_PATTERNS = [
    "First Solar",
    "FS Solar",
    "Space Exploration",
]

STEP3_KEYWORDS = [
    "panel",
    "module",
    "alum",
    "lithium",
    "rubber",
    "tracking",
    "system",
    "glass",
    "cell",
    "scissor",
    "mounting",
    "lamping",
    "inverter",
    "smart",
    "flexible",
    "off-grid",
    "junction",
    "construction",
    "steel",
    "machine",
    "portable",
    "camera",
    "conditioner",
    "frame",
    "simulator",
    "battery",
    "controller",
    "light",
    "equipment",
    "furniture",
    "emitting",
    "power",
    "supply",
    "laminator",
    "accessory",
    "outdoor",
    "semiconductor",
    "charger",
    "fold",
    "tool",
    "tshirt",
    "plastic",
    "seal",
    "insulation",
    "amorphous",
    "material",
    "part",
    "kit",
    "branket",
]
STEP3_BOUNDARY_KEYWORDS = ["LED", "EVA", "USB"]
STEP3_SHIPPER_PATTERNS = [
    "extrusion",
    "SEOUL",
    "VON ARDENNE",
    "FHR ANLAGENBAU",
]
POSITIVE_CELL_PHRASES = [
    "solar cell",
    "solar cells",
    "photovoltaic cell",
    "photovoltaic cells",
    "mono solar cell",
    "mono solar cells",
    "mono perc",
    "perc solar cell",
    "silicon cell",
    "silicon cells",
]
POSITIVE_CELL_HINTS = [
    "solar",
    "photovoltaic",
    "mono",
    "perc",
    "topcon",
    "hjt",
    "bifacial",
    "silicon",
    "q.antum",
]
MODULE_EXCLUSION_PHRASES = [
    "solar module",
    "solar modules",
    "pv module",
    "pv modules",
    "assembled in module",
    "assembled in modules",
    "installed in module",
    "installed in modules",
]

HS_CODE_REGEX = re.compile(r"(?<!\d)\d{6}(?!\d)")


def sanitize_file_component(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._")
    return cleaned or "cell_output"


def normalize_for_match(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Z0-9]+", "", str(value).upper())


def compile_boundary_pattern(keyword: str) -> re.Pattern[str]:
    return re.compile(rf"(?i)(?<![A-Z0-9]){re.escape(keyword)}(?![A-Z0-9])")


BOUNDARY_PATTERNS = {keyword: compile_boundary_pattern(keyword) for keyword in STEP3_BOUNDARY_KEYWORDS}


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
        "run_log": output_dir / f"{safe_stem}_cell_processing_run_log.md",
    }


def open_file_dialog(initial_dir: Path) -> Path | None:
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="Select the cell workbook",
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
        "description": "description",
        "marks_numbers": "marks & numbers",
    }

    missing = [display_name for display_name in required.values() if display_name not in headers]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    return {key: headers[value] for key, value in required.items()}


def fuzzy_contains(value: object, patterns: list[str]) -> bool:
    normalized_value = normalize_for_match(value)
    if not normalized_value:
        return False
    for pattern in patterns:
        normalized_pattern = normalize_for_match(pattern)
        if normalized_pattern and normalized_pattern in normalized_value:
            return True
    return False


def keyword_match(value: object, keywords: list[str], boundary_keywords: list[str]) -> bool:
    text = str(value or "")
    lowered = text.lower()
    normalized = normalize_for_match(text)

    for keyword in keywords:
        if keyword.lower() in lowered:
            return True

        normalized_keyword = normalize_for_match(keyword)
        if normalized_keyword and normalized_keyword in normalized:
            return True

    for keyword in boundary_keywords:
        if BOUNDARY_PATTERNS[keyword].search(text):
            return True

    return False


def row_looks_like_cell_import(hs_code: object, description: object, marks_numbers: object) -> bool:
    hs_text = normalize_for_match(hs_code)
    combined_text = " ".join(str(value or "") for value in [description, marks_numbers]).lower()
    has_strong_cell_phrase = any(phrase in combined_text for phrase in POSITIVE_CELL_PHRASES)
    mentions_module_or_panel = "module" in combined_text or "panel" in combined_text
    has_explicit_module_phrase = any(phrase in combined_text for phrase in MODULE_EXCLUSION_PHRASES)

    if has_explicit_module_phrase:
        return False

    if mentions_module_or_panel and not has_strong_cell_phrase and "854142" not in hs_text:
        return False

    if has_strong_cell_phrase:
        return True

    if "cell" in combined_text and any(hint in combined_text for hint in POSITIVE_CELL_HINTS) and not mentions_module_or_panel:
        return True

    if "854142" in hs_text:
        return True

    return False


def build_snapshot_workbook(source_ws, rows_to_keep: list[int], max_column: int) -> Workbook:
    workbook = Workbook()
    target_ws = workbook.active
    target_ws.title = source_ws.title
    target_ws.freeze_panes = source_ws.freeze_panes
    target_ws.sheet_view.showGridLines = source_ws.sheet_view.showGridLines

    for column_index in range(1, max_column + 1):
        column_letter = get_column_letter(column_index)
        source_dimension = source_ws.column_dimensions[column_letter]
        target_dimension = target_ws.column_dimensions[column_letter]
        target_dimension.width = source_dimension.width
        target_dimension.hidden = source_dimension.hidden
        target_dimension.bestFit = source_dimension.bestFit

    all_rows = [1] + rows_to_keep
    for target_row_index, source_row_index in enumerate(all_rows, start=1):
        source_row_dimension = source_ws.row_dimensions[source_row_index]
        target_row_dimension = target_ws.row_dimensions[target_row_index]
        target_row_dimension.height = source_row_dimension.height
        target_row_dimension.hidden = source_row_dimension.hidden

        for column_index in range(1, max_column + 1):
            source_cell = source_ws.cell(row=source_row_index, column=column_index)
            target_cell = target_ws.cell(row=target_row_index, column=column_index, value=source_cell.value)
            if source_cell.has_style:
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


def process_cell_file(source_path: Path, output_path: Path, zip_output_path: Path) -> dict[str, int | str]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    zip_output_path.parent.mkdir(parents=True, exist_ok=True)

    source_workbook = load_workbook(source_path)
    source_ws = source_workbook["DATA"] if "DATA" in source_workbook.sheetnames else source_workbook.active
    required_columns = get_required_columns(source_ws)
    max_output_column = required_columns["marks_numbers"]

    initial_rows = list(range(2, source_ws.max_row + 1))

    step1_rows = []
    for row_index in initial_rows:
        hs_code_value = str(source_ws.cell(row=row_index, column=required_columns["hs_code"]).value or "")
        distinct_codes = set(HS_CODE_REGEX.findall(hs_code_value))
        if len(distinct_codes) <= 4:
            step1_rows.append(row_index)

    step2_rows = []
    for row_index in step1_rows:
        consignee = source_ws.cell(row=row_index, column=required_columns["consignee_declared"]).value
        shipper = source_ws.cell(row=row_index, column=required_columns["shipper_declared"]).value
        if not (fuzzy_contains(consignee, STEP2_NAME_PATTERNS) or fuzzy_contains(shipper, STEP2_NAME_PATTERNS)):
            step2_rows.append(row_index)

    step3_rows = []
    for row_index in step2_rows:
        description = source_ws.cell(row=row_index, column=required_columns["description"]).value
        marks_numbers = source_ws.cell(row=row_index, column=required_columns["marks_numbers"]).value
        shipper = source_ws.cell(row=row_index, column=required_columns["shipper_declared"]).value
        hs_code = source_ws.cell(row=row_index, column=required_columns["hs_code"]).value
        is_cell_candidate = row_looks_like_cell_import(hs_code, description, marks_numbers)

        if not is_cell_candidate or fuzzy_contains(shipper, STEP3_SHIPPER_PATTERNS):
            continue

        step3_rows.append(row_index)

    summary: dict[str, int | str] = {
        "source_file": str(source_path),
        "worksheet": source_ws.title,
        "output_directory": str(output_path.parent),
        "initial_data_rows": len(initial_rows),
        "step1_removed_hs_code_rows": len(initial_rows) - len(step1_rows),
        "rows_after_step1": len(step1_rows),
        "step2_removed_name_rows": len(step1_rows) - len(step2_rows),
        "rows_after_step2": len(step2_rows),
        "step3_removed_keyword_rows": len(step2_rows) - len(step3_rows),
        "final_rows_left": len(step3_rows),
    }

    step_files: list[Path] = []
    for step_key, rows_to_keep in [("step1", step1_rows), ("step2", step2_rows), ("step3", step3_rows)]:
        workbook = build_snapshot_workbook(source_ws, rows_to_keep, max_output_column)
        saved_path = save_workbook(workbook, output_path.parent / STEP_FILE_NAMES[step_key])
        workbook.close()
        step_files.append(saved_path)

    final_workbook = build_snapshot_workbook(source_ws, step3_rows, max_output_column)
    saved_output_path = save_workbook(final_workbook, output_path)
    final_workbook.close()

    saved_zip_path = create_zip_bundle(zip_output_path, step_files + [saved_output_path])

    summary["output_file"] = str(saved_output_path)
    summary["zip_file"] = str(saved_zip_path)
    summary["step_files"] = ", ".join(str(path) for path in step_files)

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
        f"- Step 2 removed (excluded names): {summary['step2_removed_name_rows']}",
        f"- Rows after step 2: {summary['rows_after_step2']}",
        f"- Step 3 removed (excluded keywords): {summary['step3_removed_keyword_rows']}",
        f"- Final rows left: {summary['final_rows_left']}",
    ]
    with log_path.open("a", encoding="utf-8") as file_handle:
        file_handle.write("\n".join(lines) + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Filter a cell workbook and save the processed file, step snapshots, and zip bundle."
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

    summary = process_cell_file(source_path, run_paths["output_file"], run_paths["zip_file"])
    write_run_log(run_paths["run_log"], summary)

    print("Cell processing complete.")
    for key, value in summary.items():
        print(f"{key}: {value}")
    print(f"Run log updated: {run_paths['run_log']}")


if __name__ == "__main__":
    main()
