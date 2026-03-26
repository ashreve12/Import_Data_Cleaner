from __future__ import annotations

import argparse
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


DEFAULT_SOURCE_FILE_NAME = "Wafer_Raw_AI.xlsx"
DEFAULT_OUTPUT_DIR_NAME = "AI Output"
EXCEL_FILE_PATTERNS = "*.xlsx *.xlsm *.xltx *.xltm"

SUPPLIER_PATTERNS = [
    "ESWIN",
    "Shin-Etsu",
    "SEH AMERICA",
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
]

EXCLUDE_TERMS_STEP3 = [
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
]

DIMENSION_TERMS = ["12 inch", "300mm", "8 inch", "200mm"]


def normalize_for_match(value: object) -> str:
    if value is None:
        return ""
    text = str(value).upper()
    return re.sub(r"[^A-Z0-9]", "", text)


def has_any_pattern(value: object, patterns: list[str]) -> bool:
    normalized_cell = normalize_for_match(value)
    if not normalized_cell:
        return False
    for pattern in patterns:
        if normalize_for_match(pattern) in normalized_cell:
            return True
    return False


def is_subsequence(pattern: str, text: str) -> bool:
    iterator = iter(text)
    return all(char in iterator for char in pattern)


def has_any_pattern_subsequence(value: object, patterns: list[str]) -> bool:
    normalized_cell = normalize_for_match(value)
    if not normalized_cell:
        return False
    for pattern in patterns:
        normalized_pattern = normalize_for_match(pattern)
        if normalized_pattern and is_subsequence(normalized_pattern, normalized_cell):
            return True
    return False


def sanitize_file_component(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._")
    return cleaned or "wafer_output"


def delete_rows(ws, rows_to_delete: list[int]) -> None:
    for row_idx in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_idx, 1)


def create_zip_bundle(zip_path: Path, files_to_zip: list[Path]) -> Path:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for file_path in files_to_zip:
            if file_path.exists():
                zip_file.write(file_path, arcname=file_path.name)
    return zip_path


def prepare_output_copy(source_path: Path, preferred_output_path: Path) -> Path:
    try:
        shutil.copy2(source_path, preferred_output_path)
        return preferred_output_path
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = preferred_output_path.with_name(f"{preferred_output_path.stem}_{timestamp}{preferred_output_path.suffix}")
        shutil.copy2(source_path, fallback_path)
        return fallback_path


def create_paths_for_run(source_path: Path, output_dir: Path, zip_name: str | None = None) -> dict[str, Path]:
    safe_stem = sanitize_file_component(source_path.stem)
    zip_stem = sanitize_file_component(Path(zip_name).stem) if zip_name else f"{safe_stem}_step_outputs"

    return {
        "output_dir": output_dir,
        "output_file": output_dir / f"{safe_stem}_processed.xlsx",
        "zip_file": output_dir / f"{zip_stem}.zip",
        "run_log": output_dir / f"{safe_stem}_processing_run_log.md",
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


def process_wafer_file(source_path: Path, output_path: Path, zip_output_path: Path) -> dict:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    zip_output_path.parent.mkdir(parents=True, exist_ok=True)
    working_output_path = prepare_output_copy(source_path, output_path)

    wb = load_workbook(working_output_path)
    ws = wb.active
    output_dir = working_output_path.parent
    base_name = working_output_path.stem

    step_files: list[Path] = []

    def save_step_snapshot(step_label: str) -> None:
        step_path = output_dir / f"{base_name}_{step_label}.xlsx"
        wb.save(step_path)
        step_files.append(step_path)

    summary: dict[str, int | str] = {
        "source_file": str(source_path),
        "output_directory": str(output_path.parent),
        "output_file": str(working_output_path),
        "zip_file": str(zip_output_path),
        "initial_data_rows": max(ws.max_row - 1, 0),
    }

    hs_code_regex = re.compile(r"(?<!\d)\d{6}(?!\d)")

    rows_step1 = []
    for r in range(2, ws.max_row + 1):
        hs_cell = str(ws.cell(row=r, column=3).value or "")
        distinct_codes = set(hs_code_regex.findall(hs_cell))
        if len(distinct_codes) > 4:
            rows_step1.append(r)
    delete_rows(ws, rows_step1)
    save_step_snapshot("step1")
    summary["step1_removed_hs_code_rows"] = len(rows_step1)
    summary["rows_after_step1"] = max(ws.max_row - 1, 0)

    rows_step2a = []
    for r in range(2, ws.max_row + 1):
        consignee = ws.cell(row=r, column=11).value
        shipper = ws.cell(row=r, column=12).value
        if has_any_pattern_subsequence(consignee, SUPPLIER_PATTERNS) or has_any_pattern_subsequence(shipper, SUPPLIER_PATTERNS):
            rows_step2a.append(r)
    delete_rows(ws, rows_step2a)
    save_step_snapshot("step2a")
    summary["step2a_removed_supplier_rows"] = len(rows_step2a)
    summary["rows_after_step2a"] = max(ws.max_row - 1, 0)

    rows_step2b = []
    for r in range(2, ws.max_row + 1):
        description = ws.cell(row=r, column=35).value
        marks_numbers = ws.cell(row=r, column=37).value
        if has_any_pattern(description, DIMENSION_TERMS) or has_any_pattern(marks_numbers, DIMENSION_TERMS):
            rows_step2b.append(r)
    delete_rows(ws, rows_step2b)
    save_step_snapshot("step2b")
    summary["step2b_removed_dimension_rows"] = len(rows_step2b)
    summary["rows_after_step2b"] = max(ws.max_row - 1, 0)

    rows_step3_keep_wafer = []
    for r in range(2, ws.max_row + 1):
        description = ws.cell(row=r, column=35).value
        marks_numbers = ws.cell(row=r, column=37).value
        has_wafer = has_any_pattern(description, ["WAFER"]) or has_any_pattern(marks_numbers, ["WAFER"])
        if not has_wafer:
            rows_step3_keep_wafer.append(r)
    delete_rows(ws, rows_step3_keep_wafer)
    summary["step3_removed_non_wafer_rows"] = len(rows_step3_keep_wafer)
    summary["rows_after_step3_keep_wafer"] = max(ws.max_row - 1, 0)

    rows_step3_exclusions = []
    for r in range(2, ws.max_row + 1):
        description = ws.cell(row=r, column=35).value
        marks_numbers = ws.cell(row=r, column=37).value
        if has_any_pattern(description, EXCLUDE_TERMS_STEP3) or has_any_pattern(marks_numbers, EXCLUDE_TERMS_STEP3):
            rows_step3_exclusions.append(r)
    delete_rows(ws, rows_step3_exclusions)
    save_step_snapshot("step3")
    summary["step3_removed_irrelevant_rows"] = len(rows_step3_exclusions)
    summary["final_rows_left"] = max(ws.max_row - 1, 0)

    wb.save(working_output_path)
    create_zip_bundle(zip_output_path, step_files + [working_output_path])

    summary["step_files"] = ", ".join(str(path) for path in step_files)
    summary["zip_file"] = str(zip_output_path)

    return summary


def write_run_log(log_path: Path, summary: dict) -> None:
    lines = [
        f"\n## {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"- Source: {summary['source_file']}",
        f"- Output directory: {summary['output_directory']}",
        f"- Output: {summary['output_file']}",
        f"- Zip bundle: {summary['zip_file']}",
        f"- Step snapshots: {summary['step_files']}",
        f"- Initial data rows: {summary['initial_data_rows']}",
        f"- Step 1 removed (HS 4+): {summary['step1_removed_hs_code_rows']}",
        f"- Rows after step 1: {summary['rows_after_step1']}",
        f"- Step 2a removed (supplier patterns): {summary['step2a_removed_supplier_rows']}",
        f"- Rows after step 2a: {summary['rows_after_step2a']}",
        f"- Step 2b removed (12 inch/300mm/8 inch/200mm): {summary['step2b_removed_dimension_rows']}",
        f"- Rows after step 2b: {summary['rows_after_step2b']}",
        f"- Step 3 removed (non-wafer rows): {summary['step3_removed_non_wafer_rows']}",
        f"- Rows after wafer keep rule: {summary['rows_after_step3_keep_wafer']}",
        f"- Step 3 removed (irrelevant terms): {summary['step3_removed_irrelevant_rows']}",
        f"- Final rows left: {summary['final_rows_left']}",
    ]
    with log_path.open("a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Filter a wafer workbook and save the processed file, step snapshots, and zip bundle."
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

    summary = process_wafer_file(source_path, run_paths["output_file"], run_paths["zip_file"])
    write_run_log(run_paths["run_log"], summary)

    print("Wafer processing complete.")
    for key, value in summary.items():
        print(f"{key}: {value}")
    print(f"Run log updated: {run_paths['run_log']}")


if __name__ == "__main__":
    main()
