import base64
import json as _json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, List, Literal, Optional

import requests
from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel, ConfigDict, Field


APP_DIR = Path(__file__).parent
WORKSPACE_ROOT = APP_DIR.parents[1]
WAFER_SCRIPT_PATH = APP_DIR.parent / "process_wafer_workbook.py"
CELL_SCRIPT_PATH = WORKSPACE_ROOT / "Cell Cleaner" / "process_cell_workbook.py"
MODULE_SCRIPT_PATH = WORKSPACE_ROOT / "Module Cleaner" / "process_module_workbook.py"
MERGE_SCRIPT_PATH = APP_DIR.parent / "process_merge_workbooks.py"
HTML_PAGE_PATH = APP_DIR / "static" / "wafer_cleaner.html"
MAX_FILE_BYTES = 10 * 1024 * 1024  # GPT Action return limit per file
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "https://wafer-cleaner-api-production.up.railway.app").rstrip("/")
CONTACT_EMAIL = os.getenv("CONTACT_EMAIL", "your-email@example.com")
BROWSER_JOB_TTL_HOURS = int(os.getenv("BROWSER_JOB_TTL_HOURS", "12"))
BROWSER_JOB_STORAGE = Path(
    os.getenv(
        "BROWSER_JOB_STORAGE",
        str(Path(tempfile.gettempdir()) / "wafer-cleaner-browser-jobs"),
    )
)
EXCEL_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")
JOB_ID_PATTERN = re.compile(r"^[a-f0-9]{32}$")
SUMMARY_FIELD_LABELS = [
    ("initial_data_rows", "Initial data rows"),
    ("step1_removed_hs_code_rows", "Removed at step 1 (HS 4+)"),
    ("step2_removed_supplier_rows", "Removed at step 2 (supplier filter)"),
    ("step2_removed_name_rows", "Removed at step 2 (name filters)"),
    ("step2_removed_keyword_rows", "Removed at step 2 (keyword filters)"),
    ("step3_removed_rows", "Removed at step 3 (non-wafer / exclusion terms)"),
    ("step3_removed_keyword_rows", "Removed at step 3 (keyword filters)"),
    ("final_rows_left", "Final rows left"),
]
BROWSER_CLEANER_OPTIONS = {
    "wafer-imports": {
        "label": "Wafer Imports",
        "available": True,
        "script_path": WAFER_SCRIPT_PATH,
        "success_message": "Wafer cleaner completed successfully.",
    },
    "cell-imports": {
        "label": "Cell Imports",
        "available": True,
        "script_path": CELL_SCRIPT_PATH,
        "success_message": "Cell cleaner completed successfully.",
    },
    "module-imports": {
        "label": "Module Imports",
        "available": True,
        "script_path": MODULE_SCRIPT_PATH,
        "success_message": "Module cleaner completed successfully.",
    },
}
DEFAULT_BROWSER_CLEANER = "wafer-imports"

KEYWORDS_PATH = WORKSPACE_ROOT / "keywords.json"

BUILTIN_KEYWORDS: dict[str, dict[str, dict]] = {
    "wafer": {
        "supplier_patterns": {
            "label": "Supplier / Company Exclusions",
            "help": "Rows whose Shipper Declared or Consignee Declared contains any of these names are removed.",
            "values": [
                "ESWIN", "Shin-Etsu", "SEH AMERICA", "S.E.H. AMERICA",
                "Zing Semiconductor", "SK SILTRON", "silicon valley",
                "Helitek Company", "WAFER WORKS CORPORATION", "A1 SILICON",
                "A-1 SILCON INC.", "SILTRONIC", "WAFERNET", "ECHEM SOLUTIONS",
                "XXX semiconductor", "SOUTHWEST SILICON", "SemiStar Corp",
                "PURE WAFER", "SAMSUNG AUSTIN SEMICONDUCTOR",
                "LUMENTUM OPERATIONS LLC", "Ramco Technology",
                "SCIENTECH CORPORATION", "Formosa Sumco Technology Corporation",
                "West European Silicon Technologies",
                "TEKNICAL MATERIAL RECYCLING", "INC. RAMCO TECHNOLOGY INC",
                "SOITEC", "LUXFER MEL TECHNOLOGIES",
                "WEST EUROPEAN SILICON TECH", "MEMC", "FOODS", "WOLFSPEED",
                "KINIK COMPANY", "FORMOSA SUMCO",
            ],
        },
        "exclude_terms": {
            "label": "Description Exclusion Terms",
            "help": "Rows whose container description contains any of these terms are removed (after the WAFER keep filter).",
            "values": [
                "Semi", "SEMI CONDUCTOR", "CLEANING", "ADDITIVE", "TEXTURING",
                "Coating", "agent", "OVEN", "GALLIUM ARSENIDE", "Mixed",
                "CUTTING", "COOLANT", "MIXED", "RECYCLE", "RECLAIM", "SEED",
                "TAPE", "WAFFY", "COOKIES", "cakes", "peeler",
            ],
        },
    },
    "cell": {
        "name_patterns": {
            "label": "Company Name Exclusions",
            "help": "Rows whose Shipper Declared or Consignee Declared contains any of these names are removed.",
            "values": ["First Solar", "FS Solar", "Space Exploration", "SpaceX"],
        },
        "description_keywords": {
            "label": "Description Keywords",
            "help": "Rows whose container description contains any of these keywords are removed.",
            "values": [
                "accessory", "accessories", "additive", "aluminum", "amorphous",
                "amplifier", "apparatus", "baby", "backsheet", "battery", "bifacial",
                "boxes", "bracket", "blanket", "bug trap", "camera", "carrier",
                "charger", "chip", "circuit", "clay", "cleaning", "coil",
                "conditioner", "connector", "construction", "control", "controller",
                "conduct", "convertor", "cooking", "coupler", "children", "crystal",
                "current", "daughter card", "decoration", "diodes", "doors",
                "electronic products", "encoder", "emitting", "equipment", "filter",
                "film", "flexible", "foil", "fold", "fountain", "frame", "furniture",
                "glass", "hand wash", "heat", "heater", "igniter", "infrared",
                "insulation", "inverter", "igbt", "junction", "jewelry", "kit",
                "laminator", "lamping", "laser", "light", "lithium", "machine",
                "main amp", "material", "men's", "module", "mount", "mounting",
                "neutral bar", "off-grid", "outdoor", "packaging", "panel", "part",
                "plastic", "portable", "polyester", "power", "piezo", "photocell",
                "production line", "photo cell", "rectifier", "rubber", "scissor",
                "seal", "sealing", "semiconductor", "sensor", "shaver", "simulator",
                "smart", "solar light", "steel", "supply", "system", "thyristor",
                "tool", "tracking", "transducer", "transformer", "transistor",
                "tshirt", "t-shirt", "tube", "underwear", "wire",
            ],
        },
        "boundary_keywords": {
            "label": "Boundary Keywords (whole-word match)",
            "help": "Removed when found as whole words in the container description (e.g. LED won't match FLED).",
            "values": ["LED", "EVA", "USB"],
        },
    },
    "module": {
        "company_exclusions": {
            "label": "Company Exclusions",
            "help": "Rows whose Shipper Declared or Consignee Declared contains any of these names are removed.",
            "values": ["extrusion", "SEOUL", "VON ARDENNE", "FHR ANLAGENBAU"],
        },
        "description_keywords": {
            "label": "Description Keywords",
            "help": "Rows whose container description contains any of these keywords are removed.",
            "values": [
                "accessory", "accessories", "additive", "aluminum", "amorphous",
                "amplifier", "apparatus", "baby", "backsheet", "battery", "boxes",
                "bracket", "branket", "bug trap", "camera", "carrier", "charger",
                "chip", "circuit", "clay", "cleaning", "coil", "conditioner",
                "connector", "construction", "control", "controller", "conduct",
                "convertor", "cooking", "coupler", "children", "crystal", "current",
                "daughter card", "decoration", "diodes", "doors", "electronic products",
                "encoder", "emitting", "equipment", "filter", "film", "flexible",
                "foil", "fold", "for solar", "fountain", "frame", "furniture", "glass",
                "hand wash", "heat", "heater", "igniter", "infrared", "insulation",
                "inverter", "igbt", "junction", "jewelry", "kit", "laminator",
                "lamping", "laser", "light", "lithium", "machine", "main amp",
                "material", "men's", "mount", "mounting", "neutral bar", "off-grid",
                "outdoor", "packaging", "part", "plastic", "portable", "polyester",
                "power", "piezo", "photocell", "photo cell", "production line",
                "rectifier", "rubber", "ribbon", "scissor", "seal", "sealing",
                "semiconductor", "sensor", "shaver", "simulator", "smart",
                "solar light", "sport", "steel", "supply", "system", "thyristor",
                "tool", "toy", "tracking", "transducer", "transformer", "transistor",
                "tshirt", "t-shirt", "tube", "underwear", "wire",
            ],
        },
        "boundary_exclusions": {
            "label": "Boundary Exclusions (whole-word match)",
            "help": "Removed when found as whole words in the container description.",
            "values": ["LED", "EVA", "USB"],
        },
    },
}


app = FastAPI(
    title="Import Data Cleaner API",
    version="1.2.0",
    description=(
        "Process uploaded import-data Excel workbooks and return the cleaned workbook "
        "plus a zip bundle of step-by-step snapshots. The same deployment serves "
        "both the ChatGPT Action flow and the browser upload flow."
    ),
    servers=[{"url": PUBLIC_BASE_URL, "description": "Public API base URL"}],
)


class OpenAIFileRef(BaseModel):
    model_config = ConfigDict(extra="allow")

    name: str
    id: str
    mime_type: str
    download_link: str


class RunRequest(BaseModel):
    openaiFileIdRefs: List[OpenAIFileRef] = Field(
        ...,
        description="Exactly one Excel workbook uploaded by the user.",
    )
    zip_name: Optional[str] = Field(
        default=None,
        description="Optional zip filename to use for the output bundle.",
    )


class OpenAIFileResponseItem(BaseModel):
    name: str
    mime_type: str
    content: str


class RunResponse(BaseModel):
    status: str
    message: str
    stdout: str
    openaiFileResponse: List[OpenAIFileResponseItem]


class BrowserSummaryItem(BaseModel):
    label: str
    value: str


class BrowserDownloadItem(BaseModel):
    label: str
    name: str
    url: str


class BrowserRunResponse(BaseModel):
    status: str
    message: str
    job_id: str
    expires_at: str
    summary: List[BrowserSummaryItem]
    downloads: List[BrowserDownloadItem]


class MergeSummaryItem(BaseModel):
    label: str
    value: str


class KeywordBody(BaseModel):
    keyword: str


class MergeRunResponse(BaseModel):
    status: str
    message: str
    job_id: str
    expires_at: str
    summary: List[MergeSummaryItem]
    downloads: List[BrowserDownloadItem]


def _read_custom_keywords() -> dict:
    if not KEYWORDS_PATH.exists():
        return {}
    try:
        return _json.loads(KEYWORDS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write_custom_keywords(data: dict) -> None:
    KEYWORDS_PATH.write_text(_json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


# In-memory store for background merge jobs: job_id → status dict
_merge_jobs: dict[str, dict] = {}
_merge_jobs_lock = threading.Lock()


def _run_merge_job(
    job_id: str,
    source_paths: list,
    dest_in_path: Path,
    output_path: Path,
    source_dir: Path,
    dest_sheet_name: str | None,
    source_names: list,
) -> None:
    import importlib.util as _ilu
    _spec = _ilu.spec_from_file_location("process_merge_workbooks", MERGE_SCRIPT_PATH)
    _mod = _ilu.module_from_spec(_spec)
    _spec.loader.exec_module(_mod)

    try:
        # Merge all sources into dest in a single pass — dest loaded once, sources streamed
        result = _mod.merge_all_workbooks(
            source_paths=source_paths,
            dest_path=dest_in_path,
            output_path=output_path,
            dest_sheet_name=dest_sheet_name or None,
        )

        expires_at = (
            datetime.now(timezone.utc) + timedelta(hours=BROWSER_JOB_TTL_HOURS)
        ).isoformat()

        total_rows = result["rows_added"]
        summary = [
            {"label": "Files merged",             "value": str(len(source_paths))},
            {"label": "Total rows added",         "value": str(total_rows)},
            {"label": "Destination sheet",        "value": result.get("destination_sheet", "")},
            {"label": "Destination table",        "value": result.get("destination_table", "none")},
            {"label": "Matched columns",          "value": str(result.get("matched_columns", ""))},
            {"label": "Formula columns adjusted", "value": str(result.get("formula_columns_adjusted", ""))},
            {"label": "Source files",             "value": ", ".join(source_names)},
        ]

        with _merge_jobs_lock:
            _merge_jobs[job_id] = {
                "status": "complete",
                "message": (
                    f"Merge complete. {total_rows} row(s) added across "
                    f"{len(source_paths)} file(s) into "
                    f"'{result.get('destination_sheet', 'the destination sheet')}'."
                ),
                "summary": summary,
                "download_url": f"/download-merge-result/{job_id}",
                "download_name": output_path.name,
                "expires_at": expires_at,
            }
    except Exception as exc:
        with _merge_jobs_lock:
            _merge_jobs[job_id] = {"status": "error", "message": str(exc)}


@app.get("/health", include_in_schema=False)
def health() -> dict[str, bool]:
    return {"ok": True}


@app.get("/", include_in_schema=False, response_class=HTMLResponse)
@app.get("/wafer-cleaner", include_in_schema=False, response_class=HTMLResponse)
def wafer_cleaner_page() -> str:
    if not HTML_PAGE_PATH.exists():
        raise HTTPException(status_code=500, detail="wafer_cleaner.html not found on server.")
    return HTML_PAGE_PATH.read_text(encoding="utf-8")


@app.get("/privacy", include_in_schema=False, response_class=HTMLResponse)
def privacy_policy() -> str:
    return f"""
    <html>
      <head>
        <title>Privacy Policy</title>
        <meta charset=\"utf-8\" />
      </head>
      <body>
        <h1>Privacy Policy</h1>
        <p>This service processes uploaded Excel workbooks and optional user inputs solely to run the selected import cleaner workflow and return output files.</p>
        <p>Data processed may include uploaded workbook contents, filenames, optional zip names, and standard server logs.</p>
        <p>Files are used only to complete the requested processing workflow and are not sold.</p>
        <p>Data may be handled by infrastructure providers required to operate the service, including the hosting platform used to run this API.</p>
        <p>Temporary files are created only for processing and should not be retained after the request finishes, except as needed for routine logs and platform operations.</p>
        <p>Contact: {CONTACT_EMAIL}</p>
      </body>
    </html>
    """


def download_file(url: str, destination: Path) -> None:
    response = requests.get(url, timeout=120)
    response.raise_for_status()
    destination.write_bytes(response.content)


def newest_matching_file(folder: Path, glob_pattern: str) -> Optional[Path]:
    matches = list(folder.rglob(glob_pattern))
    if not matches:
        return None
    return max(matches, key=lambda path: path.stat().st_mtime)


def as_openai_file(path: Path, mime_type: str) -> OpenAIFileResponseItem:
    size = path.stat().st_size
    if size > MAX_FILE_BYTES:
        raise HTTPException(
            status_code=500,
            detail=f"{path.name} is larger than 10 MB. Use URL-based file return for larger files.",
        )

    return OpenAIFileResponseItem(
        name=path.name,
        mime_type=mime_type,
        content=base64.b64encode(path.read_bytes()).decode("utf-8"),
    )


def ensure_cleaner_script(script_path: Path) -> None:
    if not script_path.exists():
        raise HTTPException(status_code=500, detail=f"{script_path.name} not found on server.")


def ensure_browser_job_root() -> Path:
    BROWSER_JOB_STORAGE.mkdir(parents=True, exist_ok=True)
    return BROWSER_JOB_STORAGE


def normalize_zip_name(zip_name: Optional[str]) -> Optional[str]:
    if zip_name is None:
        return None
    cleaned = zip_name.strip()
    return cleaned or None


def is_excel_filename(filename: str) -> bool:
    return Path(filename).suffix.lower() in EXCEL_SUFFIXES


def validate_browser_cleaner_type(cleaner_type: Optional[str]) -> str:
    normalized = (cleaner_type or DEFAULT_BROWSER_CLEANER).strip().lower()
    option = BROWSER_CLEANER_OPTIONS.get(normalized)

    if not option:
        raise HTTPException(status_code=400, detail="Please choose a valid import cleaner option.")

    if not option["available"]:
        raise HTTPException(
            status_code=400,
            detail=f"{option['label']} cleaning is not available yet. Please choose Wafer Imports or Cell Imports for now.",
        )

    return normalized


def resolve_artifact_path(raw_value: Any) -> Optional[Path]:
    if not raw_value:
        return None

    candidate = Path(str(raw_value))
    if candidate.exists() and candidate.is_file():
        return candidate

    return None


def resolve_cleaned_workbook(output_dir: Path, parsed_summary: Optional[dict[str, Any]] = None) -> Optional[Path]:
    summary = parsed_summary or {}
    cleaned_workbook = resolve_artifact_path(summary.get("output_file"))
    if cleaned_workbook:
        return cleaned_workbook

    workbook_candidates = [
        path
        for path in output_dir.rglob("*")
        if path.is_file() and path.suffix.lower() in EXCEL_SUFFIXES and "step" not in path.stem.lower()
    ]
    if not workbook_candidates:
        workbook_candidates = [
            path for path in output_dir.rglob("*") if path.is_file() and path.suffix.lower() in EXCEL_SUFFIXES
        ]

    return max(workbook_candidates, key=lambda path: path.stat().st_mtime) if workbook_candidates else None


def cleanup_expired_browser_jobs() -> None:
    job_root = ensure_browser_job_root()
    cutoff = datetime.now(timezone.utc) - timedelta(hours=BROWSER_JOB_TTL_HOURS)

    for child in job_root.iterdir():
        try:
            last_modified = datetime.fromtimestamp(child.stat().st_mtime, tz=timezone.utc)
        except FileNotFoundError:
            continue

        if child.is_dir() and last_modified < cutoff:
            shutil.rmtree(child, ignore_errors=True)


def parse_summary_value(raw_value: str) -> Any:
    value = raw_value.strip()
    if re.fullmatch(r"-?\d+", value):
        return int(value)
    return value


def parse_cleaner_stdout(stdout: str) -> dict[str, Any]:
    parsed: dict[str, Any] = {}

    for raw_line in stdout.splitlines():
        line = raw_line.strip()
        if ": " not in line:
            continue

        key, value = line.split(": ", 1)
        normalized_key = key.strip()
        if normalized_key == "Run log updated":
            parsed["run_log"] = value.strip()
            continue

        parsed[normalized_key] = parse_summary_value(value)

    return parsed


def build_browser_summary(parsed_summary: dict[str, Any]) -> List[BrowserSummaryItem]:
    summary_items: List[BrowserSummaryItem] = []

    for key, label in SUMMARY_FIELD_LABELS:
        if key in parsed_summary:
            summary_items.append(BrowserSummaryItem(label=label, value=str(parsed_summary[key])))

    if not summary_items:
        summary_items.append(BrowserSummaryItem(label="Status", value="Processing completed successfully."))

    return summary_items


def run_cleaner(source_path: Path, output_dir: Path, zip_name: Optional[str], script_path: Path) -> dict[str, Any]:
    ensure_cleaner_script(script_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        sys.executable,
        str(script_path),
        "--source-file",
        str(source_path),
        "--output-dir",
        str(output_dir),
        "--no-dialogs",
    ]

    normalized_zip_name = normalize_zip_name(zip_name)
    if normalized_zip_name:
        cmd += ["--zip-name", normalized_zip_name]

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
    )

    parsed_summary = parse_cleaner_stdout(result.stdout)

    if result.returncode != 0:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "Cleaner script failed.",
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
        )

    cleaned_workbook = resolve_cleaned_workbook(output_dir, parsed_summary)
    zip_bundle = resolve_artifact_path(parsed_summary.get("zip_file")) or newest_matching_file(output_dir, "*.zip")

    if not cleaned_workbook:
        raise HTTPException(status_code=500, detail="Cleaner finished but no XLSX output was found.")

    return {
        "stdout": result.stdout,
        "parsed_summary": parsed_summary,
        "summary_items": build_browser_summary(parsed_summary),
        "cleaned_workbook": cleaned_workbook,
        "zip_bundle": zip_bundle,
    }


def build_browser_download_url(job_id: str, artifact: Literal["cleaned-workbook", "zip-bundle"]) -> str:
    return f"/download-wafer-cleaner/{job_id}/{artifact}"


def resolve_browser_job_dir(job_id: str) -> Path:
    if not JOB_ID_PATTERN.fullmatch(job_id):
        raise HTTPException(status_code=404, detail="Processing job not found.")

    job_dir = ensure_browser_job_root() / job_id
    if not job_dir.exists() or not job_dir.is_dir():
        raise HTTPException(status_code=404, detail="Processing job not found.")

    return job_dir


@app.post(
    "/run-wafer-cleaner",
    operation_id="runWaferCleaner",
    summary="Clean a wafer workbook and return output files.",
    description=(
        "Clean one uploaded wafer workbook and return the processed Excel file "
        "plus a zip bundle. The request must include exactly one workbook in "
        "openaiFileIdRefs."
    ),
    response_model=RunResponse,
)
def run_wafer_cleaner(request: RunRequest) -> RunResponse:
    # Original ChatGPT Action flow: keep returning files inline as base64 payloads.
    if len(request.openaiFileIdRefs) != 1:
        raise HTTPException(status_code=400, detail="Upload exactly one workbook.")

    uploaded_file = request.openaiFileIdRefs[0]
    if not is_excel_filename(uploaded_file.name):
        raise HTTPException(status_code=400, detail="Please upload an Excel workbook.")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        source_path = temp_path / Path(uploaded_file.name).name
        output_dir = temp_path / "output"

        try:
            download_file(uploaded_file.download_link, source_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Could not download uploaded file: {exc}") from exc

        run_result = run_cleaner(source_path, output_dir, request.zip_name, WAFER_SCRIPT_PATH)
        cleaned_workbook = run_result["cleaned_workbook"]
        zip_bundle = run_result["zip_bundle"]

        returned_files = [
            as_openai_file(
                cleaned_workbook,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        ]

        if zip_bundle:
            returned_files.append(as_openai_file(zip_bundle, "application/zip"))

        return RunResponse(
            status="success",
            message="Wafer cleaner completed successfully.",
            stdout=run_result["stdout"][-4000:],
            openaiFileResponse=returned_files,
        )


@app.post(
    "/upload-wafer-cleaner",
    summary="Upload an import workbook from a browser and get download links.",
    description=(
        "Browser-friendly upload endpoint that accepts one Excel workbook via "
        "multipart/form-data, runs the selected cleaner script, and returns "
        "summary data plus download links for the processed outputs."
    ),
    response_model=BrowserRunResponse,
)
def upload_wafer_cleaner(
    workbook: UploadFile = File(..., description="Exactly one Excel workbook uploaded from a browser."),
    zip_name: Optional[str] = Form(default=None, description="Optional zip filename for the output bundle."),
    cleaner_type: str = Form(
        default=DEFAULT_BROWSER_CLEANER,
        description="Import data cleaner selected in the browser upload form.",
    ),
) -> BrowserRunResponse:
    # Added browser flow: accept multipart uploads and return download URLs instead of base64 file blobs.
    cleanup_expired_browser_jobs()
    selected_cleaner = validate_browser_cleaner_type(cleaner_type)
    cleaner_config = BROWSER_CLEANER_OPTIONS[selected_cleaner]

    if not workbook.filename:
        raise HTTPException(status_code=400, detail="Please choose one Excel workbook to upload.")

    original_name = Path(workbook.filename).name
    if not is_excel_filename(original_name):
        raise HTTPException(status_code=400, detail="Please upload an Excel workbook.")

    job_id = uuid.uuid4().hex
    job_dir = ensure_browser_job_root() / job_id
    source_dir = job_dir / "source"
    output_dir = job_dir / "output"
    source_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_path = source_dir / original_name

    try:
        with source_path.open("wb") as destination:
            shutil.copyfileobj(workbook.file, destination)

        run_result = run_cleaner(source_path, output_dir, zip_name, cleaner_config["script_path"])
    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Could not process uploaded workbook: {exc}") from exc
    finally:
        workbook.file.close()

    downloads = [
        BrowserDownloadItem(
            label="Download cleaned workbook",
            name=run_result["cleaned_workbook"].name,
            url=build_browser_download_url(job_id, "cleaned-workbook"),
        )
    ]

    if run_result["zip_bundle"]:
        downloads.append(
            BrowserDownloadItem(
                label="Download zip bundle",
                name=run_result["zip_bundle"].name,
                url=build_browser_download_url(job_id, "zip-bundle"),
            )
        )

    expires_at = (datetime.now(timezone.utc) + timedelta(hours=BROWSER_JOB_TTL_HOURS)).isoformat()

    return BrowserRunResponse(
        status="success",
        message=cleaner_config["success_message"],
        job_id=job_id,
        expires_at=expires_at,
        summary=run_result["summary_items"],
        downloads=downloads,
    )


@app.get("/download-wafer-cleaner/{job_id}/{artifact}", include_in_schema=False)
def download_wafer_cleaner_artifact(
    job_id: str,
    artifact: Literal["cleaned-workbook", "zip-bundle"],
) -> FileResponse:
    cleanup_expired_browser_jobs()
    job_dir = resolve_browser_job_dir(job_id)
    output_dir = job_dir / "output"

    if artifact == "cleaned-workbook":
        file_path = resolve_cleaned_workbook(output_dir)
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        file_path = newest_matching_file(output_dir, "*.zip")
        media_type = "application/zip"

    if not file_path:
        raise HTTPException(status_code=404, detail="Requested output file is no longer available.")

    return FileResponse(
        path=file_path,
        media_type=media_type,
        filename=file_path.name,
    )


# ---------------------------------------------------------------------------
# Merge workbooks endpoint
# ---------------------------------------------------------------------------

MERGE_SUMMARY_FIELD_LABELS = [
    ("rows_added", "Rows added"),
    ("matched_columns", "Matched columns"),
    ("formula_columns_adjusted", "Formula columns adjusted"),
    ("source_sheet", "Source sheet"),
    ("destination_sheet", "Destination sheet"),
    ("destination_table", "Destination table"),
]


@app.post("/upload-merge-workbooks", summary="Merge up to 3 cleaned workbooks into a destination master workbook.")
def upload_merge_workbooks(
    background_tasks: BackgroundTasks,
    source_workbooks: List[UploadFile] = File(..., description="Up to 3 cleaned output workbooks."),
    dest_workbook: UploadFile = File(..., description="Master destination workbook to merge into."),
    dest_sheet_name: Optional[str] = Form(default=None, description="Sheet name in the destination workbook (leave blank to use the active sheet)."),
) -> dict:
    cleanup_expired_browser_jobs()

    if not source_workbooks:
        raise HTTPException(status_code=400, detail="At least one source workbook is required.")
    if len(source_workbooks) > 3:
        raise HTTPException(status_code=400, detail="Upload at most 3 source workbooks at a time.")
    for upload in source_workbooks:
        if not upload.filename or not is_excel_filename(upload.filename):
            raise HTTPException(status_code=400, detail=f"'{upload.filename}' is not a supported Excel workbook.")
    if not dest_workbook.filename or not is_excel_filename(dest_workbook.filename):
        raise HTTPException(status_code=400, detail="Destination file must be an Excel workbook.")

    job_id = uuid.uuid4().hex
    job_dir = ensure_browser_job_root() / job_id
    source_dir = job_dir / "source"
    output_dir = job_dir / "output"
    source_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = output_dir / Path(dest_workbook.filename).name

    try:
        dest_in_path = source_dir / Path(dest_workbook.filename).name
        with dest_in_path.open("wb") as fh:
            shutil.copyfileobj(dest_workbook.file, fh)

        source_paths: list[Path] = []
        source_names: list[str] = []
        for upload in source_workbooks:
            p = source_dir / Path(upload.filename).name
            if p in source_paths:
                p = source_dir / f"{p.stem}_{len(source_paths)}{p.suffix}"
            with p.open("wb") as fh:
                shutil.copyfileobj(upload.file, fh)
            source_paths.append(p)
            source_names.append(Path(upload.filename).name)
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Failed to save uploaded files: {exc}") from exc
    finally:
        dest_workbook.file.close()
        for upload in source_workbooks:
            upload.file.close()

    with _merge_jobs_lock:
        _merge_jobs[job_id] = {"status": "processing"}

    background_tasks.add_task(
        _run_merge_job,
        job_id, source_paths, dest_in_path, output_path, source_dir,
        dest_sheet_name, source_names,
    )

    return {"status": "processing", "job_id": job_id}


@app.get("/merge-status/{job_id}", include_in_schema=False)
def merge_status(job_id: str) -> dict:
    if not JOB_ID_PATTERN.match(job_id):
        raise HTTPException(status_code=400, detail="Invalid job ID.")
    with _merge_jobs_lock:
        job = _merge_jobs.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    return job


@app.get("/download-merge-result/{job_id}", include_in_schema=False)
def download_merge_result(job_id: str) -> FileResponse:
    cleanup_expired_browser_jobs()
    job_dir = resolve_browser_job_dir(job_id)
    output_dir = job_dir / "output"

    candidates = [p for p in output_dir.iterdir() if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES]
    if not candidates:
        raise HTTPException(status_code=404, detail="Merged workbook is no longer available.")

    file_path = max(candidates, key=lambda p: p.stat().st_mtime)
    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=file_path.name,
    )


# ---------------------------------------------------------------------------
# Keyword editor endpoints
# ---------------------------------------------------------------------------

@app.get("/keywords/{cleaner_type}", include_in_schema=False)
def get_keywords(cleaner_type: str) -> dict:
    if cleaner_type not in BUILTIN_KEYWORDS:
        raise HTTPException(status_code=404, detail=f"Unknown cleaner type: {cleaner_type}")
    custom_data = _read_custom_keywords()
    cleaner_custom = custom_data.get(cleaner_type, {})
    lists = []
    for list_name, meta in BUILTIN_KEYWORDS[cleaner_type].items():
        entry = cleaner_custom.get(list_name, {})
        if isinstance(entry, list):
            custom_kws, removed_kws = entry, []
        else:
            custom_kws = entry.get("custom", [])
            removed_kws = entry.get("removed", [])
        lists.append({
            "name": list_name,
            "label": meta["label"],
            "help": meta["help"],
            "builtin": meta["values"],
            "custom": custom_kws,
            "removed": removed_kws,
        })
    return {"cleaner_type": cleaner_type, "lists": lists}


@app.post("/keywords/{cleaner_type}/{list_name}", include_in_schema=False)
def add_keyword(cleaner_type: str, list_name: str, body: KeywordBody) -> dict:
    if cleaner_type not in BUILTIN_KEYWORDS:
        raise HTTPException(status_code=404, detail=f"Unknown cleaner type: {cleaner_type}")
    if list_name not in BUILTIN_KEYWORDS[cleaner_type]:
        raise HTTPException(status_code=404, detail=f"Unknown list: {list_name}")
    keyword = body.keyword.strip()
    if not keyword:
        raise HTTPException(status_code=400, detail="keyword must not be empty")
    custom_data = _read_custom_keywords()
    entry = custom_data.setdefault(cleaner_type, {}).setdefault(list_name, {})
    if isinstance(entry, list):
        entry = {"custom": entry, "removed": []}
        custom_data[cleaner_type][list_name] = entry
    kw_list = entry.setdefault("custom", [])
    if keyword not in kw_list:
        kw_list.append(keyword)
        _write_custom_keywords(custom_data)
    return {"ok": True, "keyword": keyword}


@app.delete("/keywords/{cleaner_type}/{list_name}", include_in_schema=False)
def delete_keyword(cleaner_type: str, list_name: str, keyword: str) -> dict:
    """Remove a custom keyword or hide a built-in keyword."""
    if cleaner_type not in BUILTIN_KEYWORDS:
        raise HTTPException(status_code=404, detail=f"Unknown cleaner type: {cleaner_type}")
    builtin_values = BUILTIN_KEYWORDS[cleaner_type].get(list_name, {}).get("values", [])
    custom_data = _read_custom_keywords()
    entry = custom_data.setdefault(cleaner_type, {}).setdefault(list_name, {})
    if isinstance(entry, list):
        entry = {"custom": entry, "removed": []}
        custom_data[cleaner_type][list_name] = entry
    if keyword in builtin_values:
        removed = entry.setdefault("removed", [])
        if keyword not in removed:
            removed.append(keyword)
            _write_custom_keywords(custom_data)
    else:
        kw_list = entry.setdefault("custom", [])
        if keyword in kw_list:
            kw_list.remove(keyword)
            _write_custom_keywords(custom_data)
    return {"ok": True}


@app.post("/keywords/{cleaner_type}/{list_name}/restore", include_in_schema=False)
def restore_keyword(cleaner_type: str, list_name: str, body: KeywordBody) -> dict:
    """Restore a previously removed built-in keyword."""
    if cleaner_type not in BUILTIN_KEYWORDS:
        raise HTTPException(status_code=404, detail=f"Unknown cleaner type: {cleaner_type}")
    keyword = body.keyword.strip()
    custom_data = _read_custom_keywords()
    entry = custom_data.get(cleaner_type, {}).get(list_name, {})
    if isinstance(entry, dict):
        removed = entry.get("removed", [])
        if keyword in removed:
            removed.remove(keyword)
            _write_custom_keywords(custom_data)
    return {"ok": True}
