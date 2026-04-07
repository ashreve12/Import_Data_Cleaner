import base64
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, List, Literal, Optional

import requests
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel, ConfigDict, Field


APP_DIR = Path(__file__).parent
WORKSPACE_ROOT = APP_DIR.parents[1]
WAFER_SCRIPT_PATH = APP_DIR / "process_wafer_workbook.py"
CELL_SCRIPT_PATH = WORKSPACE_ROOT / "Cell Cleaner" / "process_cell_workbook.py"
MODULE_SCRIPT_PATH = WORKSPACE_ROOT / "Module Cleaner" / "process_module_workbook.py"
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
    ("step2a_removed_supplier_rows", "Removed at step 2a (supplier patterns)"),
    ("step2b_removed_dimension_rows", "Removed at step 2b (dimension filters)"),
    ("step2_removed_name_rows", "Removed at step 2 (name filters)"),
    ("step2_removed_keyword_rows", "Removed at step 2 (keyword filters)"),
    ("step3_removed_non_wafer_rows", "Removed at step 3 (non-wafer rows)"),
    ("step3_removed_irrelevant_rows", "Removed at step 3 (irrelevant terms)"),
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
