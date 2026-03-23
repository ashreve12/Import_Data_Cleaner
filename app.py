import base64
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List, Optional

import requests
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, ConfigDict, Field


APP_DIR = Path(__file__).parent
SCRIPT_PATH = APP_DIR / "process_wafer_workbook.py"
MAX_FILE_BYTES = 10 * 1024 * 1024  # GPT Action return limit per file
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "https://wafer-cleaner-api.onrender.com").rstrip("/")
CONTACT_EMAIL = os.getenv("CONTACT_EMAIL", "your-email@example.com")

app = FastAPI(
    title="Wafer Cleaner API",
    version="1.0.0",
    description=(
        "Process one uploaded wafer Excel workbook and return the cleaned workbook "
        "plus a zip bundle of the step-by-step snapshots."
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


@app.get("/health", include_in_schema=False)
def health() -> dict[str, bool]:
    return {"ok": True}


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
        <p>This service processes uploaded Excel workbooks and optional user inputs solely to run the wafer cleaner workflow and return output files.</p>
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
    if not SCRIPT_PATH.exists():
        raise HTTPException(status_code=500, detail="process_wafer_workbook.py not found on server.")

    if len(request.openaiFileIdRefs) != 1:
        raise HTTPException(status_code=400, detail="Upload exactly one workbook.")

    uploaded_file = request.openaiFileIdRefs[0]

    if not uploaded_file.name.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise HTTPException(status_code=400, detail="Please upload an Excel workbook.")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        source_path = temp_path / uploaded_file.name
        output_dir = temp_path / "output"
        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            download_file(uploaded_file.download_link, source_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Could not download uploaded file: {exc}") from exc

        cmd = [
            sys.executable,
            str(SCRIPT_PATH),
            "--source-file",
            str(source_path),
            "--output-dir",
            str(output_dir),
            "--no-dialogs",
        ]

        if request.zip_name:
            cmd += ["--zip-name", request.zip_name]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
        )

        if result.returncode != 0:
            raise HTTPException(
                status_code=500,
                detail={
                    "message": "Cleaner script failed.",
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                },
            )

        cleaned_workbook = newest_matching_file(output_dir, "*_processed.xlsx")
        zip_bundle = newest_matching_file(output_dir, "*.zip")

        if not cleaned_workbook:
            raise HTTPException(status_code=500, detail="Cleaner finished but no XLSX output was found.")

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
            stdout=result.stdout[-4000:],
            openaiFileResponse=returned_files,
        )