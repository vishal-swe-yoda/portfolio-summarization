"""
FastAPI application for Institutional Portfolio PPT Generator.
Provides an interactive Swagger UI (OpenAPI) and REST endpoints to ingest
Excel workbooks, allocate sequential filenames, execute the presentation
generation pipeline, and return or download matching PowerPoint decks.

Interactive Swagger UI:
- http://127.0.0.1:8000/docs
- http://127.0.0.1:8000/redoc
"""

from datetime import datetime
import os
import re
import logging
import threading
from typing import List, Optional
from pydantic import BaseModel, Field
from fastapi import FastAPI, File, UploadFile, status
from fastapi.responses import JSONResponse, FileResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

from main import run_pipeline

# Configure application logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("portfolio_api")

# Directory configurations
DATA_DIR = "data"
OUTPUT_DIR = "output"

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Lock to serialize sequence allocation and pipeline execution safely across concurrent threads
pipeline_lock = threading.Lock()


# =============================================================================
# Pydantic Schemas for Interactive Swagger UI Documentation
# =============================================================================

class UploadSuccessResponse(BaseModel):
    status: str = Field("success", description="Status of the operation", example="success")
    input_file: str = Field(..., description="Sequential input Excel filename saved to data/", example="sample_input_1.xlsx")
    output_file: str = Field(..., description="Generated PowerPoint presentation filename in output/", example="output_1.pptx")
    download_url: str = Field(..., description="Direct endpoint URL to download the generated presentation", example="/download/output_1.pptx")
    message: str = Field("File processed successfully", description="User-facing execution status message", example="File processed successfully")


class ErrorResponse(BaseModel):
    status: str = Field("error", description="Error status indicator", example="error")
    message: str = Field(..., description="Descriptive error explanation without sensitive details", example="Invalid file format. Only .xlsx and .xls files are supported.")


class PresentationItem(BaseModel):
    filename: str = Field(..., description="Output PowerPoint presentation filename", example="output_1.pptx")
    size_bytes: int = Field(..., description="File size in bytes", example=125430)
    modified_time: str = Field(..., description="ISO 8601 modification timestamp", example="2026-10-19T14:15:00")
    download_url: str = Field(..., description="Direct download URL", example="/download/output_1.pptx")


class HealthResponse(BaseModel):
    status: str = Field("healthy", example="healthy")
    service: str = Field("Portfolio PPT Generator API", example="Portfolio PPT Generator API")
    version: str = Field("1.0.0", example="1.0.0")
    docs_url: str = Field("/docs", example="/docs")


# =============================================================================
# FastAPI Application with Interactive Swagger UI & OpenAPI Specification
# =============================================================================

app = FastAPI(
    title="Portfolio PPT Generator - Interactive API",
    description="""
## Institutional Portfolio PowerPoint Presentation Generator

Automated, configuration-driven REST API that transforms institutional portfolio Excel workbooks into executive 16:9 widescreen PowerPoint (`.pptx`) decks.

### Key Capabilities
* **Interactive Swagger Upload**: Upload `.xlsx` or `.xls` workbooks directly through Swagger UI (`/docs`).
* **Sequential Numbering & Input-Output Identity**: Monotonically allocates sequence numbers (`sample_input_{N}.xlsx` → `output_{N}.pptx`).
* **Deterministic Analytics**: Analyzes sector exposures, active weights, and concentration ratios.
* **Qualitative AI Synthesis**: Produces evidence-based observations with Google Gemini, DeepSeek, or automatic offline fallback.
* **Instant Presentation Download**: Download generated presentations directly via `/download/{filename}`.
    """,
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json",
    swagger_ui_parameters={
        "defaultModelsExpandDepth": 1,
        "docExpansion": "list",
        "persistAuthorization": True,
    }
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def allocate_next_sequence_number(input_dir: str, output_dir: str) -> int:
    """
    Finds the next unused monotonically increasing sequence number by inspecting
    both input (sample_input_{N}.*) and output (output_{N}.pptx) directories.
    Must be called inside a thread lock.
    """
    existing_numbers = set()

    if os.path.exists(input_dir):
        for fname in os.listdir(input_dir):
            match = re.match(r"^sample_input_(\d+)\.(xlsx|xls)$", fname, re.IGNORECASE)
            if match:
                existing_numbers.add(int(match.group(1)))

    if os.path.exists(output_dir):
        for fname in os.listdir(output_dir):
            match = re.match(r"^output_(\d+)\.pptx$", fname, re.IGNORECASE)
            if match:
                existing_numbers.add(int(match.group(1)))

    return max(existing_numbers, default=0) + 1


# =============================================================================
# Interactive API Endpoints
# =============================================================================

@app.get(
    "/",
    include_in_schema=False,
    summary="Root Redirect to Swagger UI"
)
def root():
    """Redirects visitors to the interactive Swagger UI documentation."""
    return RedirectResponse(url="/docs")


@app.get(
    "/health",
    response_model=HealthResponse,
    tags=["System"],
    summary="Health Check",
    description="Returns the health status and metadata for the generator service."
)
def health():
    return {
        "status": "healthy",
        "service": "Portfolio PPT Generator API",
        "version": "1.0.0",
        "docs_url": "/docs"
    }


@app.post(
    "/upload",
    response_model=UploadSuccessResponse,
    status_code=status.HTTP_200_OK,
    tags=["Presentation Generator"],
    summary="Upload Excel Workbook & Generate Presentation",
    description="""
Upload an Excel portfolio workbook (`.xlsx` or `.xls`) to automatically run the 5-stage PowerPoint generation pipeline.

### Processing Steps:
1. **Validation**: Confirms file extension is `.xlsx` or `.xls` and non-empty.
2. **Sequential Saving**: Saves the file as `data/sample_input_{N}.xlsx` using thread-safe locking.
3. **Pipeline Execution**: Deterministically processes holdings, runs LLM commentary, renders sector chart, and builds slide deck.
4. **Output Alignment**: Saves matching PowerPoint presentation to `output/output_{N}.pptx`.
    """,
    responses={
        200: {
            "model": UploadSuccessResponse,
            "description": "Presentation generated successfully."
        },
        400: {
            "model": ErrorResponse,
            "description": "Invalid file extension or empty file."
        },
        500: {
            "model": ErrorResponse,
            "description": "Pipeline execution or file saving failure."
        }
    }
)
def upload_excel(
    file: UploadFile = File(..., description="Excel workbook containing portfolio holdings (.xlsx or .xls)")
):
    original_filename = file.filename or ""
    _, ext = os.path.splitext(original_filename)
    ext = ext.lower()

    # 1. Validate File Extension
    if ext not in [".xlsx", ".xls"]:
        logger.warning(f"Rejected upload with invalid extension: '{original_filename}'")
        return JSONResponse(
            status_code=status.HTTP_400_BAD_REQUEST,
            content={
                "status": "error",
                "message": "Invalid file format. Only .xlsx and .xls files are supported."
            }
        )

    # 2. Read and Validate File Content
    try:
        content = file.file.read()
    except Exception as e:
        logger.error(f"Failed to read uploaded file '{original_filename}': {e}", exc_info=True)
        return JSONResponse(
            status_code=status.HTTP_400_BAD_REQUEST,
            content={
                "status": "error",
                "message": "Failed to read uploaded file."
            }
        )

    if not content or len(content) == 0:
        logger.warning(f"Rejected empty file upload: '{original_filename}'")
        return JSONResponse(
            status_code=status.HTTP_400_BAD_REQUEST,
            content={
                "status": "error",
                "message": "Uploaded file is empty."
            }
        )

    # 3. Synchronous Thread-Safe Sequence Allocation, File Saving, and Pipeline Execution
    with pipeline_lock:
        seq_num = allocate_next_sequence_number(DATA_DIR, OUTPUT_DIR)
        saved_input_filename = f"sample_input_{seq_num}{ext}"
        saved_input_path = os.path.join(DATA_DIR, saved_input_filename)

        # Extra collision guard
        while os.path.exists(saved_input_path):
            seq_num += 1
            saved_input_filename = f"sample_input_{seq_num}{ext}"
            saved_input_path = os.path.join(DATA_DIR, saved_input_filename)

        # Save input file to data/
        try:
            with open(saved_input_path, "wb") as f:
                f.write(content)
            logger.info(f"Saved uploaded file to '{saved_input_path}' (Sequence #{seq_num})")
        except Exception as save_err:
            logger.error(f"File saving failure for sequence #{seq_num}: {save_err}", exc_info=True)
            return JSONResponse(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                content={
                    "status": "error",
                    "message": "Failed to save uploaded file to the server."
                }
            )

        # 4. Run Existing Pipeline Automatically
        expected_output_filename = f"output_{seq_num}.pptx"
        expected_output_path = os.path.join(OUTPUT_DIR, expected_output_filename)

        try:
            logger.info(f"Starting pipeline execution for '{saved_input_path}' -> '{expected_output_path}'")
            run_pipeline(
                excel_path=saved_input_path,
                output_path=expected_output_path,
                sequence_num=seq_num
            )

            # Verify that the output presentation exists
            if not os.path.exists(expected_output_path) or os.path.getsize(expected_output_path) == 0:
                raise IOError(f"Output file '{expected_output_path}' was not created or is empty.")

            logger.info(f"Pipeline completed successfully: {expected_output_filename}")
            return {
                "status": "success",
                "input_file": saved_input_filename,
                "output_file": expected_output_filename,
                "download_url": f"/download/{expected_output_filename}",
                "message": "File processed successfully"
            }

        except Exception as pipe_err:
            logger.error(f"Pipeline processing error for sequence #{seq_num}: {pipe_err}", exc_info=True)
            return JSONResponse(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                content={
                    "status": "error",
                    "message": "Pipeline processing error while generating presentation."
                }
            )


@app.get(
    "/presentations",
    response_model=List[PresentationItem],
    tags=["Presentation Generator"],
    summary="List Generated Presentations",
    description="Returns a list of all PowerPoint presentations currently available in the output directory."
)
def list_presentations():
    items = []
    if os.path.exists(OUTPUT_DIR):
        for fname in sorted(os.listdir(OUTPUT_DIR)):
            if fname.lower().endswith(".pptx"):
                fpath = os.path.join(OUTPUT_DIR, fname)
                stat = os.stat(fpath)
                mtime_str = datetime.fromtimestamp(stat.st_mtime).isoformat()
                items.append({
                    "filename": fname,
                    "size_bytes": stat.st_size,
                    "modified_time": mtime_str,
                    "download_url": f"/download/{fname}"
                })
    return items


@app.get(
    "/download/{filename}",
    tags=["Presentation Generator"],
    summary="Download Generated Presentation",
    description="Downloads a specific PowerPoint file (.pptx) by filename."
)
def download_output(filename: str):
    clean_filename = os.path.basename(filename)
    file_path = os.path.join(OUTPUT_DIR, clean_filename)

    if not os.path.exists(file_path):
        return JSONResponse(
            status_code=status.HTTP_404_NOT_FOUND,
            content={
                "status": "error",
                "message": f"File '{clean_filename}' not found."
            }
        )

    return FileResponse(
        path=file_path,
        filename=clean_filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
