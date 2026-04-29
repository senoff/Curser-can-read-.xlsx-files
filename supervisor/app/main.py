"""FastAPI entry point for xlsx-supervisor."""

import io
import tempfile
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.responses import HTMLResponse, Response, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import storage
from .processor import process_file

# Locate templates and static dirs relative to the project root
_HERE = Path(__file__).parent
_PROJECT_ROOT = _HERE.parent
TEMPLATES_DIR = _PROJECT_ROOT / "templates"
STATIC_DIR = _PROJECT_ROOT / "static"


app = FastAPI(
    title="xlsx-supervisor",
    description="Server-side AI review for .xlsx files",
    version="0.0.1",
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    """Serve the upload page."""
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    """Accept an xlsx upload, run the review, store the result."""
    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Only .xlsx / .xlsm files are accepted")

    contents = await file.read()
    if len(contents) == 0:
        raise HTTPException(status_code=400, detail="File is empty")

    # Process via temp files (xlsxwriter writes to a path, openpyxl reads from one)
    with tempfile.TemporaryDirectory() as td:
        in_path = Path(td) / "input.xlsx"
        out_path = Path(td) / "output.xlsx"
        in_path.write_bytes(contents)

        try:
            review = process_file(in_path, out_path)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Processing failed: {e}")

        out_bytes = out_path.read_bytes()

    file_id = storage.put(
        filename=file.filename,
        content=out_bytes,
        review_summary={
            "issue_count": len(review.issues),
            "by_type": review.summary_counts,
        },
    )

    return {
        "file_id": file_id,
        "filename": file.filename,
        "issue_count": len(review.issues),
        "by_type": review.summary_counts,
        "download_url": f"/download/{file_id}",
    }


@app.post("/upload-form", response_class=HTMLResponse)
async def upload_form(request: Request, file: UploadFile = File(...)):
    """HTMX-friendly upload — returns an HTML fragment instead of JSON."""
    try:
        result = await upload(file)
    except HTTPException as e:
        return templates.TemplateResponse(
            "upload_error.html",
            {"request": request, "error": e.detail},
            status_code=e.status_code,
        )
    return templates.TemplateResponse(
        "upload_result.html",
        {"request": request, "result": result},
    )


@app.get("/download/{file_id}")
def download(file_id: str):
    """Return the processed file."""
    stored = storage.get(file_id)
    if stored is None:
        raise HTTPException(status_code=404, detail="File not found or expired")

    # Suggest a sensible download name: original-name + "-reviewed.xlsx"
    base = stored.filename.rsplit(".", 1)[0]
    download_name = f"{base}-reviewed.xlsx"

    return Response(
        content=stored.content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
    )
