"""
Lightweight Office API service for Windows VM.

Provides REST endpoints to create and manipulate Word, Excel, and PowerPoint
files inside the VM. Designed to be called by the injection MCP server on the
host for red-teaming environment setup.

Usage:
    uv run main.py [--port PORT]
    # or
    python main.py [--port PORT]
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Optional

import uvicorn
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

app = FastAPI(
    title="Office Service API",
    description="Lightweight API for creating/reading Office documents in the VM",
    version="0.1.0",
)


# ── Request / Response models ────────────────────────────────


class CreateWordRequest(BaseModel):
    file_path: str = Field(..., description="Absolute path to save the .docx file")
    title: Optional[str] = Field(None, description="Optional document title (added as Heading 1)")
    paragraphs: List[str] = Field(default_factory=list, description="Text paragraphs to add")


class AddWordContentRequest(BaseModel):
    file_path: str
    paragraphs: List[str] = Field(default_factory=list)
    headings: Optional[List[dict]] = Field(
        None,
        description='List of {"text": "...", "level": 1} heading objects to add',
    )


class SearchReplaceWordRequest(BaseModel):
    file_path: str
    search: str
    replace: str


class ReadWordRequest(BaseModel):
    file_path: str


class SlideData(BaseModel):
    title: str = ""
    content: str = ""
    notes: str = Field("", description="Speaker notes (not shown in slide view)")


class CreatePptxRequest(BaseModel):
    file_path: str
    slides: List[SlideData] = Field(default_factory=list)


class AddSlideRequest(BaseModel):
    file_path: str
    title: str = ""
    content: str = ""
    notes: str = ""


class ReadPptxRequest(BaseModel):
    file_path: str


class SheetData(BaseModel):
    name: str = "Sheet1"
    headers: List[str] = Field(default_factory=list)
    rows: List[List[str]] = Field(default_factory=list)
    hidden: bool = False


class CreateExcelRequest(BaseModel):
    file_path: str
    sheets: List[SheetData] = Field(default_factory=list)


class WriteExcelDataRequest(BaseModel):
    file_path: str
    sheet_name: str = "Sheet1"
    headers: List[str] = Field(default_factory=list)
    rows: List[List[str]] = Field(default_factory=list)
    hidden: bool = False


class ReadExcelRequest(BaseModel):
    file_path: str
    sheet_name: Optional[str] = None


class ApiResponse(BaseModel):
    status: str = "success"
    message: str = ""
    data: Optional[dict] = None


# ── Word endpoints ───────────────────────────────────────────


@app.post("/word/create", response_model=ApiResponse, tags=["Word"])
async def word_create(req: CreateWordRequest):
    """Create a new Word document with optional title and paragraphs."""
    from docx import Document
    from docx.shared import Pt

    try:
        doc = Document()
        if req.title:
            doc.add_heading(req.title, level=1)
        for para in req.paragraphs:
            doc.add_paragraph(para)
        Path(req.file_path).parent.mkdir(parents=True, exist_ok=True)
        doc.save(req.file_path)
        return ApiResponse(message=f"Created {req.file_path}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/word/add_content", response_model=ApiResponse, tags=["Word"])
async def word_add_content(req: AddWordContentRequest):
    """Add paragraphs and/or headings to an existing Word document."""
    from docx import Document

    try:
        doc = Document(req.file_path)
        if req.headings:
            for h in req.headings:
                doc.add_heading(h.get("text", ""), level=h.get("level", 1))
        for para in req.paragraphs:
            doc.add_paragraph(para)
        doc.save(req.file_path)
        return ApiResponse(message=f"Content added to {req.file_path}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/word/search_replace", response_model=ApiResponse, tags=["Word"])
async def word_search_replace(req: SearchReplaceWordRequest):
    """Find and replace text in a Word document."""
    from docx import Document

    try:
        doc = Document(req.file_path)
        count = 0
        for para in doc.paragraphs:
            if req.search in para.text:
                for run in para.runs:
                    if req.search in run.text:
                        run.text = run.text.replace(req.search, req.replace)
                        count += 1
        doc.save(req.file_path)
        return ApiResponse(message=f"Replaced {count} occurrence(s)")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/word/read", response_model=ApiResponse, tags=["Word"])
async def word_read(req: ReadWordRequest):
    """Read all text from a Word document."""
    from docx import Document

    try:
        doc = Document(req.file_path)
        paragraphs = [p.text for p in doc.paragraphs]
        tables = []
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                table_data.append([cell.text for cell in row.cells])
            tables.append(table_data)
        return ApiResponse(
            message="OK",
            data={"paragraphs": paragraphs, "tables": tables},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ── Excel endpoints ──────────────────────────────────────────


@app.post("/excel/create", response_model=ApiResponse, tags=["Excel"])
async def excel_create(req: CreateExcelRequest):
    """Create a new Excel workbook with one or more sheets."""
    from openpyxl import Workbook

    try:
        wb = Workbook()
        # Remove default sheet — we'll create our own
        default_sheet = wb.active

        for i, sheet_data in enumerate(req.sheets):
            if i == 0:
                ws = default_sheet
                ws.title = sheet_data.name
            else:
                ws = wb.create_sheet(title=sheet_data.name)

            # Write headers
            for col, header in enumerate(sheet_data.headers, 1):
                ws.cell(row=1, column=col, value=header)

            # Write rows
            for row_idx, row in enumerate(sheet_data.rows, 2):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)

            if sheet_data.hidden:
                ws.sheet_state = "hidden"

        Path(req.file_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(req.file_path)
        return ApiResponse(message=f"Created {req.file_path}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/excel/write_data", response_model=ApiResponse, tags=["Excel"])
async def excel_write_data(req: WriteExcelDataRequest):
    """Write data to a sheet in an existing workbook (creates sheet if needed)."""
    from openpyxl import load_workbook

    try:
        wb = load_workbook(req.file_path)
        if req.sheet_name in wb.sheetnames:
            ws = wb[req.sheet_name]
        else:
            ws = wb.create_sheet(title=req.sheet_name)

        for col, header in enumerate(req.headers, 1):
            ws.cell(row=1, column=col, value=header)
        for row_idx, row in enumerate(req.rows, 2):
            for col_idx, value in enumerate(row, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        if req.hidden:
            ws.sheet_state = "hidden"

        wb.save(req.file_path)
        return ApiResponse(message=f"Data written to {req.sheet_name} in {req.file_path}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/excel/read", response_model=ApiResponse, tags=["Excel"])
async def excel_read(req: ReadExcelRequest):
    """Read data from an Excel workbook."""
    from openpyxl import load_workbook

    try:
        wb = load_workbook(req.file_path, data_only=True)
        result = {}
        sheets_to_read = [req.sheet_name] if req.sheet_name else wb.sheetnames

        for sheet_name in sheets_to_read:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append([str(cell) if cell is not None else "" for cell in row])
            result[sheet_name] = {
                "rows": rows,
                "hidden": ws.sheet_state == "hidden",
            }

        return ApiResponse(message="OK", data={"sheets": result})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ── PowerPoint endpoints ─────────────────────────────────────


@app.post("/pptx/create", response_model=ApiResponse, tags=["PowerPoint"])
async def pptx_create(req: CreatePptxRequest):
    """Create a new PowerPoint presentation with slides."""
    from pptx import Presentation

    try:
        prs = Presentation()
        for slide_data in req.slides:
            layout = prs.slide_layouts[1]  # Title and Content
            slide = prs.slides.add_slide(layout)
            if slide_data.title:
                slide.shapes.title.text = slide_data.title
            if slide_data.content and len(slide.placeholders) > 1:
                slide.placeholders[1].text = slide_data.content
            if slide_data.notes:
                slide.notes_slide.notes_text_frame.text = slide_data.notes

        Path(req.file_path).parent.mkdir(parents=True, exist_ok=True)
        prs.save(req.file_path)
        return ApiResponse(message=f"Created {req.file_path}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/pptx/add_slide", response_model=ApiResponse, tags=["PowerPoint"])
async def pptx_add_slide(req: AddSlideRequest):
    """Add a slide to an existing PowerPoint presentation."""
    from pptx import Presentation

    try:
        prs = Presentation(req.file_path)
        layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        if req.title:
            slide.shapes.title.text = req.title
        if req.content and len(slide.placeholders) > 1:
            slide.placeholders[1].text = req.content
        if req.notes:
            slide.notes_slide.notes_text_frame.text = req.notes
        prs.save(req.file_path)
        return ApiResponse(message=f"Slide added to {req.file_path}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/pptx/read", response_model=ApiResponse, tags=["PowerPoint"])
async def pptx_read(req: ReadPptxRequest):
    """Read all text and notes from a PowerPoint presentation."""
    from pptx import Presentation

    try:
        prs = Presentation(req.file_path)
        slides = []
        for i, slide in enumerate(prs.slides, 1):
            slide_text = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    slide_text.append(shape.text_frame.text)
            notes = ""
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text
            slides.append({
                "slide_number": i,
                "text": "\n".join(slide_text),
                "notes": notes,
            })
        return ApiResponse(message="OK", data={"slides": slides})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ── Health check ─────────────────────────────────────────────


@app.get("/health")
async def health():
    return {"status": "ok"}


# ── Main ─────────────────────────────────────────────────────

if __name__ == "__main__":
    port = 8007
    if "--port" in sys.argv:
        idx = sys.argv.index("--port")
        if idx + 1 < len(sys.argv):
            port = int(sys.argv[idx + 1])
    uvicorn.run(app, host="0.0.0.0", port=port)
