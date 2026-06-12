import json
from typing import Annotated

from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response

from core import (
    analyze_supplier_pdf_folder,
    extract_supplier_pdf_data,
    fill_workbook_with_ordered_pdfs,
    find_photo_for_code,
    generate_workbook_for_codes,
    inspect_pdf_from_folder,
    list_products,
    load_inputs,
    ordered_pdf_folder_suggestions,
    ordered_pdf_preview,
    read_pdf_files_from_folder_selection,
)


app = FastAPI(title="Sistema de Produtos API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


async def read_optional_file(file: UploadFile | None):
    if not file:
        return None
    data = await file.read()
    return data or None


async def read_pdf_files(files: list[UploadFile] | None):
    out = []
    for file in files or []:
        data = await file.read()
        if data:
            out.append({"name": file.filename or "ficha.pdf", "data": data})
    return out


@app.get("/api/health")
def health():
    return {"ok": True}


@app.get("/api/photo/{code}")
def photo(code: str, folder_path: str = Query(default="")):
    try:
        image_path = find_photo_for_code(code, folder_path)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if not image_path:
        raise HTTPException(status_code=404, detail="Foto nao encontrada.")

    return FileResponse(
        image_path,
        headers={
            "Cache-Control": "no-store, max-age=0",
            "Pragma": "no-cache",
        },
    )


@app.post("/api/supplier-pdf/extract")
async def supplier_pdf_extract(
    supplier_pdf_file: Annotated[UploadFile | None, File()] = None,
):
    try:
        return extract_supplier_pdf_data(await read_optional_file(supplier_pdf_file))
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/supplier-pdfs/analyze-folder")
async def supplier_pdfs_analyze_folder(
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
    csv_file: Annotated[UploadFile | None, File()] = None,
    zip_file: Annotated[UploadFile | None, File()] = None,
    folder_path: Annotated[str, Form()] = "",
):
    try:
        rows, _, _, _ = load_inputs(
            csv_bytes=await read_optional_file(csv_file),
            zip_bytes=await read_optional_file(zip_file),
            folder_path=folder_path,
        )
        return analyze_supplier_pdf_folder(supplier_pdf_folder_path, rows)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/excel-pdf-order/preview")
async def excel_pdf_order_preview(
    workbook_file: Annotated[UploadFile | None, File()] = None,
    pdf_files: Annotated[list[UploadFile], File()] = [],
):
    try:
        return ordered_pdf_preview(
            await read_optional_file(workbook_file),
            await read_pdf_files(pdf_files),
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/excel-pdf-order/suggest-folder")
async def excel_pdf_order_suggest_folder(
    workbook_file: Annotated[UploadFile | None, File()] = None,
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
):
    try:
        return ordered_pdf_folder_suggestions(
            await read_optional_file(workbook_file),
            supplier_pdf_folder_path,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/excel-pdf-order/inspect-folder-file")
async def excel_pdf_order_inspect_folder_file(
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
    selected_pdf_file: Annotated[str, Form()] = "",
):
    try:
        return inspect_pdf_from_folder(supplier_pdf_folder_path, selected_pdf_file)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/excel-pdf-order/fill")
async def excel_pdf_order_fill(
    workbook_file: Annotated[UploadFile | None, File()] = None,
    pdf_files: Annotated[list[UploadFile], File()] = [],
):
    try:
        data = fill_workbook_with_ordered_pdfs(
            await read_optional_file(workbook_file),
            await read_pdf_files(pdf_files),
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="produtos_com_fotos_preenchido.xlsx"',
        },
    )


@app.post("/api/excel-pdf-order/fill-folder")
async def excel_pdf_order_fill_folder(
    workbook_file: Annotated[UploadFile | None, File()] = None,
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
    selected_pdf_files: Annotated[str, Form()] = "[]",
    replacement_pdf_files: Annotated[list[UploadFile], File()] = [],
    replacement_pdf_indexes: Annotated[list[int], Form()] = [],
):
    try:
        selected_files = json.loads(selected_pdf_files)
        if not isinstance(selected_files, list):
            raise ValueError("Selecao de fichas invalida.")
        pdf_files = read_pdf_files_from_folder_selection(supplier_pdf_folder_path, selected_files)
        replacements = await read_pdf_files(replacement_pdf_files)
        for index, replacement in zip(replacement_pdf_indexes or [], replacements):
            if index < 0:
                continue
            while len(pdf_files) <= index:
                pdf_files.append(None)
            pdf_files[index] = replacement
        data = fill_workbook_with_ordered_pdfs(
            await read_optional_file(workbook_file),
            pdf_files,
        )
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="Selecao de fichas invalida.") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="produtos_com_fotos_preenchido.xlsx"',
        },
    )


@app.post("/api/products")
async def products(
    csv_file: Annotated[UploadFile | None, File()] = None,
    zip_file: Annotated[UploadFile | None, File()] = None,
    folder_path: Annotated[str, Form()] = "",
    search: Annotated[str, Form()] = "",
    only_with_photo: Annotated[bool, Form()] = False,
    page: Annotated[int, Form()] = 1,
    page_size: Annotated[int, Form()] = 120,
):
    try:
        return list_products(
            csv_bytes=await read_optional_file(csv_file),
            zip_bytes=await read_optional_file(zip_file),
            folder_path=folder_path,
            search=search,
            only_with_photo=only_with_photo,
            page=page,
            page_size=page_size,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/generate")
async def generate(
    selected_codes: Annotated[str, Form()],
    include_product_sheets: Annotated[bool, Form()] = True,
    csv_file: Annotated[UploadFile | None, File()] = None,
    zip_file: Annotated[UploadFile | None, File()] = None,
    supplier_pdf_file: Annotated[UploadFile | None, File()] = None,
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
    folder_path: Annotated[str, Form()] = "",
):
    try:
        codes = json.loads(selected_codes)
        data = generate_workbook_for_codes(
            codes,
            csv_bytes=await read_optional_file(csv_file),
            zip_bytes=await read_optional_file(zip_file),
            folder_path=folder_path,
            include_product_sheets=include_product_sheets,
            supplier_pdf_bytes=await read_optional_file(supplier_pdf_file),
            supplier_pdf_folder_path=supplier_pdf_folder_path,
        )
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="Selecao invalida.") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="produtos_com_fotos.xlsx"',
        },
    )
