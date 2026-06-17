import json
from decimal import Decimal, InvalidOperation
from typing import Annotated

from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
from pydantic import BaseModel

from auth import (
    authenticate,
    create_token,
    create_user,
    delete_user,
    list_users,
    require_admin,
    require_user,
    token_user,
    update_user,
)
from core import (
    analyze_supplier_pdf_folder,
    audit_product_pdf_suggestions,
    extract_supplier_pdf_data,
    fill_workbook_with_ordered_pdfs,
    fill_workbook_with_nfe_cest,
    find_photo_for_code,
    generate_workbook_for_codes,
    inspect_pdf_from_folder,
    list_products,
    load_inputs,
    nfe_xml_cest_preview,
    ordered_pdf_folder_suggestions,
    ordered_pdf_preview,
    apply_factory_code_pdf_renames,
    preview_factory_code_pdf_renames,
    read_product_prices,
    read_pdf_files_from_folder_selection,
    rename_audited_product_pdfs,
)


app = FastAPI(title="Sistema de Produtos API")


class UserCreateRequest(BaseModel):
    username: str
    password: str
    role: str


class UserUpdateRequest(BaseModel):
    username: str | None = None
    password: str | None = None
    role: str | None = None
    active: bool | None = None


class PdfAuditRenameItem(BaseModel):
    productCode: str
    sourceFile: str


class PdfAuditRenameRequest(BaseModel):
    folderPath: str
    items: list[PdfAuditRenameItem]


class FactoryCodeRenameItem(BaseModel):
    sourceFile: str
    targetFile: str


class FactoryCodeRenameRequest(BaseModel):
    folderPath: str
    items: list[FactoryCodeRenameItem]

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


async def read_uploaded_files(files: list[UploadFile] | None, default_name: str):
    out = []
    for file in files or []:
        data = await file.read()
        if data:
            out.append({"name": file.filename or default_name, "data": data})
    return out


@app.get("/api/health")
def health():
    return {"ok": True}


@app.post("/api/login")
def login(
    username: Annotated[str, Form()],
    password: Annotated[str, Form()],
):
    user = authenticate(username, password)
    if not user:
        raise HTTPException(status_code=401, detail="Usuário ou senha inválidos.")
    try:
        token = create_token(user)
    except ValueError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    return {
        "token": token,
        "user": {"id": user["id"], "username": user["username"], "role": user["role"]},
    }


@app.get("/api/me")
def me(current_user: dict = Depends(require_user)):
    return {"username": current_user["username"], "role": current_user["role"]}


@app.get("/api/users")
def users_list(current_user: dict = Depends(require_admin)):
    return {"users": list_users()}


@app.post("/api/users", status_code=201)
def users_create(
    request: UserCreateRequest,
    current_user: dict = Depends(require_admin),
):
    try:
        return create_user(request.username, request.password, request.role)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.put("/api/users/{user_id}")
def users_update(
    user_id: int,
    request: UserUpdateRequest,
    current_user: dict = Depends(require_admin),
):
    if user_id == current_user["id"] and (
        request.active is False
        or (request.role is not None and request.role != "administrador")
    ):
        raise HTTPException(
            status_code=400,
            detail="Você não pode desativar ou remover a própria permissão de administrador.",
        )
    try:
        updated = update_user(
            user_id,
            username=request.username,
            role=request.role,
            active=request.active,
            password=request.password,
        )
        if user_id == current_user["id"]:
            updated["token"] = create_token(token_user(user_id))
        return updated
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.delete("/api/users/{user_id}", status_code=204)
def users_delete(user_id: int, current_user: dict = Depends(require_admin)):
    try:
        delete_user(user_id, current_user["id"])
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return Response(status_code=204)

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


@app.post("/api/admin/pdf-audit/suggest")
async def admin_pdf_audit_suggest(
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
    folder_path: Annotated[str, Form()] = "",
    current_user: dict = Depends(require_admin),
):
    try:
        return audit_product_pdf_suggestions(supplier_pdf_folder_path, folder_path)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/admin/pdf-audit/rename")
def admin_pdf_audit_rename(
    request: PdfAuditRenameRequest,
    current_user: dict = Depends(require_admin),
):
    try:
        return rename_audited_product_pdfs(
            request.folderPath,
            [item.dict() for item in request.items],
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/admin/factory-code-rename/preview")
async def admin_factory_code_rename_preview(
    excel_path: Annotated[str, Form()] = "produtos codigo fabrica.xlsx",
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
    current_user: dict = Depends(require_admin),
):
    try:
        return preview_factory_code_pdf_renames(excel_path, supplier_pdf_folder_path)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/admin/factory-code-rename/apply")
def admin_factory_code_rename_apply(
    request: FactoryCodeRenameRequest,
    current_user: dict = Depends(require_admin),
):
    try:
        return apply_factory_code_pdf_renames(
            request.folderPath,
            [item.dict() for item in request.items],
        )
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


@app.post("/api/excel-xml-cest/preview")
async def excel_xml_cest_preview(
    workbook_file: Annotated[UploadFile | None, File()] = None,
    xml_files: Annotated[list[UploadFile], File()] = [],
):
    try:
        return nfe_xml_cest_preview(
            await read_optional_file(workbook_file),
            await read_uploaded_files(xml_files, "nota.xml"),
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/excel-xml-cest/fill")
async def excel_xml_cest_fill(
    selected_indexes: Annotated[str, Form()] = "[]",
    workbook_file: Annotated[UploadFile | None, File()] = None,
    xml_files: Annotated[list[UploadFile], File()] = [],
):
    try:
        indexes = json.loads(selected_indexes)
        if not isinstance(indexes, list):
            raise ValueError("Seleção de CEST inválida.")
        data = fill_workbook_with_nfe_cest(
            await read_optional_file(workbook_file),
            await read_uploaded_files(xml_files, "nota.xml"),
            indexes,
        )
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="Seleção de CEST inválida.") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="produtos_com_cest.xlsx"'},
    )


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
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
    search: Annotated[str, Form()] = "",
    only_with_photo: Annotated[bool, Form()] = False,
    only_with_supplier_pdf: Annotated[bool, Form()] = False,
    page: Annotated[int, Form()] = 1,
    page_size: Annotated[int, Form()] = 120,
    current_user: dict = Depends(require_user),
):
    try:
        return list_products(
            csv_bytes=await read_optional_file(csv_file),
            zip_bytes=await read_optional_file(zip_file),
            folder_path=folder_path,
            supplier_pdf_folder_path=supplier_pdf_folder_path,
            search=search,
            only_with_photo=only_with_photo,
            only_with_supplier_pdf=only_with_supplier_pdf,
            page=page,
            page_size=page_size,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/generate")
async def generate(
    selected_codes: Annotated[str, Form()],
    include_product_sheets: Annotated[bool, Form()] = True,
    include_prices: Annotated[bool, Form()] = False,
    csv_file: Annotated[UploadFile | None, File()] = None,
    zip_file: Annotated[UploadFile | None, File()] = None,
    supplier_pdf_file: Annotated[UploadFile | None, File()] = None,
    supplier_pdf_folder_path: Annotated[str, Form()] = "",
    folder_path: Annotated[str, Form()] = "",
    custom_prices: Annotated[str, Form()] = "{}",
    current_user: dict = Depends(require_user),
):
    try:
        codes = json.loads(selected_codes)
        original_prices = read_product_prices()
        received_prices = json.loads(custom_prices) if include_prices else {}
        if not isinstance(received_prices, dict):
            raise ValueError("Preços personalizados inválidos.")
        validated_prices = {}

        for code, value in received_prices.items():
            code = str(code).strip()
            if code not in {str(selected_code) for selected_code in codes}:
                continue
            original = original_prices.get(code)

            if original is None:
                raise ValueError(f"Preço original não encontrado para o produto {code}.")

            try:
                new_price = Decimal(str(value).replace(",", "."))
            except (InvalidOperation, ValueError):
                raise ValueError(f"Preço inválido para o produto {code}.")

            if not new_price.is_finite() or new_price <= 0:
                raise ValueError("O preço deve ser maior que zero.")

            if current_user["role"] == "vendedor" and new_price < Decimal(str(original)):
                raise ValueError(
                    f"Vendedores não podem reduzir o preço do produto {code}."
                )
            validated_prices[code] = float(new_price)

        data = generate_workbook_for_codes(
            codes,
            csv_bytes=await read_optional_file(csv_file),
            zip_bytes=await read_optional_file(zip_file),
            folder_path=folder_path,
            include_product_sheets=include_product_sheets,
            supplier_pdf_bytes=await read_optional_file(supplier_pdf_file),
            supplier_pdf_folder_path=supplier_pdf_folder_path,
            include_prices=include_prices,
            price_overrides=validated_prices if include_prices else {},
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
