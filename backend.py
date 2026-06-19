import json
import csv
import io
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Annotated

from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
from pydantic import BaseModel

from audit import get_audit_event, list_audit_events, log_audit_event
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
    undo_pdf_renames,
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


class SettingsTestRequest(BaseModel):
    folderPath: str = ""
    supplierPdfFolderPath: str = ""
    factoryRenameExcelPath: str = ""
    csvPath: str = ""
    zipPath: str = ""
    pricePath: str = ""


def audit_failure(user, action, exc, **details):
    log_audit_event(
        user,
        action,
        "error",
        {**details, "error": str(exc)},
    )

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


def path_status(value, expected, label, optional=False):
    raw = str(value or "").strip()
    if not raw:
        if optional:
            return {"label": label, "path": "", "ok": True, "message": "Opcional não informado."}
        return {"label": label, "path": "", "ok": False, "message": "Não informado."}
    path = Path(raw).expanduser()
    exists = path.exists()
    if expected == "dir":
        ok = exists and path.is_dir()
        expected_message = "Pasta encontrada." if ok else "Pasta não encontrada."
    else:
        ok = exists and path.is_file()
        expected_message = "Arquivo encontrado." if ok else "Arquivo não encontrado."
    return {"label": label, "path": str(path), "ok": ok, "message": expected_message}


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


@app.get("/api/admin/audit-log")
def admin_audit_log(
    limit: int = Query(default=200, ge=1, le=1000),
    current_user: dict = Depends(require_admin),
):
    return {"events": list_audit_events(limit)}


@app.post("/api/admin/settings/test")
def admin_settings_test(
    request: SettingsTestRequest,
    current_user: dict = Depends(require_admin),
):
    return {
        "items": [
            path_status(request.folderPath, "dir", "Pasta de fotos"),
            path_status(request.supplierPdfFolderPath, "dir", "Pasta de fichas PDF"),
            path_status(request.factoryRenameExcelPath, "file", "Planilha de código de fábrica"),
            path_status(request.csvPath, "file", "CSV/planilha de produtos"),
            path_status(request.zipPath, "file", "ZIP das fotos", optional=True),
            path_status(request.pricePath, "file", "Planilha de preços"),
        ]
    }


@app.get("/api/admin/audit-log/export")
def admin_audit_log_export(
    limit: int = Query(default=1000, ge=1, le=5000),
    current_user: dict = Depends(require_admin),
):
    events = list_audit_events(limit)
    out = io.StringIO()
    writer = csv.writer(out, delimiter=";")
    writer.writerow(["id", "data", "usuario", "perfil", "acao", "status", "detalhes"])
    for event in events:
        writer.writerow([
            event["id"],
            event["createdAt"],
            event["username"] or "",
            event["role"] or "",
            event["action"],
            event["status"],
            json.dumps(event["details"], ensure_ascii=False),
        ])
    return Response(
        content="\ufeff" + out.getvalue(),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="historico_operacoes.csv"'},
    )


@app.post("/api/admin/audit-log/{event_id}/undo-rename")
def admin_audit_log_undo_rename(
    event_id: int,
    current_user: dict = Depends(require_admin),
):
    event = get_audit_event(event_id)
    if not event:
        raise HTTPException(status_code=404, detail="Evento nao encontrado.")
    if event["status"] != "success" or event["action"] not in {"pdf_audit.rename", "factory_code_rename.apply"}:
        raise HTTPException(status_code=400, detail="Este evento nao pode ser desfeito.")
    details = event.get("details") or {}
    try:
        result = undo_pdf_renames(details.get("folderPath", ""), details.get("items", []))
        log_audit_event(
            current_user,
            "rename.undo",
            details={
                "eventId": event_id,
                "originalAction": event["action"],
                "folderPath": details.get("folderPath", ""),
                "undoneCount": result["undoneCount"],
                "items": result["items"],
            },
        )
        return result
    except ValueError as exc:
        audit_failure(
            current_user,
            "rename.undo",
            exc,
            eventId=event_id,
            originalAction=event["action"],
            folderPath=details.get("folderPath", ""),
        )
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/users", status_code=201)
def users_create(
    request: UserCreateRequest,
    current_user: dict = Depends(require_admin),
):
    try:
        created = create_user(request.username, request.password, request.role)
        log_audit_event(
            current_user,
            "user.create",
            details={"createdUserId": created["id"], "username": created["username"], "role": created["role"]},
        )
        return created
    except ValueError as exc:
        audit_failure(current_user, "user.create", exc, username=request.username, role=request.role)
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
        log_audit_event(
            current_user,
            "user.update",
            details={
                "updatedUserId": user_id,
                "username": updated["username"],
                "role": updated["role"],
                "active": updated["active"],
                "passwordChanged": bool(request.password),
            },
        )
        return updated
    except ValueError as exc:
        audit_failure(current_user, "user.update", exc, updatedUserId=user_id)
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.delete("/api/users/{user_id}", status_code=204)
def users_delete(user_id: int, current_user: dict = Depends(require_admin)):
    try:
        delete_user(user_id, current_user["id"])
        log_audit_event(current_user, "user.delete", details={"deletedUserId": user_id})
    except ValueError as exc:
        audit_failure(current_user, "user.delete", exc, deletedUserId=user_id)
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
    current_user: dict = Depends(require_user),
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
    csv_path: Annotated[str, Form()] = "",
    zip_path: Annotated[str, Form()] = "",
    current_user: dict = Depends(require_user),
):
    try:
        rows, _, _, _ = load_inputs(
            csv_bytes=await read_optional_file(csv_file),
            zip_bytes=await read_optional_file(zip_file),
            folder_path=folder_path,
            csv_path=csv_path,
            zip_path=zip_path,
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
        result = rename_audited_product_pdfs(
            request.folderPath,
            [item.dict() for item in request.items],
        )
        log_audit_event(
            current_user,
            "pdf_audit.rename",
            details={
                "folderPath": request.folderPath,
                "requestedCount": len(request.items),
                "renamedCount": result["renamedCount"],
                "items": result["items"],
            },
        )
        return result
    except ValueError as exc:
        audit_failure(
            current_user,
            "pdf_audit.rename",
            exc,
            folderPath=request.folderPath,
            requestedCount=len(request.items),
        )
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
        result = apply_factory_code_pdf_renames(
            request.folderPath,
            [item.dict() for item in request.items],
        )
        log_audit_event(
            current_user,
            "factory_code_rename.apply",
            details={
                "folderPath": request.folderPath,
                "requestedCount": len(request.items),
                "renamedCount": result["renamedCount"],
                "items": result["items"],
            },
        )
        return result
    except ValueError as exc:
        audit_failure(
            current_user,
            "factory_code_rename.apply",
            exc,
            folderPath=request.folderPath,
            requestedCount=len(request.items),
        )
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/excel-pdf-order/preview")
async def excel_pdf_order_preview(
    workbook_file: Annotated[UploadFile | None, File()] = None,
    pdf_files: Annotated[list[UploadFile], File()] = [],
    current_user: dict = Depends(require_user),
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
    current_user: dict = Depends(require_user),
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
    current_user: dict = Depends(require_user),
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
        log_audit_event(
            current_user,
            "excel_xml_cest.fill",
            details={"selectedCount": len(indexes), "xmlCount": len(xml_files or [])},
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
    current_user: dict = Depends(require_user),
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
    current_user: dict = Depends(require_user),
):
    try:
        return inspect_pdf_from_folder(supplier_pdf_folder_path, selected_pdf_file)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/excel-pdf-order/fill")
async def excel_pdf_order_fill(
    workbook_file: Annotated[UploadFile | None, File()] = None,
    pdf_files: Annotated[list[UploadFile], File()] = [],
    current_user: dict = Depends(require_user),
):
    try:
        data = fill_workbook_with_ordered_pdfs(
            await read_optional_file(workbook_file),
            await read_pdf_files(pdf_files),
        )
        log_audit_event(
            current_user,
            "excel_pdf_order.fill_upload",
            details={"pdfCount": len(pdf_files or [])},
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
    current_user: dict = Depends(require_user),
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
        log_audit_event(
            current_user,
            "excel_pdf_order.fill_folder",
            details={
                "folderPath": supplier_pdf_folder_path,
                "selectedCount": sum(1 for item in selected_files if str(item or "").strip()),
                "replacementCount": len(replacement_pdf_files or []),
            },
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
    csv_path: Annotated[str, Form()] = "",
    zip_path: Annotated[str, Form()] = "",
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
            csv_path=csv_path,
            zip_path=zip_path,
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
    csv_path: Annotated[str, Form()] = "",
    zip_path: Annotated[str, Form()] = "",
    price_path: Annotated[str, Form()] = "",
    custom_prices: Annotated[str, Form()] = "{}",
    current_user: dict = Depends(require_user),
):
    try:
        codes = json.loads(selected_codes)
        original_prices = read_product_prices(price_path=price_path) if include_prices else {}
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
            csv_path=csv_path,
            zip_path=zip_path,
            include_product_sheets=include_product_sheets,
            supplier_pdf_bytes=await read_optional_file(supplier_pdf_file),
            supplier_pdf_folder_path=supplier_pdf_folder_path,
            include_prices=include_prices,
            price_overrides=validated_prices if include_prices else {},
            price_path=price_path,
        )
        log_audit_event(
            current_user,
            "workbook.generate",
            details={
                "selectedCount": len(codes) if isinstance(codes, list) else 0,
                "includeProductSheets": include_product_sheets,
                "includePrices": include_prices,
                "customPriceCount": len(validated_prices),
                "folderPath": folder_path,
                "supplierPdfFolderPath": supplier_pdf_folder_path,
            },
        )
    except json.JSONDecodeError as exc:
        audit_failure(current_user, "workbook.generate", exc)
        raise HTTPException(status_code=400, detail="Selecao invalida.") from exc
    except ValueError as exc:
        audit_failure(current_user, "workbook.generate", exc)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="produtos_com_fotos.xlsx"',
        },
    )
