import sys
import types

from backend import app as fastapi_app


app = fastapi_app
__path__ = []

main_module = types.ModuleType("app.main")
main_module.app = fastapi_app
sys.modules.setdefault("app.main", main_module)


def run_streamlit_app():
    import streamlit as st

    from core import (
        build_search_text,
        create_workbook_bytes,
        find_template_file,
        load_inputs,
        row_code,
        row_summary,
    )

    st.set_page_config(page_title="Gerador de cadastro com fotos", layout="wide")
    st.title("Gerador de planilha de produtos com fotos")
    st.caption("Interface Streamlit legada. O frontend principal do projeto e o React.")

    left, right = st.columns(2)

    with left:
        csv_file = st.file_uploader("CSV de produtos", type=["csv"])

    with right:
        zip_file_upload = st.file_uploader("ZIP das fotos (opcional)", type=["zip"])

    folder_path = st.text_input(
        "Ou informe o caminho da pasta de fotos no computador/servidor",
        value="",
    )

    csv_bytes = csv_file.read() if csv_file else None
    zip_bytes = zip_file_upload.read() if zip_file_upload else None

    try:
        rows, zip_obj, zip_mapping, folder_mapping = load_inputs(
            csv_bytes=csv_bytes,
            zip_bytes=zip_bytes,
            folder_path=folder_path,
        )
    except ValueError as exc:
        st.error(str(exc))
        st.stop()

    image_codes = set(zip_mapping) | set(folder_mapping)

    st.success(f"{len(rows)} produtos carregados.")
    st.write(f"Fotos indexadas por codigo: {len(image_codes)}")

    template_file = find_template_file()
    if not template_file.exists():
        st.warning("Arquivo modelo nao encontrado. Sera criada uma planilha simples.")

    if "selected_codes" not in st.session_state:
        st.session_state.selected_codes = set()

    search = st.text_input("Buscar por codigo, descricao, fornecedor ou marca")
    only_with_photo = st.checkbox("Mostrar apenas produtos com foto encontrada")
    max_show = st.slider("Quantidade exibida", 20, 500, 100, 20)

    filtered = []
    query = search.lower().strip()

    for row in rows:
        if query and query not in build_search_text(row):
            continue

        summary = row_summary(row, image_codes)
        if only_with_photo and not summary["hasPhoto"]:
            continue

        filtered.append(row)

    if st.button("Selecionar todos com foto"):
        for row in filtered:
            code = row_code(row)
            if code in image_codes:
                st.session_state.selected_codes.add(code)
        st.rerun()

    view_rows = []
    for row in filtered[:max_show]:
        summary = row_summary(row, image_codes)
        view_rows.append({
            "Selecionar": summary["code"] in st.session_state.selected_codes,
            "Codigo": summary["code"],
            "Descricao": summary["description"],
            "Fornecedor": summary["supplier"],
            "Marca": summary["brand"],
            "Foto?": "SIM" if summary["hasPhoto"] else "NAO",
        })

    edited = st.data_editor(
        view_rows,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="editor_produtos",
    )

    visible_codes = {str(row["Codigo"]) for row in view_rows}

    for row in edited:
        code = str(row["Codigo"])
        if row.get("Selecionar"):
            st.session_state.selected_codes.add(code)
        elif code in visible_codes:
            st.session_state.selected_codes.discard(code)

    selected_codes = st.session_state.selected_codes
    selected = [row for row in rows if row_code(row) in selected_codes]

    col_a, col_b = st.columns([1, 4])

    with col_a:
        st.write(f"Selecionados: {len(selected)}")

    with col_b:
        if st.button("Limpar selecao"):
            st.session_state.selected_codes = set()
            st.rerun()

    include_product_sheets = st.checkbox(
        "Incluir fichas individuais por produto",
        value=True,
        help="Desmarque para gerar apenas a aba Produtos. Fica bem mais rapido para muitos itens.",
    )

    if st.button("Gerar planilha .xlsx", disabled=not selected):
        try:
            workbook_bytes = create_workbook_bytes(
                selected,
                zip_obj,
                zip_mapping,
                folder_mapping,
                include_product_sheets=include_product_sheets,
            )
        except Exception as exc:
            st.error(f"Nao consegui gerar a planilha: {exc}")
            st.stop()

        st.download_button(
            "Baixar planilha gerada",
            data=workbook_bytes,
            file_name="produtos_com_fotos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    run_streamlit_app()
