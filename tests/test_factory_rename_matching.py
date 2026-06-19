from pathlib import Path

from core import FactoryRenameProduct, find_factory_rename_product_code_in_text


def test_finds_spreadsheet_factory_code_in_carapreta_table():
    product = FactoryRenameProduct("3096", "ACEM COWBOY CONG CARAPRETA", "9000000375")
    exact = {"9000000375": [product]}
    pdf_text = """
    CONSERVACAO CODIGO INTERNO ATAK CODIGO SAP
    RESFRIADA 09010021038 9000000046
    CONGELADA 09020022024 9000000375
    PRODUTO ACEM (COWBOY)
    MARCA CARAPRETA
    """

    matched, code, status = find_factory_rename_product_code_in_text(
        exact, Path("ACEM COWBOY.pdf"), pdf_text
    )

    assert matched == product
    assert code == "9000000375"
    assert status == "codigo da planilha no PDF"


def test_uses_description_when_pdf_contains_multiple_spreadsheet_codes():
    acem = FactoryRenameProduct("3096", "ACEM COWBOY CONG CARAPRETA", "9000000375")
    picanha = FactoryRenameProduct("2699", "PICANHA CONG CARAPRETA", "9000000405")
    exact = {"9000000375": [acem], "9000000405": [picanha]}

    matched, code, status = find_factory_rename_product_code_in_text(
        exact,
        Path("ACEM COWBOY.pdf"),
        "Codigos relacionados 9000000375 9000000405. Produto ACEM COWBOY.",
    )

    assert matched == acem
    assert code == "9000000375"
    assert status == "codigo da planilha no PDF + descricao"


def test_ignores_short_numbers_from_regular_pdf_content():
    product = FactoryRenameProduct("1637", "FILE DE TILAPIA SEM PELE", "134")

    matched, code, status = find_factory_rename_product_code_in_text(
        {"134": [product]}, Path("FILE DE TILAPIA.pdf"), "Validade 134 dias"
    )

    assert matched is None
    assert code == ""
    assert status == ""
