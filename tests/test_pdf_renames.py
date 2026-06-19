from pathlib import Path

import pytest

from core import apply_factory_code_pdf_renames, undo_pdf_renames


def write_pdf(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"%PDF-1.4\n% test\n")


def test_apply_factory_code_pdf_renames_moves_file(tmp_path):
    write_pdf(tmp_path / "origem.pdf")

    result = apply_factory_code_pdf_renames(
        tmp_path,
        [{"sourceFile": "origem.pdf", "targetFile": "destino.pdf"}],
    )

    assert result["renamedCount"] == 1
    assert not (tmp_path / "origem.pdf").exists()
    assert (tmp_path / "destino.pdf").is_file()


def test_undo_pdf_renames_moves_file_back(tmp_path):
    write_pdf(tmp_path / "destino.pdf")

    result = undo_pdf_renames(
        tmp_path,
        [{"sourceFile": "origem.pdf", "targetFile": "destino.pdf"}],
    )

    assert result["undoneCount"] == 1
    assert (tmp_path / "origem.pdf").is_file()
    assert not (tmp_path / "destino.pdf").exists()


def test_undo_pdf_renames_refuses_existing_original(tmp_path):
    write_pdf(tmp_path / "origem.pdf")
    write_pdf(tmp_path / "destino.pdf")

    with pytest.raises(ValueError, match="nome original"):
        undo_pdf_renames(
            tmp_path,
            [{"sourceFile": "origem.pdf", "targetFile": "destino.pdf"}],
        )
