from pathlib import Path


def verified_rename_pdf(source, target, display_name):
    source = Path(source)
    target = Path(target)
    if source == target:
        if not target.is_file():
            raise ValueError(f"Ficha nao encontrada apos conferir: {display_name}")
        return

    try:
        source.rename(target)
    except OSError as exc:
        raise ValueError(f"Nao consegui renomear a ficha {display_name}: {exc}") from exc

    if not target.is_file():
        raise ValueError(f"A ficha {display_name} nao apareceu com o novo nome: {target.name}")
    if source.exists():
        raise ValueError(f"A ficha {display_name} continuou com o nome antigo.")


def undo_pdf_renames(pdf_folder_path, renamed_items):
    folder = Path(pdf_folder_path).expanduser()
    if not str(pdf_folder_path or "").strip() or not folder.is_dir():
        raise ValueError("Pasta de fichas invalida.")
    folder = folder.resolve()
    prepared = []
    seen_sources = set()
    seen_targets = set()

    for item in renamed_items or []:
        original_name = str(item.get("sourceFile", "")).strip()
        current_name = str(item.get("targetFile", "")).strip()
        if not original_name or not current_name:
            continue
        source = (folder / current_name).resolve()
        target = (folder / original_name).resolve()
        if folder not in source.parents or not source.is_file() or source.suffix.lower() != ".pdf":
            raise ValueError(f"Ficha renomeada nao encontrada: {current_name}")
        if folder not in target.parents or target.suffix.lower() != ".pdf":
            raise ValueError(f"Nome original invalido: {original_name}")
        if source in seen_sources:
            raise ValueError(f"A mesma ficha foi selecionada mais de uma vez: {current_name}")
        if target in seen_targets:
            raise ValueError(f"Duas fichas voltariam para o mesmo nome: {target.name}")
        if target != source and target.exists():
            raise ValueError(f"Ja existe uma ficha com o nome original: {target.name}")
        prepared.append((source, target, current_name))
        seen_sources.add(source)
        seen_targets.add(target)

    undone = []
    for source, target, source_name in prepared:
        verified_rename_pdf(source, target, source_name)
        undone.append({
            "sourceFile": str(source.relative_to(folder)),
            "targetFile": str(target.relative_to(folder)),
        })
    return {"undoneCount": len(undone), "items": undone}
