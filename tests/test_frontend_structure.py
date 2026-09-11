"""Smoke checks that do not instantiate the desktop GUI."""
from __future__ import annotations

import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "app.py"


def main() -> None:
    source = SOURCE.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(SOURCE))

    classes = {node.name for node in tree.body if isinstance(node, ast.ClassDef)}
    assert "App" in classes, "Classe App não encontrada"

    app_class = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "App")
    methods = {node.name for node in app_class.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))}

    required = {
        "show_home",
        "select_csv",
        "_on_drop",
        "_selecionar_caminho",
        "remove_csv",
        "start",
        "show_processing",
        "show_success",
        "show_error",
        "open_folder",
        "copy_files",
        "reset",
    }
    missing = required - methods
    assert not missing, f"Métodos obrigatórios ausentes: {sorted(missing)}"

    imports = {
        alias.name.split(".")[0]
        for node in tree.body
        if isinstance(node, ast.Import)
        for alias in node.names
    }
    import_from = {
        node.module.split(".")[0]
        for node in tree.body
        if isinstance(node, ast.ImportFrom) and node.module
    }
    all_imports = imports | import_from
    assert "tkinter" in all_imports, "Tkinter não encontrado"
    assert "tkinterdnd2" in all_imports, "TkinterDnD2 não encontrado"
    assert "main" in all_imports, "Motor principal não encontrado"
    assert "database" in all_imports, "Camada de banco não encontrada"
    assert "config" in all_imports, "Configuração não encontrada"

    print("Frontend structure checks: OK")


if __name__ == "__main__":
    main()
