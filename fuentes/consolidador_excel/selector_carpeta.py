"""Selector de carpetas que muestra subcarpetas y archivos Excel."""

from __future__ import annotations

import os
from pathlib import Path


EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}


def _desktop_folder() -> Path:
    if os.name == "nt":
        try:
            import ctypes

            buffer = ctypes.create_unicode_buffer(260)
            result = ctypes.windll.shell32.SHGetFolderPathW(
                None, 0x0010, None, 0, buffer
            )
            if result == 0 and buffer.value:
                desktop = Path(buffer.value)
                if desktop.is_dir():
                    return desktop
        except Exception:
            pass
    desktop = Path.home() / "Desktop"
    return desktop if desktop.is_dir() else Path.home()


def _available_drives() -> list[Path]:
    if os.name != "nt":
        return [Path("/")]
    try:
        import ctypes

        mask = ctypes.windll.kernel32.GetLogicalDrives()
        return [
            Path(f"{chr(65 + index)}:/")
            for index in range(26)
            if mask & (1 << index)
        ]
    except Exception:
        return [Path(Path.home().anchor)]


def seleccionar_carpeta_excel(
    titulo: str = "Selecciona la carpeta que contiene los archivos Excel",
) -> Path | None:
    try:
        import tkinter as tk
        from tkinter import messagebox, ttk
    except ImportError as exc:
        raise RuntimeError("Tkinter no está disponible para seleccionar la carpeta") from exc

    root = tk.Tk()
    root.title(titulo)
    root.geometry("930x570")
    root.minsize(720, 430)
    root.attributes("-topmost", True)
    root.after(300, lambda: root.attributes("-topmost", False))

    selected: Path | None = None
    current: Path | None = None
    item_paths: dict[str, Path] = {}
    path_var = tk.StringVar()
    info_var = tk.StringVar()

    outer = ttk.Frame(root, padding=12)
    outer.pack(fill="both", expand=True)

    toolbar = ttk.Frame(outer)
    toolbar.pack(fill="x", pady=(0, 10))

    ttk.Label(toolbar, text="Ubicación:").pack(side="left", padx=(0, 6))
    path_entry = ttk.Entry(toolbar, textvariable=path_var)
    path_entry.pack(side="left", fill="x", expand=True)

    columns = ("nombre", "tipo", "tamano")
    tree = ttk.Treeview(outer, columns=columns, show="headings", selectmode="browse")
    tree.heading("nombre", text="Nombre")
    tree.heading("tipo", text="Tipo")
    tree.heading("tamano", text="Tamaño")
    tree.column("nombre", width=560, minwidth=250, anchor="w")
    tree.column("tipo", width=150, minwidth=100, anchor="w")
    tree.column("tamano", width=100, minwidth=80, anchor="e")

    scrollbar = ttk.Scrollbar(outer, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="left", fill="y")

    bottom = ttk.Frame(root, padding=(12, 0, 12, 12))
    bottom.pack(fill="x")
    ttk.Label(bottom, textvariable=info_var).pack(side="left", fill="x", expand=True)

    def size_text(size: int) -> str:
        if size < 1024:
            return f"{size} B"
        if size < 1024 * 1024:
            return f"{size / 1024:.1f} KB"
        return f"{size / (1024 * 1024):.1f} MB"

    def clear_tree() -> None:
        item_paths.clear()
        for item in tree.get_children():
            tree.delete(item)

    def add_item(path: Path, kind: str, size: str = "") -> None:
        item = tree.insert("", "end", values=(path.name or str(path), kind, size))
        item_paths[item] = path

    def show_drives() -> None:
        nonlocal current
        current = None
        clear_tree()
        path_var.set("Este equipo")
        for drive in _available_drives():
            add_item(drive, "Unidad")
        info_var.set("Abre una unidad o escribe una ruta. Luego selecciona una carpeta.")
        select_button.state(["disabled"])

    def show_folder(folder: Path) -> None:
        nonlocal current
        try:
            folder = folder.expanduser().resolve()
            entries = list(folder.iterdir())
        except Exception as exc:
            messagebox.showerror(
                "No se puede abrir la carpeta", str(exc), parent=root
            )
            return

        directories = sorted(
            (path for path in entries if path.is_dir()), key=lambda path: path.name.casefold()
        )
        excels = sorted(
            (
                path
                for path in entries
                if path.is_file()
                and path.suffix.lower() in EXCEL_EXTENSIONS
                and not path.name.startswith("~$")
            ),
            key=lambda path: path.name.casefold(),
        )

        current = folder
        clear_tree()
        path_var.set(str(folder))
        for directory in directories:
            add_item(directory, "Carpeta")
        for excel in excels:
            try:
                size = size_text(excel.stat().st_size)
            except OSError:
                size = ""
            add_item(excel, "Archivo Excel", size)

        info_var.set(
            f"{len(excels)} Excel visible(s). Al confirmar también se revisarán las subcarpetas."
        )
        select_button.state(["!disabled"])

    def open_selected(_event=None) -> None:
        selection = tree.selection()
        if not selection:
            return
        path = item_paths.get(selection[0])
        if path is not None and path.is_dir():
            show_folder(path)

    def go_to_entry(_event=None) -> None:
        entered = path_var.get().strip().strip('"')
        if entered:
            show_folder(Path(entered))

    def go_up() -> None:
        if current is None:
            return
        parent = current.parent
        if parent == current:
            show_drives()
        else:
            show_folder(parent)

    def confirm() -> None:
        nonlocal selected
        if current is None:
            messagebox.showwarning(
                "Selecciona una carpeta",
                "Abre primero la carpeta que quieres analizar.",
                parent=root,
            )
            return
        selected = current
        root.destroy()

    def cancel() -> None:
        root.destroy()

    ttk.Button(toolbar, text="Ir", command=go_to_entry).pack(side="left", padx=(6, 0))
    ttk.Button(toolbar, text="Subir", command=go_up).pack(side="left", padx=(6, 0))
    ttk.Button(toolbar, text="Escritorio", command=lambda: show_folder(_desktop_folder())).pack(
        side="left", padx=(6, 0)
    )
    ttk.Button(toolbar, text="Unidades", command=show_drives).pack(side="left", padx=(6, 0))

    ttk.Button(bottom, text="Cancelar", command=cancel).pack(side="right")
    select_button = ttk.Button(bottom, text="Usar esta carpeta", command=confirm)
    select_button.pack(side="right", padx=(0, 8))

    tree.bind("<Double-1>", open_selected)
    tree.bind("<Return>", open_selected)
    path_entry.bind("<Return>", go_to_entry)
    root.protocol("WM_DELETE_WINDOW", cancel)

    show_folder(_desktop_folder())
    root.mainloop()
    return selected
