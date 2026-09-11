"""Aplicativo desktop da Automação de Frota.

Camada de apresentação separada do motor existente em main.py,
services.py e database.py.
"""
from __future__ import annotations

import io
import os
import shutil
import threading
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # pragma: no cover - permite desenvolvimento sem DnD instalado
    DND_FILES = None
    TkinterDnD = None

import config
import database
import main as processamento

TITLE = "Automação de Frota"
SUBTITLE = "Controle e geração de relatórios de abastecimento"
VERSION = "Desktop"

BG = "#F3F6FA"
CARD = "#FFFFFF"
TEXT = "#172033"
MUTED = "#667085"
BORDER = "#DDE4EC"
PRIMARY = "#2463EB"
PRIMARY_DARK = "#1D4ED8"
PRIMARY_SOFT = "#EEF4FF"
SUCCESS = "#15803D"
SUCCESS_SOFT = "#ECFDF3"
DANGER = "#B42318"
DANGER_SOFT = "#FFF1F0"

BaseTk = TkinterDnD.Tk if TkinterDnD is not None else tk.Tk


class App(BaseTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(TITLE)
        self.geometry("860x680")
        self.minsize(700, 520)
        self.configure(bg=BG)

        self.csv: Path | None = None
        self.generated: list[Path] = []
        self.processing = False
        self._drop_widgets: list[tk.Widget] = []

        self._style()
        self._layout()
        self.show_home()

    def _style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 25, "bold"))
        style.configure("Subtitle.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 16, "bold"))
        style.configure("CardText.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 10))
        style.configure("Status.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 10))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(18, 10), foreground="#FFFFFF", background=PRIMARY, borderwidth=0)
        style.map("Primary.TButton", background=[("active", PRIMARY_DARK), ("disabled", "#B8C1CF")], foreground=[("disabled", "#EEF1F5")])
        style.configure("Secondary.TButton", font=("Segoe UI", 9, "bold"), padding=(14, 8), foreground=TEXT, background="#EEF2F6", borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#E1E7EE")])
        style.configure("Danger.TButton", font=("Segoe UI", 9, "bold"), padding=(12, 7), foreground=DANGER, background=DANGER_SOFT, borderwidth=0)
        style.map("Danger.TButton", background=[("active", "#FDD9D6")])
        style.configure(
            "Vertical.TScrollbar",
            troughcolor="#E9EEF5",
            background="#B9C6D8",
            bordercolor="#E9EEF5",
            darkcolor="#B9C6D8",
            lightcolor="#B9C6D8",
            arrowsize=0,
            width=10,
        )
        style.map(
            "Vertical.TScrollbar",
            background=[("pressed", PRIMARY_DARK), ("active", PRIMARY)],
        )
        style.layout(
            "Vertical.TScrollbar",
            [
                (
                    "Vertical.Scrollbar.trough",
                    {
                        "sticky": "ns",
                        "children": [
                            (
                                "Vertical.Scrollbar.thumb",
                                {"sticky": "ns", "expand": "1"},
                            )
                        ],
                    },
                )
            ],
        )
        style.configure("Horizontal.TProgressbar", troughcolor="#E8EDF3", background=PRIMARY, thickness=8)

    def _layout(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=(36, 28, 28, 22))
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root, style="App.TFrame")
        header.pack(fill="x", pady=(0, 22))
        brand = tk.Frame(header, bg=BG)
        brand.pack(side="left", fill="x", expand=True)

        logo = tk.Label(brand, text="AF", bg=PRIMARY, fg="#FFFFFF", font=("Segoe UI", 12, "bold"), padx=11, pady=6)
        logo.pack(side="left", padx=(0, 12))
        titles = tk.Frame(brand, bg=BG)
        titles.pack(side="left")
        tk.Label(titles, text=TITLE, bg=BG, fg=TEXT, font=("Segoe UI", 23, "bold")).pack(anchor="w")
        tk.Label(titles, text=SUBTITLE, bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(1, 0))

        tk.Label(header, text=VERSION, bg="#EAF0F7", fg="#526071", font=("Segoe UI", 8, "bold"), padx=10, pady=5).pack(side="right", anchor="n", pady=2)

        body = ttk.Frame(root, style="App.TFrame")
        body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, background=BG, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.content = ttk.Frame(self.canvas, style="App.TFrame")
        self.window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.content.bind("<Configure>", self._update_scrollregion)
        self.canvas.bind("<Configure>", self._resize_content)
        self.canvas.bind_all("<MouseWheel>", self._scroll_mouse)

        tk.Label(root, text="Processamento local • Seus arquivos permanecem no computador", bg=BG, fg="#8A95A5", font=("Segoe UI", 8)).pack(fill="x", pady=(10, 0))

    def _update_scrollregion(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _resize_content(self, event) -> None:
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _scroll_mouse(self, event) -> None:
        first, last = self.canvas.yview()
        if (event.delta > 0 and first > 0) or (event.delta < 0 and last < 1):
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def clear(self) -> None:
        for widget in self.content.winfo_children():
            widget.destroy()
        self.canvas.yview_moveto(0)

    def card(self, padding: int = 26) -> ttk.Frame:
        card = ttk.Frame(self.content, style="Card.TFrame", padding=padding)
        card.pack(fill="x", pady=(0, 16))
        return card

    def _pill(self, parent, text: str, bg: str, fg: str) -> tk.Label:
        return tk.Label(parent, text=text, bg=bg, fg=fg, font=("Segoe UI", 8, "bold"), padx=9, pady=4)

    def show_home(self) -> None:
        self.clear()
        card = self.card()
        top = tk.Frame(card, bg=CARD)
        top.pack(fill="x")
        tk.Label(top, text="Novo processamento", bg=CARD, fg=TEXT, font=("Segoe UI", 16, "bold")).pack(side="left")
        self._pill(top, "CSV → RELATÓRIOS", PRIMARY_SOFT, PRIMARY).pack(side="right")
        tk.Label(card, text="Importe o relatório mensal do Prime Benefícios para gerar automaticamente as fichas e o consolidado da frota.", bg=CARD, fg=MUTED, font=("Segoe UI", 10), wraplength=720, justify="left").pack(anchor="w", pady=(7, 22))

        drop_outer = tk.Frame(card, bg="#D9E2EE", padx=1, pady=1)
        drop_outer.pack(fill="x")
        drop = tk.Frame(drop_outer, bg="#F8FAFD", height=220)
        drop.pack(fill="x")
        drop.pack_propagate(False)
        icon = tk.Label(drop, text="CSV", bg="#EAF2FF", fg=PRIMARY, font=("Segoe UI", 17, "bold"), padx=18, pady=9)
        icon.pack(pady=(34, 14))
        title = tk.Label(drop, text="Solte seu arquivo CSV aqui", bg="#F8FAFD", fg=TEXT, font=("Segoe UI", 12, "bold"))
        title.pack()
        subtitle = tk.Label(drop, text="ou selecione manualmente pelo botão abaixo", bg="#F8FAFD", fg=MUTED, font=("Segoe UI", 9))
        subtitle.pack(pady=(4, 12))
        ttk.Button(drop, text="Selecionar CSV", style="Secondary.TButton", command=self.select_csv).pack()
        self._drop_widgets = [drop_outer, drop, icon, title, subtitle]
        if DND_FILES is not None and hasattr(drop, "drop_target_register"):
            self._register_drag_drop(self._drop_widgets)

        self.file_box = ttk.Frame(self.content, style="Card.TFrame", padding=16)
        self.file_box.pack(fill="x", pady=(0, 16))
        self.file_box.pack_forget()
        file_icon = tk.Label(self.file_box, text="CSV", bg="#EAF2FF", fg=PRIMARY, font=("Segoe UI", 9, "bold"), padx=8, pady=6)
        file_icon.pack(side="left", padx=(0, 12))
        file_text = tk.Frame(self.file_box, bg=CARD)
        file_text.pack(side="left", fill="x", expand=True)
        self.file_name = tk.Label(file_text, text="", bg=CARD, fg=TEXT, font=("Segoe UI", 10, "bold"), anchor="w")
        self.file_name.pack(fill="x")
        self.file_info = tk.Label(file_text, text="", bg=CARD, fg=MUTED, font=("Segoe UI", 8), anchor="w")
        self.file_info.pack(fill="x", pady=(2, 0))
        self._pill(self.file_box, "PRONTO", SUCCESS_SOFT, SUCCESS).pack(side="right", padx=(8, 10))
        ttk.Button(self.file_box, text="Remover", style="Danger.TButton", command=self.remove_csv).pack(side="right")

        actions = tk.Frame(self.content, bg=BG)
        actions.pack(fill="x", pady=(0, 16))
        self.generate = ttk.Button(actions, text="GERAR RELATÓRIOS", style="Primary.TButton", command=self.start, state="disabled")
        self.generate.pack(side="left")
        tk.Label(actions, text="O arquivo original permanece intacto.", bg=BG, fg="#8A95A5", font=("Segoe UI", 8)).pack(side="left", padx=(12, 0), pady=(2, 0))

        steps = self.card(padding=20)
        tk.Label(steps, text="Fluxo simples", bg=CARD, fg=TEXT, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 12))
        row = tk.Frame(steps, bg=CARD)
        row.pack(fill="x")
        self._step(row, "01", "Importar", "Selecione o CSV")
        self._connector(row)
        self._step(row, "02", "Processar", "Calcular e gerar")
        self._connector(row)
        self._step(row, "03", "Entregar", "Abrir ou copiar")
        self.after_idle(self._update_scrollregion)

    def _step(self, parent, number: str, title: str, detail: str) -> None:
        box = tk.Frame(parent, bg=CARD)
        box.pack(side="left", fill="x", expand=True)
        tk.Label(box, text=number, bg=PRIMARY_SOFT, fg=PRIMARY, font=("Segoe UI", 8, "bold"), padx=8, pady=5).pack(side="left", padx=(0, 8))
        text = tk.Frame(box, bg=CARD)
        text.pack(side="left")
        tk.Label(text, text=title, bg=CARD, fg=TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(text, text=detail, bg=CARD, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", pady=(1, 0))

    def _connector(self, parent) -> None:
        tk.Label(parent, text="→", bg=CARD, fg="#B0BAC8", font=("Segoe UI", 13, "bold")).pack(side="left", padx=8)

    def _register_drag_drop(self, widgets: list[tk.Widget]) -> None:
        for widget in widgets:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<DropEnter>>", self._on_drag_enter)
                widget.dnd_bind("<<DropLeave>>", self._on_drag_leave)
                widget.dnd_bind("<<Drop>>", self._on_drop)
            except tk.TclError:
                pass

    def _on_drag_enter(self, _event):
        self._set_drop_state(True)

    def _on_drag_leave(self, _event):
        self._set_drop_state(False)

    def _set_drop_state(self, active: bool) -> None:
        if not self._drop_widgets:
            return
        bg = "#EEF4FF" if active else "#F8FAFD"
        outer = "#9DBDFF" if active else "#D9E2EE"
        self._drop_widgets[0].configure(bg=outer)
        self._drop_widgets[1].configure(bg=bg)
        for widget in self._drop_widgets[2:]:
            widget.configure(bg=bg)

    def select_csv(self) -> None:
        selected = filedialog.askopenfilename(title="Selecionar relatório CSV", filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")])
        if selected:
            self._selecionar_caminho(Path(selected))

    def _on_drop(self, event) -> None:
        self._set_drop_state(False)
        try:
            paths = self.tk.splitlist(event.data)
        except tk.TclError:
            paths = [event.data]
        arquivos = [Path(path) for path in paths if path]
        if len(arquivos) != 1:
            messagebox.showwarning(TITLE, "Arraste apenas um arquivo CSV por vez.")
            return
        self._selecionar_caminho(arquivos[0])

    def _selecionar_caminho(self, path: Path) -> None:
        if path.suffix.lower() != ".csv":
            messagebox.showerror(TITLE, "Selecione um arquivo CSV válido.")
            return
        if not path.is_file():
            messagebox.showerror(TITLE, "O arquivo selecionado não foi encontrado.")
            return
        self.csv = path
        self.file_box.pack(fill="x", pady=(0, 16))
        self.file_name.configure(text=path.name)
        self.file_info.configure(text=f"{path.stat().st_size / 1024:,.1f} KB • Pronto para processamento")
        self.generate.configure(state="normal")
        self.after_idle(self._update_scrollregion)

    def remove_csv(self) -> None:
        self.csv = None
        self.file_name.configure(text="")
        self.file_info.configure(text="")
        self.file_box.pack_forget()
        self.generate.configure(state="disabled")
        self.after_idle(self._update_scrollregion)

    def start(self) -> None:
        if self.processing or self.csv is None:
            return
        self.processing = True
        self.generate.configure(state="disabled")
        self.show_processing()
        threading.Thread(target=self._worker, daemon=True).start()

    def show_processing(self) -> None:
        self.clear()
        card = self.card(padding=30)
        tk.Label(card, text="Gerando relatórios", bg=CARD, fg=TEXT, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(card, text="Estamos processando o CSV e preparando os documentos da frota.", bg=CARD, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(6, 22))
        self.progress = ttk.Progressbar(card, mode="indeterminate", style="Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 16))
        self.progress.start(10)
        status = tk.Frame(card, bg="#F8FAFD", padx=14, pady=12)
        status.pack(fill="x")
        tk.Label(status, text="PROCESSANDO", bg=PRIMARY_SOFT, fg=PRIMARY, font=("Segoe UI", 8, "bold"), padx=9, pady=4).pack(side="left")
        tk.Label(status, text="Leitura • cálculos • fichas • consolidado", bg="#F8FAFD", fg=MUTED, font=("Segoe UI", 9)).pack(side="left", padx=(10, 0))
        self.after_idle(self._update_scrollregion)

    def _worker(self) -> None:
        assert self.csv is not None
        started = datetime.now().timestamp()
        out, err = io.StringIO(), io.StringIO()
        success = False
        detail = ""
        staged: Path | None = None
        try:
            for folder in config.TODAS_AS_PASTAS:
                folder.mkdir(parents=True, exist_ok=True)
            database.init_db()
            staged = config.PASTA_ENTRADA / self.csv.name
            if staged.exists():
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                staged = config.PASTA_ENTRADA / f"{self.csv.stem}_{stamp}{self.csv.suffix}"
            shutil.copy2(self.csv, staged)
            before = {path: path.stat().st_mtime for path in config.PASTA_RELATORIOS_GERADOS.glob("*") if path.is_file()}
            with redirect_stdout(out), redirect_stderr(err):
                success = processamento.processar_arquivo_csv(staged)
            generated = [path for path in config.PASTA_RELATORIOS_GERADOS.glob("*") if path.is_file() and path.stat().st_mtime >= started]
            for path, old_mtime in before.items():
                if path.exists() and path not in generated and path.stat().st_mtime > old_mtime:
                    generated.append(path)
            self.generated = sorted(set(generated), key=lambda path: path.name.lower())
            if not success:
                detail = out.getvalue().strip() or err.getvalue().strip() or "O arquivo não pôde ser processado."
        except Exception as exc:  # noqa: BLE001
            detail = str(exc)
            if staged is not None and staged.exists():
                try:
                    staged.unlink()
                except OSError:
                    pass
        self.after(0, self.finish, success, detail)

    def finish(self, success: bool, detail: str) -> None:
        self.processing = False
        if hasattr(self, "progress"):
            self.progress.stop()
        if success:
            self.show_success()
        else:
            self.show_error(detail)

    def show_success(self) -> None:
        self.clear()
        card = self.card(padding=28)
        tk.Label(card, text="✓", bg=SUCCESS_SOFT, fg=SUCCESS, font=("Segoe UI", 20, "bold"), padx=11, pady=5).pack(anchor="w")
        tk.Label(card, text="Relatórios gerados com sucesso", bg=CARD, fg=TEXT, font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(12, 4))
        tk.Label(card, text=f"{len(self.generated)} arquivo(s) ficaram disponíveis para uso.", bg=CARD, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 18))
        listing = tk.Frame(card, bg="#F8FAFD", highlightbackground=BORDER, highlightthickness=1)
        listing.pack(fill="x")
        for path in self.generated:
            row = tk.Frame(listing, bg="#F8FAFD")
            row.pack(fill="x", padx=12, pady=6)
            tk.Label(row, text="DOC", bg="#EAF2FF", fg=PRIMARY, font=("Segoe UI", 7, "bold"), padx=7, pady=4).pack(side="left")
            tk.Label(row, text=path.name, bg="#F8FAFD", fg=TEXT, font=("Segoe UI", 9), anchor="w").pack(side="left", padx=(10, 0), fill="x", expand=True)
        if not self.generated:
            tk.Label(listing, text="Processamento concluído, mas nenhum arquivo foi localizado.", bg="#F8FAFD", fg=MUTED, font=("Segoe UI", 9)).pack(padx=12, pady=14)
        actions = tk.Frame(self.content, bg=BG)
        actions.pack(fill="x", pady=(0, 16))
        ttk.Button(actions, text="ABRIR PASTA", style="Secondary.TButton", command=self.open_folder).pack(side="left")
        ttk.Button(actions, text="COPIAR ARQUIVOS", style="Primary.TButton", command=self.copy_files).pack(side="left", padx=(10, 0))
        ttk.Button(actions, text="NOVO PROCESSAMENTO", style="Secondary.TButton", command=self.reset).pack(side="right")
        self.after_idle(self._update_scrollregion)

    def show_error(self, detail: str) -> None:
        self.clear()
        card = self.card(padding=28)
        tk.Label(card, text="!", bg=DANGER_SOFT, fg=DANGER, font=("Segoe UI", 20, "bold"), padx=13, pady=5).pack(anchor="w")
        tk.Label(card, text="Não foi possível gerar os relatórios", bg=CARD, fg=TEXT, font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(12, 4))
        tk.Label(card, text="O arquivo foi encaminhado para o fluxo de erros. Confira os detalhes abaixo.", bg=CARD, fg=MUTED, font=("Segoe UI", 10), wraplength=720, justify="left").pack(anchor="w", pady=(0, 16))
        detail_box = tk.Frame(card, bg="#FFF8F7", highlightbackground="#F2C7C3", highlightthickness=1)
        detail_box.pack(fill="both")
        text = tk.Text(detail_box, height=10, wrap="word", bg="#FFF8F7", fg="#7A271A", relief="flat", borderwidth=0, font=("Consolas", 9), padx=12, pady=10)
        text.pack(fill="both", expand=True)
        text.insert("1.0", detail)
        text.configure(state="disabled")
        actions = tk.Frame(self.content, bg=BG)
        actions.pack(fill="x")
        ttk.Button(actions, text="VOLTAR", style="Secondary.TButton", command=self.show_home).pack(side="left")
        ttk.Button(actions, text="NOVO PROCESSAMENTO", style="Primary.TButton", command=self.reset).pack(side="right")
        self.after_idle(self._update_scrollregion)

    def open_folder(self) -> None:
        config.PASTA_RELATORIOS_GERADOS.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(config.PASTA_RELATORIOS_GERADOS))
        except AttributeError:
            messagebox.showinfo(TITLE, f"Relatórios em:\n{config.PASTA_RELATORIOS_GERADOS}")

    def copy_files(self) -> None:
        if not self.generated:
            messagebox.showwarning(TITLE, "Nenhum arquivo gerado foi localizado para copiar.")
            return
        target = filedialog.askdirectory(title="Escolher pasta de destino")
        if not target:
            return
        destination = Path(target)
        count = 0
        try:
            for path in self.generated:
                if path.exists():
                    shutil.copy2(path, destination / path.name)
                    count += 1
            messagebox.showinfo(TITLE, f"{count} arquivo(s) copiado(s) com sucesso para:\n{destination}")
        except OSError as exc:
            messagebox.showerror(TITLE, f"Não foi possível copiar os arquivos:\n\n{exc}")

    def reset(self) -> None:
        self.csv = None
        self.generated = []
        self.show_home()


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
