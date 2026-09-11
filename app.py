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

BaseTk = TkinterDnD.Tk if TkinterDnD is not None else tk.Tk


class App(BaseTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(TITLE)
        self.geometry("820x650")
        self.minsize(700, 520)
        self.configure(bg="#F4F6F8")
        self.csv: Path | None = None
        self.generated: list[Path] = []
        self.processing = False

        self._style()
        self._layout()
        self.show_home()

    def _style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background="#F4F6F8")
        style.configure("Card.TFrame", background="#FFFFFF")
        style.configure("Title.TLabel", background="#F4F6F8", foreground="#172033", font=("Segoe UI", 24, "bold"))
        style.configure("Subtitle.TLabel", background="#F4F6F8", foreground="#697386", font=("Segoe UI", 11))
        style.configure("CardTitle.TLabel", background="#FFFFFF", foreground="#172033", font=("Segoe UI", 14, "bold"))
        style.configure("CardText.TLabel", background="#FFFFFF", foreground="#697386", font=("Segoe UI", 10))
        style.configure("File.TLabel", background="#FFFFFF", foreground="#172033", font=("Segoe UI", 11, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), padding=(22, 10), foreground="#FFFFFF", background="#1F5EFF", borderwidth=0)
        style.map("Primary.TButton", background=[("active", "#184DDB"), ("disabled", "#AEB8C8")])
        style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(16, 8), foreground="#172033", background="#EEF1F5", borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#E1E6ED")])
        style.configure("Success.TLabel", background="#FFFFFF", foreground="#15803D", font=("Segoe UI", 14, "bold"))
        style.configure("Error.TLabel", background="#FFFFFF", foreground="#B42318", font=("Segoe UI", 14, "bold"))
        style.configure("Status.TLabel", background="#FFFFFF", foreground="#697386", font=("Segoe UI", 10))
        style.configure("Progress.Horizontal.TProgressbar", troughcolor="#E9EDF3", background="#1F5EFF", thickness=8)
        style.configure("Vertical.TScrollbar", troughcolor="#EEF1F5", background="#C7D0DC", arrowcolor="#697386")

    def _layout(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=(34, 26, 26, 22))
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root, style="App.TFrame")
        header.pack(fill="x", pady=(0, 18))
        ttk.Label(header, text=TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text=SUBTITLE, style="Subtitle.TLabel").pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(root, style="App.TFrame")
        body.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(body, background="#F4F6F8", highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.content = ttk.Frame(self.canvas, style="App.TFrame")
        self.window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        self.content.bind("<Configure>", self._atualizar_scrollregion)
        self.canvas.bind("<Configure>", self._redimensionar_conteudo)
        self.canvas.bind_all("<MouseWheel>", self._scroll_mouse)

    def _atualizar_scrollregion(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _redimensionar_conteudo(self, event) -> None:
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _scroll_mouse(self, event) -> None:
        # Só desloca quando há conteúdo além da área visível.
        first, last = self.canvas.yview()
        if (event.delta > 0 and first > 0) or (event.delta < 0 and last < 1):
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def clear(self) -> None:
        for widget in self.content.winfo_children():
            widget.destroy()
        self.canvas.yview_moveto(0)

    def card(self) -> ttk.Frame:
        card = ttk.Frame(self.content, style="Card.TFrame", padding=24)
        card.pack(fill="x", pady=(0, 14))
        return card

    def show_home(self) -> None:
        self.clear()

        card = self.card()
        ttk.Label(card, text="Novo processamento", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Selecione o relatório CSV mensal exportado do Prime Benefícios.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(5, 16))

        drop = tk.Frame(card, bg="#F7F9FC", highlightbackground="#D7DEE8", highlightthickness=1)
        drop.pack(fill="x", ipady=24)

        tk.Label(drop, text="CSV", bg="#F7F9FC", fg="#1F5EFF", font=("Segoe UI", 20, "bold")).pack(pady=(2, 6))
        tk.Label(
            drop,
            text="Arraste e solte o arquivo aqui ou selecione manualmente",
            bg="#F7F9FC",
            fg="#697386",
            font=("Segoe UI", 10),
        ).pack()
        ttk.Button(drop, text="Selecionar CSV", style="Secondary.TButton", command=self.select_csv).pack(pady=(12, 2))

        if DND_FILES is not None and hasattr(drop, "drop_target_register"):
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self._on_drop)

        self.file_box = ttk.Frame(self.content, style="Card.TFrame", padding=18)
        self.file_name = ttk.Label(self.file_box, text="", style="File.TLabel")
        self.file_info = ttk.Label(self.file_box, text="", style="CardText.TLabel")
        self.file_name.pack(anchor="w")
        self.file_info.pack(anchor="w", pady=(4, 0))
        self.file_box.pack_forget()

        actions = ttk.Frame(self.content, style="App.TFrame")
        actions.pack(fill="x")
        self.generate = ttk.Button(actions, text="GERAR RELATÓRIOS", style="Primary.TButton", command=self.start, state="disabled")
        self.generate.pack(side="left")
        ttk.Label(actions, text="O CSV será processado e arquivado automaticamente.", style="Subtitle.TLabel").pack(side="left", padx=(14, 0))

        self.after_idle(self._atualizar_scrollregion)

    def select_csv(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecionar relatório CSV",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if selected:
            self._selecionar_caminho(Path(selected))

    def _on_drop(self, event) -> None:
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
        self.file_box.pack(fill="x", pady=(0, 14))
        self.file_name.configure(text=path.name)
        self.file_info.configure(text=f"{path.stat().st_size / 1024:,.1f} KB • Pronto para processamento")
        self.generate.configure(state="normal")
        self.after_idle(self._atualizar_scrollregion)

    def start(self) -> None:
        if self.processing or self.csv is None:
            return
        self.processing = True
        self.generate.configure(state="disabled")
        self.show_processing()
        threading.Thread(target=self._worker, daemon=True).start()

    def show_processing(self) -> None:
        self.clear()
        card = self.card()
        ttk.Label(card, text="Gerando relatórios...", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(card, text="O sistema está processando o arquivo e gerando os documentos.", style="CardText.TLabel").pack(anchor="w", pady=(5, 22))
        self.progress = ttk.Progressbar(card, mode="indeterminate", style="Progress.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 18))
        self.progress.start(10)
        ttk.Label(card, text="Processando CSV...", style="Status.TLabel").pack(anchor="w")
        self.after_idle(self._atualizar_scrollregion)

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
        card = self.card()
        ttk.Label(card, text="✓", style="Success.TLabel", font=("Segoe UI", 26, "bold")).pack(anchor="w")
        ttk.Label(card, text="Relatórios gerados!", style="CardTitle.TLabel").pack(anchor="w", pady=(4, 4))
        ttk.Label(card, text=f"{len(self.generated)} arquivo(s) foram gerados com sucesso.", style="CardText.TLabel").pack(anchor="w", pady=(0, 18))
        listing = tk.Frame(card, bg="#F7F9FC", highlightbackground="#E1E6ED", highlightthickness=1)
        listing.pack(fill="x")
        for path in self.generated:
            tk.Label(listing, text=f"• {path.name}", bg="#F7F9FC", fg="#344054", anchor="w", font=("Segoe UI", 9)).pack(fill="x", padx=14, pady=5)
        if not self.generated:
            tk.Label(listing, text="Processamento concluído, mas nenhum arquivo foi localizado.", bg="#F7F9FC", fg="#697386", anchor="w", font=("Segoe UI", 9)).pack(fill="x", padx=14, pady=12)

        actions = ttk.Frame(self.content, style="App.TFrame")
        actions.pack(fill="x", pady=(4, 0))
        ttk.Button(actions, text="ABRIR PASTA", style="Secondary.TButton", command=self.open_folder).pack(side="left")
        ttk.Button(actions, text="COPIAR ARQUIVOS", style="Primary.TButton", command=self.copy_files).pack(side="left", padx=(10, 0))
        ttk.Button(actions, text="NOVO PROCESSAMENTO", style="Secondary.TButton", command=self.reset).pack(side="right")
        self.after_idle(self._atualizar_scrollregion)

    def show_error(self, detail: str) -> None:
        self.clear()
        card = self.card()
        ttk.Label(card, text="!", style="Error.TLabel", font=("Segoe UI", 26, "bold")).pack(anchor="w")
        ttk.Label(card, text="Não foi possível gerar os relatórios", style="CardTitle.TLabel").pack(anchor="w", pady=(4, 4))
        ttk.Label(card, text="Confira os detalhes abaixo.", style="CardText.TLabel").pack(anchor="w", pady=(0, 14))
        box = tk.Frame(card, bg="#FFF5F5", highlightbackground="#F1C5C5", highlightthickness=1)
        box.pack(fill="both", expand=True)
        text = tk.Text(box, height=12, wrap="word", bg="#FFF5F5", fg="#7A271A", relief="flat", borderwidth=0, font=("Consolas", 9), padx=12, pady=10)
        text.pack(fill="both", expand=True)
        text.insert("1.0", detail)
        text.configure(state="disabled")

        actions = ttk.Frame(self.content, style="App.TFrame")
        actions.pack(fill="x", pady=(4, 0))
        ttk.Button(actions, text="VOLTAR", style="Secondary.TButton", command=self.show_home).pack(side="left")
        ttk.Button(actions, text="NOVO PROCESSAMENTO", style="Primary.TButton", command=self.reset).pack(side="right")
        self.after_idle(self._atualizar_scrollregion)

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
        try:
            count = 0
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
