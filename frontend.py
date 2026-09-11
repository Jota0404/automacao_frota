"""
frontend.py

Interface desktop do sistema de automação de controle de frota.

A interface é deliberadamente separada do motor de processamento:
- main.py continua responsável pelo fluxo de negócio;
- frontend.py apenas coleta o CSV, dispara o processamento em segundo
  plano e apresenta o resultado ao usuário.

O frontend usa apenas Tkinter/ttk, disponíveis na distribuição padrão
 do Python para Windows, evitando uma dependência adicional de UI.
"""

from __future__ import annotations

import io
import os
import shutil
import threading
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import config
import main as processamento


APP_TITLE = "Automação de Frota"
APP_SUBTITLE = "Controle e geração de relatórios de abastecimento"
WINDOW_SIZE = "820x620"


class AutomacaoFrotaApp(tk.Tk):
    """Janela principal do aplicativo desktop."""

    def __init__(self) -> None:
        super().__init__()

        self.title(APP_TITLE)
        self.geometry(WINDOW_SIZE)
        self.minsize(720, 560)
        self.configure(bg="#F4F6F8")

        self.csv_origem: Path | None = None
        self.csv_processado: Path | None = None
        self.arquivos_gerados: list[Path] = []
        self._processing = False

        self._configurar_estilo()
        self._criar_interface()
        self._mostrar_inicio()

    # ------------------------------------------------------------------
    # ESTILO
    # ------------------------------------------------------------------

    def _configurar_estilo(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "App.TFrame",
            background="#F4F6F8",
        )
        style.configure(
            "Card.TFrame",
            background="#FFFFFF",
            relief="flat",
        )
        style.configure(
            "Title.TLabel",
            background="#F4F6F8",
            foreground="#172033",
            font=("Segoe UI", 24, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background="#F4F6F8",
            foreground="#697386",
            font=("Segoe UI", 11),
        )
        style.configure(
            "CardTitle.TLabel",
            background="#FFFFFF",
            foreground="#172033",
            font=("Segoe UI", 14, "bold"),
        )
        style.configure(
            "CardText.TLabel",
            background="#FFFFFF",
            foreground="#697386",
            font=("Segoe UI", 10),
        )
        style.configure(
            "File.TLabel",
            background="#FFFFFF",
            foreground="#172033",
            font=("Segoe UI", 11, "bold"),
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(22, 10),
            foreground="#FFFFFF",
            background="#1F5EFF",
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#184DDB"), ("disabled", "#AEB8C8")],
            foreground=[("disabled", "#E9EDF3")],
        )
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            padding=(16, 8),
            foreground="#172033",
            background="#EEF1F5",
            borderwidth=0,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#E1E6ED")],
        )
        style.configure(
            "Success.TLabel",
            background="#FFFFFF",
            foreground="#15803D",
            font=("Segoe UI", 13, "bold"),
        )
        style.configure(
            "Error.TLabel",
            background="#FFFFFF",
            foreground="#B42318",
            font=("Segoe UI", 13, "bold"),
        )
        style.configure(
            "Status.TLabel",
            background="#FFFFFF",
            foreground="#697386",
            font=("Segoe UI", 10),
        )
        style.configure(
            "Progress.Horizontal.TProgressbar",
            troughcolor="#E9EDF3",
            background="#1F5EFF",
            borderwidth=0,
            thickness=8,
        )

    # ------------------------------------------------------------------
    # CONSTRUÇÃO DA INTERFACE
    # ------------------------------------------------------------------

    def _criar_interface(self) -> None:
        self.container = ttk.Frame(self, style="App.TFrame", padding=(42, 34, 42, 28))
        self.container.pack(fill="both", expand=True)

        header = ttk.Frame(self.container, style="App.TFrame")
        header.pack(fill="x", pady=(0, 26))

        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text=APP_SUBTITLE, style="Subtitle.TLabel").pack(anchor="w", pady=(4, 0))

        self.content = ttk.Frame(self.container, style="App.TFrame")
        self.content.pack(fill="both", expand=True)

    def _limpar_conteudo(self) -> None:
        for widget in self.content.winfo_children():
            widget.destroy()

    def _criar_card(self) -> ttk.Frame:
        card = ttk.Frame(self.content, style="Card.TFrame", padding=28)
        card.pack(fill="x", pady=(0, 18))
        return card

    # ------------------------------------------------------------------
    # TELA INICIAL
    # ------------------------------------------------------------------

    def _mostrar_inicio(self) -> None:
        self._limpar_conteudo()

        card = self._criar_card()
        ttk.Label(card, text="Novo processamento", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Selecione o relatório CSV mensal exportado do Prime Benefícios.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(5, 20))

        upload = tk.Frame(card, bg="#F7F9FC", highlightbackground="#D7DEE8", highlightthickness=1)
        upload.pack(fill="x", ipady=28)

        tk.Label(
            upload,
            text="CSV",
            bg="#F7F9FC",
            fg="#1F5EFF",
            font=("Segoe UI", 20, "bold"),
        ).pack(pady=(4, 8))
        tk.Label(
            upload,
            text="Selecione um arquivo .csv para iniciar",
            bg="#F7F9FC",
            fg="#697386",
            font=("Segoe UI", 10),
        ).pack()

        ttk.Button(
            upload,
            text="Selecionar CSV",
            style="Secondary.TButton",
            command=self._selecionar_csv,
        ).pack(pady=(14, 4))

        self.file_card = ttk.Frame(self.content, style="Card.TFrame", padding=22)
        self.file_card.pack(fill="x", pady=(0, 18))
        self.file_card.pack_forget()

        self.file_name_label = ttk.Label(self.file_card, text="", style="File.TLabel")
        self.file_name_label.pack(anchor="w")
        self.file_info_label = ttk.Label(self.file_card, text="", style="CardText.TLabel")
        self.file_info_label.pack(anchor="w", pady=(4, 0))

        actions = ttk.Frame(self.content, style="App.TFrame")
        actions.pack(fill="x")

        self.generate_button = ttk.Button(
            actions,
            text="GERAR RELATÓRIOS",
            style="Primary.TButton",
            command=self._iniciar_processamento,
            state="disabled",
        )
        self.generate_button.pack(side="left")

        ttk.Label(
            actions,
            text="O CSV será processado e arquivado automaticamente.",
            style="Subtitle.TLabel",
        ).pack(side="left", padx=(14, 0))

    def _selecionar_csv(self) -> None:
        caminho = filedialog.askopenfilename(
            title="Selecionar relatório CSV",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return

        arquivo = Path(caminho)
        if arquivo.suffix.lower() != ".csv":
            messagebox.showerror(APP_TITLE, "Selecione um arquivo CSV válido.")
            return

        self.csv_origem = arquivo
        tamanho_kb = arquivo.stat().st_size / 1024

        self.file_card.pack(fill="x", pady=(0, 18))
        self.file_name_label.configure(text=arquivo.name)
        self.file_info_label.configure(text=f"{tamanho_kb:,.1f} KB • Pronto para processamento")
        self.generate_button.configure(state="normal")

    # ------------------------------------------------------------------
    # PROCESSAMENTO
    # ------------------------------------------------------------------

    def _iniciar_processamento(self) -> None:
        if self._processing or self.csv_origem is None:
            return

        self._processing = True
        self.generate_button.configure(state="disabled")
        self._mostrar_processamento()

        thread = threading.Thread(target=self._processar_em_background, daemon=True)
        thread.start()

    def _mostrar_processamento(self) -> None:
        self._limpar_conteudo()

        card = self._criar_card()
        ttk.Label(card, text="Gerando relatórios...", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="O sistema está processando o arquivo e gerando os documentos.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(5, 22))

        self.progress = ttk.Progressbar(
            card,
            mode="indeterminate",
            style="Progress.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", pady=(0, 18))
        self.progress.start(10)

        self.status_label = ttk.Label(card, text="Processando CSV...", style="Status.TLabel")
        self.status_label.pack(anchor="w")

    def _processar_em_background(self) -> None:
        assert self.csv_origem is not None

        inicio = datetime.now()
        buffer_stdout = io.StringIO()
        buffer_stderr = io.StringIO()
        caminho_entrada: Path | None = None
        sucesso = False
        erro_detalhado = ""

        try:
            config.PASTA_ENTRADA.mkdir(parents=True, exist_ok=True)
            caminho_entrada = config.PASTA_ENTRADA / self.csv_origem.name

            # Evita sobrescrever silenciosamente um arquivo já colocado na fila.
            if caminho_entrada.exists():
                sufixo = datetime.now().strftime("%Y%m%d_%H%M%S")
                caminho_entrada = config.PASTA_ENTRADA / (
                    f"{self.csv_origem.stem}_{sufixo}{self.csv_origem.suffix}"
                )

            shutil.copy2(self.csv_origem, caminho_entrada)

            arquivos_antes = {
                p: p.stat().st_mtime
                for p in config.PASTA_RELATORIOS_GERADOS.glob("*")
                if p.is_file()
            }

            with redirect_stdout(buffer_stdout), redirect_stderr(buffer_stderr):
                sucesso = processamento.processar_arquivo_csv(caminho_entrada)

            arquivos_depois = [
                p
                for p in config.PASTA_RELATORIOS_GERADOS.glob("*")
                if p.is_file() and p.stat().st_mtime >= inicio.timestamp()
            ]

            # Se o Windows tiver uma precisão de timestamp mais baixa,
            # inclui também arquivos que já existiam mas foram atualizados.
            for caminho, mtime in arquivos_antes.items():
                if caminho.exists() and caminho not in arquivos_depois and caminho.stat().st_mtime > mtime:
                    arquivos_depois.append(caminho)

            self.arquivos_gerados = sorted(set(arquivos_depois), key=lambda p: p.name.lower())

            if not sucesso:
                erro_detalhado = buffer_stdout.getvalue().strip() or buffer_stderr.getvalue().strip()
                if not erro_detalhado:
                    erro_detalhado = "O arquivo não pôde ser processado."

        except Exception as exc:  # noqa: BLE001
            sucesso = False
            erro_detalhado = str(exc)

            # Se uma cópia temporária ficou na entrada por uma falha anterior,
            # remove somente essa cópia para não deixar lixo operacional.
            if caminho_entrada is not None and caminho_entrada.exists():
                try:
                    caminho_entrada.unlink()
                except OSError:
                    pass

        self.after(0, self._finalizar_processamento, sucesso, erro_detalhado)

    def _finalizar_processamento(self, sucesso: bool, erro_detalhado: str) -> None:
        self._processing = False

        if hasattr(self, "progress"):
            self.progress.stop()

        if sucesso:
            self._mostrar_sucesso()
        else:
            self._mostrar_erro(erro_detalhado)

    # ------------------------------------------------------------------
    # TELA DE SUCESSO
    # ------------------------------------------------------------------

    def _mostrar_sucesso(self) -> None:
        self._limpar_conteudo()

        card = self._criar_card()
        ttk.Label(card, text="✓", style="Success.TLabel", font=("Segoe UI", 26, "bold")).pack(anchor="w")
        ttk.Label(card, text="Relatórios gerados!", style="CardTitle.TLabel").pack(anchor="w", pady=(4, 4))
        ttk.Label(
            card,
            text=f"{len(self.arquivos_gerados)} arquivo(s) foram gerados com sucesso.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(0, 18))

        lista = tk.Frame(card, bg="#F7F9FC", highlightbackground="#E1E6ED", highlightthickness=1)
        lista.pack(fill="x")

        if self.arquivos_gerados:
            for arquivo in self.arquivos_gerados:
                tk.Label(
                    lista,
                    text=f"• {arquivo.name}",
                    bg="#F7F9FC",
                    fg="#344054",
                    anchor="w",
                    font=("Segoe UI", 9),
                ).pack(fill="x", padx=14, pady=5)
        else:
            tk.Label(
                lista,
                text="Os relatórios foram processados, mas não foi possível listar os arquivos gerados.",
                bg="#F7F9FC",
                fg="#697386",
                anchor="w",
                font=("Segoe UI", 9),
            ).pack(fill="x", padx=14, pady=12)

        actions = ttk.Frame(self.content, style="App.TFrame")
        actions.pack(fill="x", pady=(4, 0))

        ttk.Button(actions, text="ABRIR PASTA", style="Secondary.TButton", command=self._abrir_pasta_relatorios).pack(side="left")
        ttk.Button(actions, text="COPIAR ARQUIVOS", style="Primary.TButton", command=self._copiar_arquivos).pack(side="left", padx=(10, 0))
        ttk.Button(actions, text="NOVO PROCESSAMENTO", style="Secondary.TButton", command=self._novo_processamento).pack(side="right")

    # ------------------------------------------------------------------
    # TELA DE ERRO
    # ------------------------------------------------------------------

    def _mostrar_erro(self, detalhe: str) -> None:
        self._limpar_conteudo()

        card = self._criar_card()
        ttk.Label(card, text="!", style="Error.TLabel", font=("Segoe UI", 26, "bold")).pack(anchor="w")
        ttk.Label(card, text="Não foi possível gerar os relatórios", style="CardTitle.TLabel").pack(anchor="w", pady=(4, 4))
        ttk.Label(
            card,
            text="O arquivo não pôde ser processado. Confira os detalhes abaixo.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(0, 14))

        detalhe_frame = tk.Frame(card, bg="#FFF5F5", highlightbackground="#F1C5C5", highlightthickness=1)
        detalhe_frame.pack(fill="both", expand=True)

        text = tk.Text(
            detalhe_frame,
            height=12,
            wrap="word",
            bg="#FFF5F5",
            fg="#7A271A",
            relief="flat",
            borderwidth=0,
            font=("Consolas", 9),
            padx=12,
            pady=10,
        )
        text.pack(fill="both", expand=True)
        text.insert("1.0", detalhe)
        text.configure(state="disabled")

        actions = ttk.Frame(self.content, style="App.TFrame")
        actions.pack(fill="x", pady=(4, 0))

        ttk.Button(actions, text="VOLTAR", style="Secondary.TButton", command=self._mostrar_inicio).pack(side="left")
        ttk.Button(actions, text="NOVO PROCESSAMENTO", style="Primary.TButton", command=self._novo_processamento).pack(side="right")

    # ------------------------------------------------------------------
    # AÇÕES DE ARQUIVOS
    # ------------------------------------------------------------------

    def _abrir_pasta_relatorios(self) -> None:
        config.PASTA_RELATORIOS_GERADOS.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(config.PASTA_RELATORIOS_GERADOS))
        except AttributeError:
            messagebox.showinfo(APP_TITLE, f"Relatórios em:\n{config.PASTA_RELATORIOS_GERADOS}")

    def _copiar_arquivos(self) -> None:
        if not self.arquivos_gerados:
            messagebox.showwarning(APP_TITLE, "Nenhum arquivo gerado foi localizado para copiar.")
            return

        destino = filedialog.askdirectory(title="Escolher pasta de destino")
        if not destino:
            return

        destino_path = Path(destino)
        copiados = 0

        try:
            for arquivo in self.arquivos_gerados:
                if arquivo.exists():
                    shutil.copy2(arquivo, destino_path / arquivo.name)
                    copiados += 1

            messagebox.showinfo(
                APP_TITLE,
                f"{copiados} arquivo(s) copiado(s) com sucesso para:\n{destino_path}",
            )
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Não foi possível copiar os arquivos:\n\n{exc}")

    def _novo_processamento(self) -> None:
        self.csv_origem = None
        self.csv_processado = None
        self.arquivos_gerados = []
        self._mostrar_inicio()


# ----------------------------------------------------------------------
# PONTO DE ENTRADA
# ----------------------------------------------------------------------


def main() -> None:
    app = AutomacaoFrotaApp()
    app.mainloop()


if __name__ == "__main__":
    main()
