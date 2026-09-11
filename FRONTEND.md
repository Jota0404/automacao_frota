# Frontend desktop

A interface gráfica fica em `app.py` e usa apenas Tkinter/ttk, mantendo o motor existente em `main.py`, `services.py` e `database.py`.

## Desenvolvimento

```bash
python app.py
```

## Gerar executável Windows

```bash
python -m PyInstaller --clean --onefile --noconsole --name AutomacaoFrota app.py
```

O executável será criado em:

```text
dist/AutomacaoFrota.exe
```

O executável deve permanecer na raiz do projeto ou ser distribuído junto das pastas `modelos/`, `entrada/`, `processados/`, `erros/` e `relatorios_gerados/` e do arquivo `frota.db` quando ele já existir.

A interface não precisa que o usuário coloque o CSV manualmente em `entrada/`: ao selecionar o arquivo, ela cria uma cópia de trabalho nessa pasta, executa o processamento e deixa o original intacto.
