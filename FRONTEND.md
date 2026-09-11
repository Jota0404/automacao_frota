# Frontend desktop

A interface gráfica fica em `app.py` e mantém o motor existente em `main.py`, `services.py` e `database.py` separado da apresentação.

## Desenvolvimento

Instale as dependências principais e as do frontend:

```bash
python -m pip install -r requirements.txt
python -m pip install -r requirements-frontend.txt
```

Execute:

```bash
python app.py
```

A interface permite selecionar um CSV ou arrastá-lo diretamente para a área de upload. O arquivo original permanece intacto; uma cópia de trabalho é criada em `entrada/` para usar o fluxo existente.

A janela possui rolagem vertical quando o espaço disponível for menor que o conteúdo, evitando cortes quando não estiver maximizada.

## Gerar executável Windows

```bash
python -m PyInstaller --clean --onefile --noconsole --name AutomacaoFrota app.py
```

O executável será criado em:

```text
dist/AutomacaoFrota.exe
```

O executável deve permanecer na raiz do projeto ou ser distribuído junto das pastas `modelos/`, `entrada/`, `processados/`, `erros/` e `relatorios_gerados` e do arquivo `frota.db` quando ele já existir.
