"""
config.py

Arquivo de configuração central do sistema de automação de controle de
frota de veículos da Administração Regional (GDF).

Referência normativa: Decreto nº 47.091/2025.

Este módulo concentra:
    - Definição dos caminhos (pastas) utilizados pelo fluxo de trabalho;
    - Caminho do banco de dados SQLite;
    - Cadastro fixo dos veículos da frota (CADASTRO_FROTA).

Nenhuma rotina de banco de dados ou lógica de processamento é definida
aqui — este arquivo é apenas de configuração/constantes, para ser
importado pelos demais módulos do sistema (ex.: main.py, db.py, etc.).
"""

from pathlib import Path

# ---------------------------------------------------------------------------
# CAMINHOS BASE DO PROJETO
# ---------------------------------------------------------------------------

# Diretório raiz do projeto (pasta onde este arquivo config.py está localizado).
# Usar Path(__file__).resolve().parent garante que os caminhos funcionem
# corretamente independentemente de onde o script for executado.
BASE_DIR = Path(__file__).resolve().parent

# Pasta onde o relatório CSV mensal do portal "Prime Benefícios" deve ser
# depositado para processamento (entrada dos dados brutos).
PASTA_ENTRADA = BASE_DIR / "entrada"

# Pasta para onde os arquivos CSV já processados com sucesso são movidos,
# servindo como histórico/arquivamento e evitando reprocessamento indevido.
PASTA_PROCESSADOS = BASE_DIR / "processados"

# Pasta para onde são movidos (ou onde são registrados) os arquivos que
# apresentaram falha durante a leitura/validação/processamento.
PASTA_ERROS = BASE_DIR / "erros"

# Pasta onde ficam armazenados os modelos (templates) .docx utilizados
# pelo docxtpl para a geração das fichas individuais de cada veículo.
PASTA_MODELOS = BASE_DIR / "modelos"

# Pasta onde serão salvos os relatórios finais gerados pelo sistema:
# as fichas individuais em Word (.docx) e a planilha consolidada (.xlsx).
PASTA_RELATORIOS_GERADOS = BASE_DIR / "relatorios_gerados"

# Lista utilitária com todas as pastas do fluxo de trabalho, útil para que
# outro módulo (ex.: main.py) possa garantir a criação de todas elas de
# uma só vez (ex.: for pasta in TODAS_AS_PASTAS: pasta.mkdir(...)).
TODAS_AS_PASTAS = (
    PASTA_ENTRADA,
    PASTA_PROCESSADOS,
    PASTA_ERROS,
    PASTA_MODELOS,
    PASTA_RELATORIOS_GERADOS,
)

# ---------------------------------------------------------------------------
# BANCO DE DADOS
# ---------------------------------------------------------------------------

# Caminho completo do arquivo do banco de dados SQLite que armazenará o
# histórico de consumo/quilometragem processado a partir dos relatórios CSV.
CAMINHO_BANCO_DADOS = BASE_DIR / "frota.db"

# String de conexão no formato aceito pelo SQLAlchemy, pronta para ser
# utilizada por create_engine() no módulo de banco de dados (a ser criado
# em etapa futura).
DATABASE_URL = f"sqlite:///{CAMINHO_BANCO_DADOS}"

# ---------------------------------------------------------------------------
# CADASTRO FIXO DA FROTA
# ---------------------------------------------------------------------------
# Dicionário com o mapeamento cadastral das 5 placas fixas da frota da
# Administração Regional. A chave é a placa do veículo (sem hífen, em
# caixa alta, no padrão utilizado pelo portal Prime Benefícios) e o valor
# é um dicionário com os dados cadastrais correspondentes.
#
# Campos de cada veículo:
#   - "marca":     Fabricante do veículo.
#   - "modelo":    Modelo do veículo.
#   - "ano":       Ano de fabricação/modelo, no formato "AAAA/AAAA".
CADASTRO_FROTA = {
    "SSK9F34": {
        "marca": "Chevrolet",
        "modelo": "Spin",
        "ano": "2022/2023",
    },
    "SSL4B53": {
        "marca": "Chevrolet",
        "modelo": "Spin",
        "ano": "2023/2024",
    },
    "TVK1E28": {
        "marca": "Fiat",
        "modelo": "Strada",
        "ano": "2023/2023",
    },
    "SGY0B50": {
        "marca": "Nissan",
        "modelo": "Sentra",
        "ano": "2020/2021",
    },
    "SSF2D14": {
        "marca": "Chevrolet",
        "modelo": "Onix",
        "ano": "2021/2022",
    },
}