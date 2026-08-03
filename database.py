"""
database.py

Camada de banco de dados do sistema de automação de controle de frota
de veículos da Administração Regional (GDF) - Decreto nº 47.091/2025.

Este módulo utiliza SQLAlchemy (ORM) para persistir, em um banco SQLite
(arquivo definido em config.CAMINHO_BANCO_DADOS), os seguintes dados:

    - Abastecimento: cada linha bruta do relatório CSV do "Prime
      Benefícios" (um registro por abastecimento realizado).
    - ResumoMensal: a consolidação mensal por veículo (fechamento do
      mês), utilizada para consultas históricas e cálculo de médias.

Funções expostas:
    - init_db(): cria as tabelas no banco, caso ainda não existam.
    - salvar_abastecimentos(df): persiste um DataFrame de abastecimentos.
    - salvar_resumo_mensal(dados_resumo): persiste/atualiza o fechamento
      mensal de um veículo.
    - obter_km_medio_acumulado(placa): calcula a média móvel do Km médio
      (km/l) acumulado ao longo do histórico de resumos mensais de uma
      placa.
"""

from datetime import datetime, date
from typing import Optional

import pandas as pd
from sqlalchemy import (
    create_engine,
    Column,
    Integer,
    String,
    Float,
    Date,
    DateTime,
    UniqueConstraint,
)
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy.exc import SQLAlchemyError

import config

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DO SQLALCHEMY (ENGINE, SESSION, BASE)
# ---------------------------------------------------------------------------

# O "engine" é o ponto de entrada de conexão do SQLAlchemy com o banco.
# echo=False evita que o SQL gerado seja impresso no console (pode ser
# alterado para True durante depuração, se necessário).
# connect_args={"check_same_thread": False} é recomendado para SQLite em
# cenários onde a conexão pode ser usada por mais de uma thread (ex.:
# aplicações com interface gráfica ou agendadores).
engine = create_engine(
    config.DATABASE_URL,
    echo=False,
    connect_args={"check_same_thread": False},
)

# Fábrica de sessões (SessionLocal) — cada chamada a SessionLocal() cria
# uma nova sessão de trabalho com o banco de dados.
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)

# Classe base declarativa da qual todos os modelos ORM devem herdar.
Base = declarative_base()


# ---------------------------------------------------------------------------
# MODELOS (TABELAS) ORM
# ---------------------------------------------------------------------------

class Abastecimento(Base):
    """
    Representa um único registro de abastecimento extraído do relatório
    CSV mensal do portal "Prime Benefícios".

    Cada linha do CSV processado corresponde a uma linha nesta tabela.
    """

    __tablename__ = "abastecimentos"

    id = Column(Integer, primary_key=True, autoincrement=True)

    # Placa do veículo (ex.: "SSK9F34"), utilizada para vincular o
    # abastecimento ao cadastro fixo da frota (CADASTRO_FROTA).
    placa = Column(String(10), nullable=False, index=True)

    # Data em que o abastecimento foi realizado.
    data = Column(Date, nullable=False)

    # Quantidade de combustível abastecida, em litros.
    litros = Column(Float, nullable=False)

    # Valor total pago no abastecimento (em reais).
    valor_total = Column(Float, nullable=False)

    # Leitura do hodômetro (quilometragem) no momento do abastecimento.
    km_hodometro = Column(Float, nullable=False)

    # Nome/identificação do posto onde o abastecimento foi realizado.
    posto = Column(String(150), nullable=True)

    # Data/hora em que o registro foi inserido no banco (auditoria interna).
    criado_em = Column(DateTime, default=datetime.now, nullable=False)

    def __repr__(self) -> str:
        return (
            f"<Abastecimento(placa={self.placa!r}, data={self.data!r}, "
            f"litros={self.litros!r}, km_hodometro={self.km_hodometro!r})>"
        )


class ResumoMensal(Base):
    """
    Representa o fechamento/consolidação mensal de um veículo: totais de
    quilometragem rodada, litros consumidos, valor gasto e a média de
    consumo (km/l) no período.

    Existe, no máximo, um registro de ResumoMensal por combinação
    (placa, mes_ano), garantido pela restrição de unicidade abaixo.
    """

    __tablename__ = "resumos_mensais"
    __table_args__ = (
        UniqueConstraint("placa", "mes_ano", name="uq_placa_mes_ano"),
    )

    id = Column(Integer, primary_key=True, autoincrement=True)

    # Placa do veículo ao qual o resumo se refere.
    placa = Column(String(10), nullable=False, index=True)

    # Mês/ano de referência do resumo, no formato "AAAA-MM" (ex.: "2025-01"),
    # o que facilita tanto a ordenação cronológica quanto a exibição.
    mes_ano = Column(String(7), nullable=False, index=True)

    # Total de quilômetros rodados pelo veículo no mês (calculado a partir
    # da diferença entre leituras de hodômetro).
    total_km = Column(Float, nullable=False, default=0.0)

    # Total de litros de combustível consumidos no mês.
    total_litros = Column(Float, nullable=False, default=0.0)

    # Valor total gasto com combustível no mês (em reais).
    total_valor = Column(Float, nullable=False, default=0.0)

    # Média de consumo do veículo no mês, em quilômetros por litro (km/l).
    media_kml = Column(Float, nullable=False, default=0.0)

    # Data/hora em que o registro foi inserido/atualizado (auditoria interna).
    atualizado_em = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)

    def __repr__(self) -> str:
        return (
            f"<ResumoMensal(placa={self.placa!r}, mes_ano={self.mes_ano!r}, "
            f"total_km={self.total_km!r}, media_kml={self.media_kml!r})>"
        )


# ---------------------------------------------------------------------------
# FUNÇÕES PRINCIPAIS
# ---------------------------------------------------------------------------

def init_db() -> None:
    """
    Cria todas as tabelas mapeadas (Abastecimento, ResumoMensal) no banco
    de dados SQLite, caso elas ainda não existam.

    Também garante que a pasta onde o arquivo frota.db será criado
    realmente exista, evitando erro de "diretório não encontrado" na
    primeira execução do sistema.

    Esta função é idempotente: pode ser chamada múltiplas vezes sem
    causar erro ou duplicar tabelas já existentes.
    """
    try:
        # Garante que o diretório-base do arquivo .db exista antes de o
        # SQLite tentar criar o arquivo físico.
        config.CAMINHO_BANCO_DADOS.parent.mkdir(parents=True, exist_ok=True)

        # create_all() cria apenas as tabelas que ainda não existem no
        # banco; tabelas já existentes não são alteradas nem recriadas.
        Base.metadata.create_all(bind=engine)

        print(f"[database] Banco de dados inicializado com sucesso em: {config.CAMINHO_BANCO_DADOS}")

    except SQLAlchemyError as erro:
        print(f"[database] ERRO ao inicializar o banco de dados: {erro}")
        raise


def salvar_abastecimentos(df: pd.DataFrame) -> int:
    """
    Persiste os registros de um DataFrame de abastecimentos na tabela
    `abastecimentos`.

    Args:
        df: DataFrame do pandas já limpo/validado, contendo obrigato-
            riamente as colunas: "placa", "data", "litros", "valor_total",
            "km_hodometro". A coluna "posto" é opcional.

    Returns:
        int: quantidade de registros efetivamente inseridos no banco.

    Raises:
        ValueError: se o DataFrame não contiver as colunas obrigatórias.
    """
    colunas_obrigatorias = {"placa", "data", "litros", "valor_total", "km_hodometro"}
    colunas_faltantes = colunas_obrigatorias - set(df.columns)
    if colunas_faltantes:
        raise ValueError(
            f"DataFrame de abastecimentos está sem as colunas obrigatórias: {colunas_faltantes}"
        )

    if df.empty:
        print("[database] Aviso: DataFrame de abastecimentos está vazio. Nada foi salvo.")
        return 0

    session = SessionLocal()
    registros_inseridos = 0

    try:
        # Percorre cada linha do DataFrame, convertendo-a em uma instância
        # do modelo ORM Abastecimento antes de adicioná-la à sessão.
        for _, linha in df.iterrows():
            # Normaliza a data: aceita tanto objetos datetime/date do
            # pandas quanto strings, convertendo sempre para date puro.
            valor_data = linha["data"]
            if isinstance(valor_data, str):
                valor_data = pd.to_datetime(valor_data).date()
            elif isinstance(valor_data, datetime):
                valor_data = valor_data.date()
            elif isinstance(valor_data, pd.Timestamp):
                valor_data = valor_data.date()
            # Se já for um objeto date puro, mantém como está.

            novo_abastecimento = Abastecimento(
                placa=str(linha["placa"]).strip().upper(),
                data=valor_data,
                litros=float(linha["litros"]),
                valor_total=float(linha["valor_total"]),
                km_hodometro=float(linha["km_hodometro"]),
                posto=str(linha["posto"]).strip() if "posto" in df.columns and pd.notna(linha.get("posto")) else None,
            )
            session.add(novo_abastecimento)
            registros_inseridos += 1

        # Confirma (commit) todas as inserções de uma só vez, o que é
        # mais eficiente do que um commit por linha.
        session.commit()
        print(f"[database] {registros_inseridos} abastecimento(s) salvo(s) com sucesso.")

    except (SQLAlchemyError, ValueError, KeyError) as erro:
        # Em caso de falha, desfaz qualquer alteração parcial (rollback)
        # para não deixar o banco em estado inconsistente.
        session.rollback()
        print(f"[database] ERRO ao salvar abastecimentos: {erro}")
        raise

    finally:
        # A sessão é sempre encerrada, com sucesso ou falha, para evitar
        # vazamento de conexões com o banco.
        session.close()

    return registros_inseridos


def salvar_resumo_mensal(dados_resumo: dict) -> None:
    """
    Salva (ou atualiza, caso já exista) a consolidação mensal de um
    veículo na tabela `resumos_mensais`.

    Se já existir um resumo para a mesma combinação (placa, mes_ano), os
    valores são atualizados (upsert manual) em vez de gerar um registro
    duplicado — isso permite reprocessar um mês sem quebrar o histórico.

    Args:
        dados_resumo: dicionário com as chaves obrigatórias:
            "placa" (str), "mes_ano" (str, formato "AAAA-MM"),
            "total_km" (float), "total_litros" (float),
            "total_valor" (float), "media_kml" (float).

    Raises:
        KeyError: se alguma chave obrigatória estiver ausente no dicionário.
    """
    chaves_obrigatorias = {
        "placa", "mes_ano", "total_km", "total_litros", "total_valor", "media_kml"
    }
    chaves_faltantes = chaves_obrigatorias - set(dados_resumo.keys())
    if chaves_faltantes:
        raise KeyError(f"dados_resumo está sem as chaves obrigatórias: {chaves_faltantes}")

    session = SessionLocal()

    try:
        placa = str(dados_resumo["placa"]).strip().upper()
        mes_ano = str(dados_resumo["mes_ano"]).strip()

        # Verifica se já existe um resumo para esta placa/mês, para decidir
        # entre UPDATE (atualização) ou INSERT (novo registro).
        resumo_existente = (
            session.query(ResumoMensal)
            .filter_by(placa=placa, mes_ano=mes_ano)
            .first()
        )

        if resumo_existente:
            # Atualiza os valores do resumo já existente.
            resumo_existente.total_km = float(dados_resumo["total_km"])
            resumo_existente.total_litros = float(dados_resumo["total_litros"])
            resumo_existente.total_valor = float(dados_resumo["total_valor"])
            resumo_existente.media_kml = float(dados_resumo["media_kml"])
            resumo_existente.atualizado_em = datetime.now()
            print(f"[database] Resumo mensal de {placa} ({mes_ano}) atualizado com sucesso.")
        else:
            # Cria um novo registro de resumo mensal.
            novo_resumo = ResumoMensal(
                placa=placa,
                mes_ano=mes_ano,
                total_km=float(dados_resumo["total_km"]),
                total_litros=float(dados_resumo["total_litros"]),
                total_valor=float(dados_resumo["total_valor"]),
                media_kml=float(dados_resumo["media_kml"]),
            )
            session.add(novo_resumo)
            print(f"[database] Resumo mensal de {placa} ({mes_ano}) criado com sucesso.")

        session.commit()

    except (SQLAlchemyError, ValueError, KeyError) as erro:
        session.rollback()
        print(f"[database] ERRO ao salvar resumo mensal: {erro}")
        raise

    finally:
        session.close()


def obter_km_medio_acumulado(placa: str) -> Optional[float]:
    """
    Consulta todo o histórico de resumos mensais de um veículo e calcula
    a média móvel (média simples do histórico) do Km médio (km/l)
    acumulado, ponderando pelo número de meses já registrados.

    A média retornada é a média aritmética das médias mensais de km/l
    (media_kml) de todos os meses já processados para a placa informada,
    representando o desempenho médio anual/acumulado do veículo.

    Args:
        placa: placa do veículo a ser consultado (ex.: "SSK9F34").

    Returns:
        float: a média de km/l acumulada, arredondada em 2 casas decimais.
        None: caso não exista nenhum resumo mensal salvo para a placa.
    """
    session = SessionLocal()

    try:
        placa_normalizada = placa.strip().upper()

        # Busca todos os resumos mensais da placa, ordenados cronologica-
        # mente (do mais antigo para o mais recente) pelo campo mes_ano,
        # que segue o formato "AAAA-MM" e é diretamente ordenável como texto.
        resumos = (
            session.query(ResumoMensal)
            .filter_by(placa=placa_normalizada)
            .order_by(ResumoMensal.mes_ano.asc())
            .all()
        )

        if not resumos:
            print(f"[database] Nenhum resumo mensal encontrado para a placa {placa_normalizada}.")
            return None

        # Calcula a média móvel/acumulada como a média aritmética simples
        # de todas as médias mensais (media_kml) já registradas.
        soma_medias = sum(resumo.media_kml for resumo in resumos)
        quantidade_meses = len(resumos)
        media_acumulada = soma_medias / quantidade_meses

        media_acumulada_arredondada = round(media_acumulada, 2)

        print(
            f"[database] Km médio acumulado da placa {placa_normalizada}: "
            f"{media_acumulada_arredondada} km/l (baseado em {quantidade_meses} mês(es))."
        )

        return media_acumulada_arredondada

    except SQLAlchemyError as erro:
        print(f"[database] ERRO ao consultar km médio acumulado da placa {placa}: {erro}")
        raise

    finally:
        session.close()


# ---------------------------------------------------------------------------
# EXECUÇÃO DIRETA (TESTE RÁPIDO MANUAL)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # Permite rodar "python database.py" isoladamente apenas para garantir
    # que o banco e as tabelas sejam criados corretamente, sem depender do
    # restante do sistema (main.py) ainda não desenvolvido.
    init_db()