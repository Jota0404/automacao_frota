"""
main.py

Ponto de entrada / orquestrador do sistema de automação de controle de
frota de veículos da Administração Regional (GDF) - Decreto nº
47.091/2025.

Fluxo executado a cada chamada (`python main.py`):

    1. Inicializa o banco de dados (database.init_db()) e garante que
       toda a estrutura de pastas do projeto exista.
    2. Varre a pasta 'entrada/' em busca de arquivos .csv.
    3. Para cada CSV encontrado:
        a. Lê e padroniza os dados (services.processar_csv_prime).
        b. Persiste os abastecimentos brutos no banco
           (database.salvar_abastecimentos).
        c. Para cada um dos 5 veículos fixos da frota
           (config.CADASTRO_FROTA):
             - calcula as métricas do mês (services.calcular_metricas_veiculo);
             - gera a ficha individual em Word (services.gerar_ficha_word);
             - salva o fechamento mensal no banco
               (database.salvar_resumo_mensal).
        d. Gera a planilha consolidada do mês em Excel
           (services.gerar_relatorio_consolidado_excel).
        e. Em caso de sucesso, move o CSV para 'processados/'; em caso
           de qualquer falha, move o CSV para 'erros/' e segue para o
           próximo arquivo (uma falha em um CSV não interrompe o
           processamento dos demais).
    4. Exibe um resumo final da execução.

Observação: assume-se que cada arquivo CSV corresponde a um único mês
de referência (conforme descrito no enunciado do projeto: "relatório
CSV mensal do portal Prime Benefícios"). Todos os 5 veículos cadastrados
em config.CADASTRO_FROTA recebem ficha e entram no relatório consolidado
todo mês, mesmo os que não tiveram nenhum abastecimento no período —
isso mantém um registro de controle completo e auditável da frota.
"""

import shutil
from datetime import datetime
from pathlib import Path
from typing import List

import config
import database
import services

# ---------------------------------------------------------------------------
# FUNÇÕES AUXILIARES DE ORQUESTRAÇÃO
# ---------------------------------------------------------------------------

def garantir_estrutura_de_pastas() -> None:
    """
    Garante que todas as pastas do fluxo de trabalho (entrada/,
    processados/, erros/, modelos/, relatorios_gerados/) existam,
    criando as que faltarem. Idempotente: pode ser chamada em toda
    execução sem causar erro.
    """
    for pasta in config.TODAS_AS_PASTAS:
        pasta.mkdir(parents=True, exist_ok=True)


def listar_arquivos_csv(pasta: Path) -> List[Path]:
    """
    Lista, em ordem alfabética, todos os arquivos .csv presentes
    diretamente em `pasta` (não entra em subpastas).

    A comparação de extensão é feita sem diferenciar maiúsculas de
    minúsculas (aceita tanto '.csv' quanto '.CSV'), pois o
    comportamento nativo de Path.glob() varia entre sistemas de
    arquivos.

    Args:
        pasta: pasta (Path) a ser varrida.

    Returns:
        Lista de Path, ordenada alfabeticamente pelo nome do arquivo.
    """
    if not pasta.exists():
        return []
    return sorted(
        arquivo for arquivo in pasta.iterdir()
        if arquivo.is_file() and arquivo.suffix.lower() == ".csv"
    )


def mover_arquivo(caminho_origem: Path, pasta_destino: Path) -> Path:
    """
    Move um arquivo para a pasta de destino informada.

    Caso já exista um arquivo com o mesmo nome na pasta de destino
    (ex.: reprocessamento de um arquivo de mesmo nome baixado em outro
    mês), um sufixo com data/hora é adicionado ao nome para evitar
    sobrescrever o arquivo anterior por engano.

    Args:
        caminho_origem: caminho atual do arquivo.
        pasta_destino: pasta para onde o arquivo deve ser movido.

    Returns:
        Path: caminho final do arquivo após a movimentação.
    """
    pasta_destino.mkdir(parents=True, exist_ok=True)
    caminho_destino = pasta_destino / caminho_origem.name

    if caminho_destino.exists():
        sufixo = datetime.now().strftime("%Y%m%d_%H%M%S")
        novo_nome = f"{caminho_origem.stem}_{sufixo}{caminho_origem.suffix}"
        caminho_destino = pasta_destino / novo_nome

    shutil.move(str(caminho_origem), str(caminho_destino))
    return caminho_destino


def processar_arquivo_csv(caminho_csv: Path) -> bool:
    """
    Processa um único arquivo CSV do início ao fim: leitura/padronização,
    persistência no banco, geração das fichas individuais e da planilha
    consolidada, e por fim move o arquivo para 'processados/' ou
    'erros/', conforme o resultado.

    Qualquer exceção ocorrida em qualquer etapa é capturada aqui: o
    arquivo é movido para 'erros/' e o processamento segue para o
    próximo CSV da fila, em vez de interromper toda a execução.

    Args:
        caminho_csv: caminho do arquivo CSV a ser processado.

    Returns:
        bool: True se o arquivo foi processado com sucesso, False caso
        tenha sido movido para a pasta de erros.
    """
    print(f"\n{'-' * 70}")
    print(f"Processando arquivo: {caminho_csv.name}")
    print(f"{'-' * 70}")

    try:
        # --- 1. Leitura e padronização do CSV ---
        print("[main] Etapa 1/4 - Lendo e padronizando o CSV...")
        df_abastecimentos = services.processar_csv_prime(caminho_csv)

        # --- 2. Persistência dos abastecimentos brutos no banco ---
        print("[main] Etapa 2/4 - Salvando abastecimentos no banco de dados...")
        database.salvar_abastecimentos(df_abastecimentos)

        # --- 3. Métricas + ficha individual + resumo mensal, por veículo ---
        print("[main] Etapa 3/4 - Calculando métricas e gerando fichas individuais...")
        lista_metricas = []
        mes_ano_arquivo = None  # formato "MM-AAAA", usado no nome dos arquivos gerados

        for placa in config.CADASTRO_FROTA:
            km_medio_historico = database.obter_km_medio_acumulado(placa)

            contexto = services.calcular_metricas_veiculo(
                df_abastecimentos, placa, km_medio_historico
            )
            lista_metricas.append(contexto)

            services.gerar_ficha_word(contexto, config.PASTA_RELATORIOS_GERADOS)
            print(f"    -> Ficha da placa {placa} gerada com sucesso.")

            # Persiste o fechamento mensal deste veículo no banco. O
            # formato "AAAA-MM" (ano primeiro) é o mesmo já utilizado
            # pela tabela ResumoMensal em database.py.
            mes_ano_banco = f"{contexto['ano_referencia']}-{contexto['mes_numero']}"
            database.salvar_resumo_mensal({
                "placa": contexto["placa"],
                "mes_ano": mes_ano_banco,
                "total_km": contexto["total_km_mensal_num"],
                "total_litros": contexto["total_litros_num"],
                "total_valor": contexto["total_valor_num"],
                "media_kml": contexto["media_kml_num"],
            })

            # Guarda o período no formato "MM-AAAA" (mesmo padrão usado
            # no nome dos arquivos gerados por services.gerar_ficha_word).
            mes_ano_arquivo = f"{contexto['mes_numero']}-{contexto['ano_referencia']}"

        # --- 4. Planilha consolidada do mês ---
        print("[main] Etapa 4/4 - Gerando planilha consolidada da frota...")
        nome_planilha = f"Resumo_Frota_{mes_ano_arquivo}.xlsx"
        caminho_planilha = config.PASTA_RELATORIOS_GERADOS / nome_planilha
        services.gerar_relatorio_consolidado_excel(lista_metricas, caminho_planilha)

        # --- Sucesso: move o CSV para a pasta de processados ---
        caminho_final = mover_arquivo(caminho_csv, config.PASTA_PROCESSADOS)
        print(f"[main] Concluído com sucesso. CSV movido para: {caminho_final}")
        return True

    except Exception as erro:
        # Qualquer falha em qualquer etapa acima cai aqui: registra o
        # erro de forma clara e move o CSV para a pasta de erros, sem
        # interromper o processamento dos demais arquivos da fila.
        print(f"[main] ERRO ao processar '{caminho_csv.name}': {erro}")
        try:
            caminho_erro = mover_arquivo(caminho_csv, config.PASTA_ERROS)
            print(f"[main] CSV movido para a pasta de erros: {caminho_erro}")
        except Exception as erro_ao_mover:
            print(f"[main] ERRO adicional ao tentar mover o CSV para 'erros/': {erro_ao_mover}")
        return False


# ---------------------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# ---------------------------------------------------------------------------

def main() -> None:
    """
    Executa o fluxo completo do sistema de controle de frota: inicializa
    o banco de dados e as pastas, processa todos os CSVs pendentes na
    pasta 'entrada/' e exibe um resumo final da execução.
    """
    inicio_execucao = datetime.now()

    print("=" * 70)
    print("  SISTEMA DE CONTROLE DE FROTA - ADMINISTRAÇÃO REGIONAL (GDF)")
    print("  Decreto nº 47.091/2025")
    print(f"  Início da execução: {inicio_execucao.strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 70)

    # --- Inicialização: banco de dados e estrutura de pastas ---
    print("\n[main] Inicializando banco de dados...")
    database.init_db()

    print("[main] Garantindo estrutura de pastas do projeto...")
    garantir_estrutura_de_pastas()

    # --- Varredura da pasta de entrada ---
    arquivos_csv = listar_arquivos_csv(config.PASTA_ENTRADA)

    if not arquivos_csv:
        print(f"\n[main] Nenhum arquivo .csv encontrado em '{config.PASTA_ENTRADA}'.")
        print("[main] Coloque o relatório mensal do Prime Benefícios nessa pasta e execute novamente.")
        return

    print(f"\n[main] {len(arquivos_csv)} arquivo(s) CSV encontrado(s) para processamento:")
    for arquivo in arquivos_csv:
        print(f"    - {arquivo.name}")

    # --- Processa cada CSV encontrado, um de cada vez ---
    arquivos_com_sucesso: List[str] = []
    arquivos_com_erro: List[str] = []

    for caminho_csv in arquivos_csv:
        sucesso = processar_arquivo_csv(caminho_csv)
        if sucesso:
            arquivos_com_sucesso.append(caminho_csv.name)
        else:
            arquivos_com_erro.append(caminho_csv.name)

    # --- Resumo final da execução ---
    duracao = (datetime.now() - inicio_execucao).total_seconds()

    print(f"\n{'=' * 70}")
    print("  RESUMO DA EXECUÇÃO")
    print(f"{'=' * 70}")
    print(f"  Arquivos processados com sucesso: {len(arquivos_com_sucesso)}")
    for nome in arquivos_com_sucesso:
        print(f"    [OK]   {nome}")
    print(f"  Arquivos com erro: {len(arquivos_com_erro)}")
    for nome in arquivos_com_erro:
        print(f"    [ERRO] {nome}")
    print(f"  Tempo total de execução: {duracao:.1f}s")
    print(f"  Fichas e planilhas disponíveis em: {config.PASTA_RELATORIOS_GERADOS}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()