# -*- coding: utf-8 -*-
"""
Este script processa arquivos Excel de relatórios localizados em múltiplas subpastas.
Ele extrai dados de abas específicas, aplica transformações e mescla as informações
antes de exportar os resultados consolidados.
"""

import os
import pandas as pd
from typing import List, Tuple, Optional

def ler_dados_excel(pasta_principal: str) -> Tuple[List[pd.DataFrame], List[pd.DataFrame], List[pd.DataFrame]]:
    """
    Percorre as subpastas de um diretório principal, lê arquivos Excel e extrai
    dados das abas 'FILTROS', 'PDV' e 'CONSULTOR'.

    Args:
        pasta_principal (str): O caminho para a pasta que contém as subpastas com os relatórios.

    Returns:
        tuple: Uma tupla contendo três listas de DataFrames:
               - lista_dfs_filtros
               - lista_dfs_pdv
               - lista_dfs_consultor
    """
    lista_dfs_filtros = []
    lista_dfs_pdv = []
    lista_dfs_consultor = []

    print(f"Iniciando varredura na pasta: {pasta_principal}")
    for raiz, _, arquivos in os.walk(pasta_principal):
        for arquivo in arquivos:
            if arquivo.endswith(('.xlsx', '.xls')):
                caminho_completo = os.path.join(raiz, arquivo)
                nome_arquivo = os.path.splitext(arquivo)[0]
                subpasta = os.path.basename(raiz)
                print(f"\nProcessando arquivo: {caminho_completo}")

                # Tenta ler cada aba e adiciona à lista correspondente
                for aba, lista_destino in [('FILTROS', lista_dfs_filtros), 
                                           ('PDV', lista_dfs_pdv), 
                                           ('CONSULTOR', lista_dfs_consultor)]:
                    try:
                        df = pd.read_excel(caminho_completo, sheet_name=aba, keep_default_na=False)
                        df["Nome_Arquivo"] = nome_arquivo
                        df["Subpasta"] = subpasta
                        lista_destino.append(df)
                        print(f"  - Aba '{aba}' processada com sucesso.")
                    except Exception as e:
                        print(f"  - Aviso: Erro ao ler aba '{aba}' de {arquivo}: {e}")
    
    return lista_dfs_filtros, lista_dfs_pdv, lista_dfs_consultor

def concatenar_dataframes(lista_dfs: List[pd.DataFrame], nome_df: str) -> Optional[pd.DataFrame]:
    """
    Concatena uma lista de DataFrames em um único DataFrame.

    Args:
        lista_dfs (list): Lista de DataFrames a serem concatenados.
        nome_df (str): Nome do DataFrame para exibição em logs.

    Returns:
        pd.DataFrame or None: O DataFrame concatenado ou None se a lista estiver vazia.
    """
    if lista_dfs:
        df_final = pd.concat(lista_dfs, ignore_index=True)
        print(f"\nDataFrame '{nome_df}' final criado com {len(df_final)} linhas.")
        return df_final
    else:
        print(f"\nNenhum dado da aba '{nome_df}' foi encontrado para concatenar.")
        return None

def transformar_df_filtros(df_filtros: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    Aplica transformações específicas ao DataFrame de Filtros.
    - Filtra por 'PERÍODO ATUAL'.
    - Divide a coluna 'SELEÇÃO' em 'DTA INI' e 'DTA FIM'.

    Args:
        df_filtros (pd.DataFrame): O DataFrame de filtros concatenado.

    Returns:
        pd.DataFrame or None: O DataFrame transformado ou None se a entrada for inválida.
    """
    if df_filtros is None or 'FILTRO' not in df_filtros.columns:
        print("DataFrame de Filtros é inválido ou não contém a coluna 'FILTRO'. Pulando transformação.")
        return df_filtros

    print("\nIniciando transformação do DataFrame de Filtros...")
    df_filtros['FILTRO'] = df_filtros['FILTRO'].astype(str)
    df_filtrado = df_filtros.loc[df_filtros['FILTRO'] == 'PERÍODO ATUAL'].copy()

    if 'SELEÇÃO' in df_filtrado.columns:
        df_filtrado['SELEÇÃO'] = df_filtrado['SELEÇÃO'].astype(str)
        df_entre_datas = df_filtrado['SELEÇÃO'].str.split('-', expand=True)
        
        if df_entre_datas.shape[1] >= 2:
            df_entre_datas = df_entre_datas.iloc[:, :2]
            df_entre_datas.columns = ["DTA INI", "DTA FIM"]
            df_resultado = pd.concat(
                [df_filtrado.reset_index(drop=True), df_entre_datas.reset_index(drop=True)], axis=1
            ).drop(columns=["SELEÇÃO"])
            print("Transformação de Filtros concluída com sucesso.")
            return df_resultado
        else:
            print("Aviso: Divisão da coluna 'SELEÇÃO' não resultou em duas colunas.")
            return df_filtrado
    else:
        print("Aviso: Coluna 'SELEÇÃO' não encontrada. Nenhuma data foi separada.")
        return df_filtrado

def transformar_df_pdv(df_pdv: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    Aplica transformações específicas ao DataFrame de PDV.
    - Renomeia colunas.
    - Remove colunas desnecessárias e linhas de totais.

    Args:
        df_pdv (pd.DataFrame): O DataFrame de PDV concatenado.

    Returns:
        pd.DataFrame or None: O DataFrame transformado ou None se a entrada for inválida.
    """
    if df_pdv is None:
        print("DataFrame de PDV não existe. Pulando transformação.")
        return None

    print("\nIniciando transformação do DataFrame de PDV...")
    colunas_rename = {
        'Unnamed: 0': 'PDV', 'RECEITA (R$)': 'RECEITA PERÍODO ANTERIOR', 'Unnamed: 2': 'RECEITA PERÍODO ATUAL',
        'RECEITA MOBSHOP (R$)': 'RECEITA MOBSHOP PERÍODO ANTERIOR', 'Unnamed: 5': 'RECEITA MOBSHOP PERÍODO ATUAL',
        'BOLETO MÉDIO ': 'BOLETO MÉDIO PERÍODO ANTERIOR', 'Unnamed: 8': 'BOLETO MÉDIO PERÍODO ATUAL',
        'BOLETO MÉDIO MOBSHOP': 'BOLETO MÉDIO MOBSHOP PERÍODO ANTERIOR', 'Unnamed: 11': 'BOLETO MÉDIO MOBSHOP PERÍODO ATUAL',
        'ITENS POR BOLETO': 'ITENS POR BOLETO PERÍODO ANTERIOR', 'Unnamed: 14': 'ITENS POR BOLETO PERÍODO ATUAL',
        'QUANTIDADE DE BOLETOS': 'QUANTIDADE DE BOLETOS PERÍODO ANTERIOR', 'Unnamed: 17': 'QUANTIDADE DE BOLETOS PERÍODO ATUAL',
        'PREÇO MÉDIO': 'PREÇO MÉDIO PERÍODO ANTERIOR', 'Unnamed: 20': 'PREÇO MÉDIO PERÍODO ATUAL',
        'QUANTIDADE DE ITENS': 'QUANTIDADE DE ITENS PERÍODO ANTERIOR', 'Unnamed: 23': 'QUANTIDADE DE ITENS PERÍODO ATUAL'
    }
    df_pdv_renomeado = df_pdv.rename(columns=colunas_rename)

    colunas_remover = ['Unnamed: 3', 'Unnamed: 6', 'Unnamed: 9', 'Unnamed: 12', 'Unnamed: 15', 'Unnamed: 18', 'Unnamed: 21', 'Unnamed: 24', 'Subpasta']
    df_pdv_limpo = df_pdv_renomeado.drop(columns=[col for col in colunas_remover if col in df_pdv_renomeado.columns])

    df_pdv_limpo = df_pdv_limpo[~df_pdv_limpo['PDV'].isin(['PDV', 'TOTAL'])]
    print("Transformação de PDV concluída com sucesso.")
    return df_pdv_limpo

def mesclar_dados(df_base: pd.DataFrame, df_pdv: Optional[pd.DataFrame], df_consultor: Optional[pd.DataFrame]) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """
    Mescla o DataFrame base (derivado de filtros) com os DataFrames de PDV e Consultor.

    Args:
        df_base (pd.DataFrame): DataFrame principal transformado.
        df_pdv (pd.DataFrame): DataFrame de PDV transformado.
        df_consultor (pd.DataFrame): DataFrame de Consultor concatenado.

    Returns:
        tuple: Uma tupla contendo os dois DataFrames mesclados (df_merge_pdv, df_merge_consultor).
    """
    if df_base is None:
        print("\nDataFrame base não existe. Impossível realizar o merge.")
        return None, None
        
    print("\nIniciando merge dos dados...")
    df_merge_pdv = None
    if df_pdv is not None:
        df_merge_pdv = pd.merge(df_base, df_pdv, on="Nome_Arquivo", how="inner")
        print(f"Merge com PDV concluído. Resultado: {len(df_merge_pdv)} linhas.")

    df_merge_consultor = None
    if df_consultor is not None:
        df_merge_consultor = pd.merge(df_base, df_consultor, on="Nome_Arquivo", how="inner")
        print(f"Merge com Consultor concluído. Resultado: {len(df_merge_consultor)} linhas.")
        
    return df_merge_pdv, df_merge_consultor

def exportar_resultados(df_final: Optional[pd.DataFrame], nome_arquivo: str, nome_aba: str):
    """
    Exporta um DataFrame para um arquivo Excel.

    Args:
        df_final (pd.DataFrame): O DataFrame a ser exportado.
        nome_arquivo (str): O nome do arquivo de saída (ex: 'output.xlsx').
        nome_aba (str): O nome da aba na planilha Excel.
    """
    if df_final is not None and not df_final.empty:
        try:
            df_final.to_excel(nome_arquivo, sheet_name=nome_aba, engine='openpyxl', index=False)
            print(f"\nArquivo '{nome_arquivo}' criado com sucesso na aba '{nome_aba}'!")
        except Exception as e:
            print(f"\nErro ao exportar o arquivo '{nome_arquivo}': {e}")
    else:
        print(f"\nNenhum dado para exportar para o arquivo '{nome_arquivo}'.")

def main():
    """
    Função principal que orquestra todo o processo de ETL.
    """
    # 1. Definição do Caminho da Pasta Principal
    pasta_principal = r'.\relatorios\2025'
    
    # 2. Leitura e Extração dos Dados
    listas_filtros, listas_pdv, listas_consultor = ler_dados_excel(pasta_principal)

    # 3. Concatenação dos Dados
    df_filtros_raw = concatenar_dataframes(listas_filtros, "FILTROS")
    df_pdv_raw = concatenar_dataframes(listas_pdv, "PDV")
    df_consultor_raw = concatenar_dataframes(listas_consultor, "CONSULTOR")

    # 4. Transformação dos Dados
    df_base_transformado = transformar_df_filtros(df_filtros_raw)
    df_pdv_transformado = transformar_df_pdv(df_pdv_raw)

    # 5. Mesclagem (Merge) dos Dados
    df_final_pdv, df_final_consultor = mesclar_dados(df_base_transformado, df_pdv_transformado, df_consultor_raw)
    
    # 6. Exportação dos Resultados
    exportar_resultados(df_final_pdv, 'output_pdv.xlsx', 'PDV')
    exportar_resultados(df_final_consultor, 'output_consultor.xlsx', 'CONSULTOR')

    print("\nProcessamento concluído.")

# Ponto de entrada do script: a função main() só será executada
# quando o script for rodado diretamente.
if __name__ == "__main__":
    main()