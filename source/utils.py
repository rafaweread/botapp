import os
from datetime import datetime

def salvar_relatorio_em_pasta(data_bytes, start_date, base_dir='relatorios', nome_arquivo='relatorio.xlsx'):
    """
    Salva o conteúdo do relatório em uma estrutura de pastas ano/mês com base na data inicial.

    :param response_data: conteúdo do arquivo (response do requests)
    :param start_date: data de início no formato YYYY-MM-DD
    :param base_dir: diretório base onde as pastas serão criadas
    :param nome_arquivo: nome do arquivo a ser salvo
    """

    # Extrai ano e mês da data
    data = datetime.strptime(start_date, '%Y-%m-%d')
    ano = str(data.year)
    mes = str(data.month).zfill(2)

    # Cria caminho da pasta ano/mês
    pasta_destino = os.path.join(base_dir, ano, mes)
    os.makedirs(pasta_destino, exist_ok=True)

    # Caminho completo do arquivo
    caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)

    # Salva o conteúdo do arquivo
    with open(caminho_arquivo, 'wb') as f:
        f.write(data_bytes)

    print(f"Arquivo salvo em: {caminho_arquivo}")
