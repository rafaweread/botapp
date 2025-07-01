from source.utils import salvar_relatorio_em_pasta
from source.reports import ReportClient
from calendar import monthrange

if __name__ == "__main__":

    client = ReportClient() 
    ano = 2025
    for mes in range(1, 13):
        # Obtém o número de dias no mês
        _, num_dias = monthrange(ano, mes)
        start_date = f"{ano}-{mes:02d}-01"
        end_date = f"{ano}-{mes:02d}-{num_dias}"
        print(f'Gerando relatório para o mês: {start_date}')
        for dia in range(1, num_dias + 1):
            # Formata a data para o dia atual
            start_date = f"{ano}-{mes:02d}-{dia:02d}"
            end_date = f"{ano}-{mes:02d}-{dia:02d}"
            print(f'Gerando relatório para o dia: {start_date}')
            try:
                # Gera o relatório
                bin_data = client.get_report(start_date=start_date, end_date=end_date)
                # Salva o relatório na pasta
                salvar_relatorio_em_pasta(bin_data, start_date=start_date, base_dir='relatorios', nome_arquivo=f'{start_date}.xlsx')
            except Exception as e:
                print(f'Erro ao gerar relatório para {start_date}: {e}')


#start_date = f"2025-01-01"
#end_date = f"2025-01-31"
