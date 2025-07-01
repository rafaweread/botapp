"""
import time
import schedule
from datetime import datetime

from source.reports import ReportClient
from source.utils import salvar_relatorio_em_pasta

# -------------------------------------------------------------------
def run_report() -> None:
    # Baixa e salva o relatório do dia atual.
    client = ReportClient()
    date_str = datetime.now().strftime("%Y-%m-%d")
    print(f"[{datetime.now():%H:%M:%S}] Gerando relatório de {date_str}")

    try:
        bin_data = client.get_report(start_date=date_str, end_date=date_str)
        salvar_relatorio_em_pasta(
            bin_data,
            start_date=date_str,
            base_dir="relatorios",
            nome_arquivo=f"{date_str}.xlsx",
        )
        print(f"[OK] Relatório salvo em relatorios/{date_str}.xlsx")
    except Exception as e:
        print(f"[ERRO] Falhou em {date_str}: {e}")

# -------------------------------------------------------------------

if __name__ == "__main__":
    # horários desejados (24 h)
    for hr in ("06:00", "11:30", "17:49", "23:00"):
        schedule.every().day.at(hr).do(run_report)
        print(f"✓ Agendado para {hr}")

    # loop infinito
    while True:
        schedule.run_pending()
        time.sleep(30)        # checa a fila a cada 30 s


if __name__ == "__main__":
    run_report()

"""

import time
import schedule
from datetime import datetime, timedelta

from source.reports import ReportClient
from source.utils import salvar_relatorio_em_pasta

# -------------------------------------------------------------------
def run_report() -> None:
    """Baixa e salva o relatório do dia atual."""
    client = ReportClient()
    date_str = datetime.now().strftime("%Y-%m-%d")
    print(f"[{datetime.now():%H:%M:%S}] Gerando relatório de {date_str}")

    try:
        bin_data = client.get_report(start_date=date_str, end_date=date_str)
        salvar_relatorio_em_pasta(
            bin_data,
            start_date=date_str,
            base_dir="relatorios",
            nome_arquivo=f"{date_str}.xlsx",
        )
        print(f"[OK] Relatório salvo em relatorios/{date_str}.xlsx")
    except Exception as e:
        print(f"[ERRO] Falhou em {date_str}: {e}")

    date_str = datetime.now() - timedelta(days=1)
    date_str = date_str.strftime("%Y-%m-%d")

    try:
        bin_data = client.get_report(start_date=date_str, end_date=date_str)
        salvar_relatorio_em_pasta(
            bin_data,
            start_date=date_str,
            base_dir="relatorios",
            nome_arquivo=f"{date_str}.xlsx",
        )
        print(f"[OK] Relatório salvo em relatorios/{date_str}.xlsx")
    except Exception as e:
        print(f"[ERRO] Falhou em {date_str}: {e}")

# -------------------------------------------------------------------

"""
if _name_ == "_main_":
    # horários desejados (24 h)
    for hr in ("06:00", "11:30", "17:00", "23:00"):
        schedule.every().day.at(hr).do(run_report)
        print(f"✓ Agendado para {hr}")

    # loop infinito
    while True:
        schedule.run_pending()
        time.sleep(30)        # checa a fila a cada 30 s
"""
if __name__ == "__main__":
    run_report()