from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from datetime import datetime, timedelta

# Inicia o navegador
driver = webdriver.Chrome()
driver.get("https://agendapib.pibcuritiba.org.br/detalhes/")

wait = WebDriverWait(driver, 15)

# Loop pelos próximos 7 dias
for i in range(7):
    data_atual = datetime.now() + timedelta(days=i)
    data_formatada = data_atual.strftime("%d/%m/%Y")
    data_arquivo = data_atual.strftime("%d-%m-%Y")

    try:
        # Espera os campos de data
        input_data_inicio = wait.until(EC.presence_of_element_located((By.ID, "dtAgendaI")))
        input_data_fim = wait.until(EC.presence_of_element_located((By.ID, "dtAgendaF")))

        # Preenche os dois campos com a mesma data
        input_data_inicio.clear()
        input_data_inicio.send_keys(data_formatada)

        input_data_fim.clear()
        input_data_fim.send_keys(data_formatada)

        # Clica no botão "Visualizar"
        botao = driver.find_element(By.CSS_SELECTOR, 'input.btn.btn-primary[type="submit"]')
        botao.click()

        # Aguarda a tabela aparecer
        wait.until(EC.presence_of_element_located((By.ID, "tableCalendar")))
        time.sleep(1)

        # Extrai os dados da tabela
        rows = driver.find_elements(By.CSS_SELECTOR, "#tableCalendar tbody tr")
        dados = []

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 8:
                Data = cols[0].text
                inicio = cols[1].text
                fim = cols[2].text
                evento = cols[3].text
                sala = cols[4].text
                area = cols[5].text
                som = cols[7].get_attribute("textContent").strip()

                # Só adiciona eventos com SOM preenchido
                if not som:
                    continue

                dados.append({
                    'Data': Data,
                    'Início': inicio,
                    'Fim': fim,
                    'Evento': evento,
                    'Sala': sala,
                    'Área': area,
                    'Som': som,
                })

        if dados:
            df = pd.DataFrame(dados)
            nome_arquivo = f'AGENDA_{data_arquivo}.xlsx'
            df.to_excel(nome_arquivo, index=False)
            print(f"Salvo: {nome_arquivo}")
        else:
            print(f"Aviso: Nenhum dado com SOM encontrado em {data_formatada}")

    except Exception as e:
        print(f"Erro ao processar {data_formatada}: {e}")

driver.quit()