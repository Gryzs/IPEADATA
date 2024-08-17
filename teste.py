import requests
from bs4 import BeautifulSoup
import csv

headers = {
    'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
}

url = 'https://www.dadosdemercado.com.br/acoes'

# Realizar a requisição ao site
site = requests.get(url, headers=headers)
status = site.status_code

if status == 200:
    soup = BeautifulSoup(site.content, 'html.parser')
    table = soup.find('table', id='stocks')

    if table:
        # Criar ou abrir o arquivo CSV para escrita
        with open('codigos.csv', mode='w', newline='') as file:
            writer = csv.writer(file)

            # Encontrar todas as linhas da tabela
            rows = table.find_all('tr')

            for row in rows[1:]:  # Pular o cabeçalho da tabela
                ticker_cell = row.find_all('td')[0]  # Pegar a primeira célula (código do ticker)
                ticker = ticker_cell.get_text(strip=True)

                # Escrever o código no arquivo CSV
                writer.writerow([ticker])

        print("Códigos salvos com sucesso em 'codigos.csv'.")
    else:
        print("Tabela não encontrada no site.")
else:
    print(f"Erro ao acessar o site. Status code: {status}")
