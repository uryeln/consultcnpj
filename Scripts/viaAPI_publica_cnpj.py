import requests
import pandas as pd
import time
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

df = pd.read_excel('.\\Search\\all.xlsx')

resultados = []

# Função para buscar os dados do CNPJ
def buscar_cnpj(cnpj, max_tentativas=3, timeout=10):
    for tentativa in range(max_tentativas):
        try:
            response = requests.get(f"https://publica.cnpj.ws/cnpj/{cnpj}", timeout=timeout)
            if response.status_code == 429:
                logging.warning("Limite de solicitações excedido. Aguardando 60 segundos para tentar novamente...")
                time.sleep(60)
                continue
            elif response.status_code == 200:
                return response.json()
            elif response.status_code == 504:
                logging.warning(f"Timeout ao buscar dados para {cnpj}. Tentando novamente após 5 segundos...")
                time.sleep(5)
                continue
            else:
                raise requests.exceptions.HTTPError(f"Erro ao buscar dados para {cnpj}: {response.status_code}")
        except requests.exceptions.RequestException as e:
            logging.error(f"Erro de conexão ao buscar dados para {cnpj}: {e}")
            time.sleep(10)
            continue
        except Exception as e:
            logging.error(f"Erro ao buscar dados para {cnpj}: {e}")
            return None

    logging.error(f"Número máximo de tentativas ({max_tentativas}) atingido para CNPJ {cnpj}.")
    return None

# Início da busca dos CNPJs
logging.info("Iniciando busca dos CNPJs...")
for index, row in df.iterrows():
    cnpj = row['CNPJ']
    logging.info(f"Buscando dados para: {cnpj} ({index + 1}/{len(df)})")

    data = buscar_cnpj(cnpj)
    logging.info(f"Resposta da API para {cnpj}: {data}")

    if data:
        razao_social = data.get('razao_social', 'n/a')
        cidade = data.get('estabelecimento', {}).get('cidade', {}).get('nome', 'n/a')
        estado = data.get('estabelecimento', {}).get('estado', {}).get('sigla', 'n/a')
        tipo = data.get('estabelecimento', {}).get('tipo', 'n/a')
        atividade = data.get('estabelecimento', {}).get('atividade_principal', {}).get('descricao', 'n/a')

        logging.info(f"Razão Social: {razao_social}")
        logging.info(f"Cidade: {cidade}")
        logging.info(f"Estado: {estado}")
        logging.info(f"Tipo: {tipo}")
        logging.info(f"Atividade: {atividade}")

        resultados.append({
            'CNPJ': cnpj,
            'Razão Social': razao_social,
            'Cidade': cidade,
            'Estado': estado,
            'Tipo': tipo,
            'Atividade': atividade
        })
    else:
        resultados.append({
            'CNPJ': cnpj,
            'Razão Social': "n/a",
            'Cidade': "n/a",
            'Estado': "n/a",
            'Tipo': "n/a",
            'Atividade': "n/a"
        })

    logging.info(f"Dados para {cnpj} obtidos com sucesso.")
    time.sleep(20)

df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel('.\\Result\\resultados_all.xlsx', index=False)
logging.info("Processo concluído. Os resultados foram salvos em 'resultados_cnpjs.xlsx'.")
