import requests
import json
import time
import pandas as pd
import openpyxl


def consultar_cnpj(codigo, cnpj):
    # Site para consultar informações do CNPJ
    url = f'https://www.receitaws.com.br/v1/cnpj/{cnpj}'

    try:
        response = requests.get(url)
        data = response.json()

        if response.status_code == 200:
            # Processar os dados do CNPJ aqui
            try:
                resposta = [codigo, data['nome'], data['abertura'], data['situacao'], data['logradouro'],
                            data['bairro'], data['numero'], data['municipio'], data['uf'], data['status']]
            except KeyError:
                resposta = None
            return resposta
        else:
            print(f'Erro na consulta: {data["message"]}')
    except requests.exceptions.RequestException as e:
        print(f'Erro na requisição: {str(e)}')
        return None


# Onde está a base com os CNPJ's a serem consultados
base = pd.read_csv('cnpj.csv', dtype='str')
cnpjs = base['CNPJ']

lista = []
contador = 0
acompanhar = 1
for i in cnpjs:
    i = str(i).zfill(14)
    print(f"{acompanhar}° CNPJ: {i}\n")
    resultado = consultar_cnpj(i, i)
    if resultado is not None:  # Verificar se o resultado não é None antes de adicionar à lista
        lista.append(resultado)
    acompanhar += 1
    contador += 1
    if contador == 3:
        print("Limite de consultas por minuto -> Esperando...\n")
        time.sleep(61)
        contador = 0
# Alterando o nome das colunas e salvando
nome_colunas = ['CNPJ', 'Nome', 'Data de Abertura', 'Situção', 'Endereço', 'Bairro', 'Número', 'Cidade', 'UF', 'Status']
arquivo = pd.DataFrame(lista, columns=nome_colunas)
caminho_arquivo = r"C:\Users\Public\Documentos\01.Scripts\01.C-Stores\InfoCstore.xlsx"
arquivo.to_excel(caminho_arquivo, index=False, engine='openpyxl')
# Verificar se o diretório existe e criar se necessário

