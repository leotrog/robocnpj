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
                atividade_principal = data['atividade_principal'][0]['text'] if 'atividade_principal' in data and data['atividade_principal'] else 'N/A'
                qsa = data.get('qsa', [])
                if qsa:
                    qsa_info = qsa[0]  # Pegando o primeiro item da lista qsa
                else:
                    qsa_info = {'nome': 'N/A', 'qual': 'N/A', 'pais_origem': 'N/A', 'nome_rep_legal': 'N/A', 'qual_rep_legal': 'N/A'}
                resposta = [codigo, data.get('nome', 'N/A'), data.get('abertura', 'N/A'), data.get('situacao', 'N/A'), data.get('logradouro', 'N/A'),
                            data.get('bairro', 'N/A'), data.get('numero', 'N/A'), data.get('municipio', 'N/A'), data.get('uf', 'N/A'), data.get('status', 'N/A'), data.get('cep', 'N/A'), data.get('complemento', 'N/A'), data.get('email', 'N/A'), data.get('telefone', 'N/A'), atividade_principal,
                            qsa_info.get('nome', 'N/A'), qsa_info.get('qual', 'N/A'), qsa_info.get('pais_origem', 'N/A'), qsa_info.get('nome_rep_legal', 'N/A'), qsa_info.get('qual_rep_legal', 'N/A')]
            except KeyError:
                resposta = [codigo, data.get('nome', 'N/A'), data.get('abertura', 'N/A'), data.get('situacao', 'N/A'), data.get('logradouro', 'N/A'),
                            data.get('bairro', 'N/A'), data.get('numero', 'N/A'), data.get('municipio', 'N/A'), data.get('uf', 'N/A'), data.get('status', 'N/A'), data.get('cep', 'N/A'), data.get('complemento', 'N/A'), data.get('email', 'N/A'), data.get('telefone', 'N/A'), atividade_principal,
                            'N/A', 'N/A', 'N/A', 'N/A', 'N/A']
            return resposta
        else:
            print(f'Erro na consulta: {data["message"]}')
            return [codigo, 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A']
    except requests.exceptions.RequestException as e:
        print(f'Erro na requisição: {str(e)}')
        return [codigo, 'N/A', 'N/A', 'N/A', 'N/A']

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
nome_colunas = ['CNPJ', 'Nome', 'Data de Abertura', 'Situação', 'Endereço', 'Bairro', 'Número', 'Cidade', 'UF', 'Status', 'cep', 'complemento', 'email', 'telefone', 'Atividade Principal',
                'QSA Nome', 'QSA Qual', 'QSA País de Origem', 'QSA Nome Rep Legal', 'QSA Qual Rep Legal']
arquivo = pd.DataFrame(lista, columns=nome_colunas)
caminho_arquivo = r"C:\Users\Public\Documentos\01.Scripts\01.C-Stores\InfoCstore.xlsx"
arquivo.to_excel(caminho_arquivo, index=False, engine='openpyxl')
