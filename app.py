
import requests
import base64
import json 
import openpyxl
from datetime import datetime 
import time

# Carrega a base de dados em uma Planilha XLSX
atendimentos = []
cont_requiscoes = 0
data_hora_atual = datetime.now().isoformat()
planilha = openpyxl.load_workbook('') # Adicionar base de dados 
planilha_ativa = planilha.active

# Realiza a Conexão
host = ""
url = "https://{}/webservice/v1/su_mensagens/{}".format(host,atendimentos)
token = "".encode('utf-8')


print(f"Horario da inicialização {data_hora_atual}/n")
print("-"*50)

# Retira os IDs da planilha Exel e adiciona dentro da lista
for linha in planilha_ativa.iter_rows(min_row=2,values_only=True):
    if linha[0] is not None:
        atendimentos.append(linha[0])
        


# Realiza o envio do Json para cada ID
for id in atendimentos:
    payload = {
        "id_ticket": id,
        "data":"CURRENT_TIMESTAMP",
        'mensagens_nao_lida_cli': 'Finalizado via API',
        "operador": "operador desejado",
        'su_status': 'S',
        "mensagem": "Finalização de atendimentos em massa ",
        "visibilidade_mensagens": "PU",
        "existe_pendencia_externa": "0",
        "id_evento_status": "0",
        "ultima_atualizacao": data_hora_atual
    }


    headers = {
            'Authorization': 'Basic {}'.format(base64.b64encode(token).decode('utf-8')),
            'Content-Type': 'application/json'
            }
    
    response = requests.post(url, data=json.dumps(payload), headers=headers)

    # A cada 1000 Requisições realiza uma pausa para não sobrecarregar o servidor e parar o codigo
    cont_requiscoes += 1
    if cont_requiscoes == 1000:
        cont_requiscoes == 0 
        time.sleep(180)
        pass
    
print(response.text)

print("-"*50)
print(f"\nHorario da abertura ", datetime.now().isoformat())
print('\nComando enviado com SUCESSO, FINALIZADO!')