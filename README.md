<h2>Finalização em Massa de Atendimentos no Sistema ERP IXC-SOFT</h2>

Este projeto tem como objetivo automatizar a finalização de atendimentos em massa no sistema ERP IXC-SOFT, utilizando a API do sistema. O script em Python faz requisições a partir de uma planilha Excel contendo os IDs dos atendimentos a serem finalizados.

<h2>Bibliotecas</h2>
<ul>
  <li>Requests</li>
  <li>Base64</li>
  <li>Json</li>
  <li>openpyxl</li>
  <li>datetime</li>
</ul>

<h2>Estrutura do Projeto</h2>
<ul>
  <li>planilha.xlsx: Arquivo Excel contendo os IDs dos atendimentos a serem finalizados.</li>
  <li>main.py: Script Python para finalizar os atendimentos em massa.</li>
</ul>

<h2>Como Configurar</h2>
<ol>
  <li>Clone o Repositório:
    <pre><code>git clone https://github.com/LucasMaciel7/Finalizar_atendimentos_IXC-API.git</code></pre>
  </li>
  <li>Instale as bibliotecas necessárias:
    <pre><code>pip install requests openpyxl</code></pre>
  </li>
  <li>Configure o script:
    <ul>
      <li>Adicione o host da API no campo <code>host</code>.</li>
      <li>Adicione o token de autenticação no campo <code>token</code>.</li>
    </ul>
  </li>
</ol>

<h2>Estrutura do Script</h2>

<h3>Importação de Bibliotecas</h3>

<pre><code>import requests
import base64
import json
import openpyxl
from datetime import datetime
import time
</code></pre>

<h3>Inicialização de Variáveis</h3>

<pre><code>atendimentos = []
response_count = 0
data_hora_atual = datetime.now().isoformat()
</code></pre>

<h3>Carregamento da Planilha</h3>

<pre><code>planilha = openpyxl.load_workbook('planilha.xlsx')
planilha_ativa = planilha.active
</code></pre>

<h3>Definição do Host, URL e Token de Autenticação</h3>

<pre><code>host = ""  # Adicione o seu Host!
url = "https://{}/webservice/v1/su_mensagens/{}".format(host, atendimentos)
token = "".encode('utf-8')  # Token de autenticação
</code></pre>

<h3>Exibição do Horário de Inicialização</h3>

<pre><code>print(f"Horario da inicialização {data_hora_atual}/n")
print("-" * 50)
</code></pre>

<h3>Extração dos IDs dos Atendimentos</h3>

<pre><code>for linha in planilha_ativa.iter_rows(min_row=2, values_only=True):
    if linha[0] is not None:
        atendimentos.append(linha[0])
</code></pre>

<h3>Criação das Requisições</h3>

<pre><code>for id in atendimentos:
    payload = {
        "id_ticket": id,
        "data": "CURRENT_TIMESTAMP",
        'mensagens_nao_lida_cli': 'Finalizado via API',
        "operador": "",
        'su_status': 'S',
        "mensagem": "Finalização de atendimentos em massa via API",
        "visibilidade_mensagens": "PU",
        "existe_pendencia_externa": "0",
        "id_evento_status": "0",
        "ultima_atualizacao": data_hora_atual
    }

    headers = {
        'Authorization': 'Basic {}'.format(base64.b64encode(token).decode('utf-8')),
        'Content-Type': 'application/json'
    }

    while True:
        try:
            response = requests.post(url, data=json.dumps(payload), headers=headers, timeout=30)
            if response.status_code == 200:
                response_count += 1

                if response_count == 1000:
                    response_count = 0
                    print("1000 Atendimentos finalizados, realizando pausa de 5 minutos")
                    time.sleep(300)
                break

            else:
                print('Resposta da API não foi 200, realizando pausa de 5 minutos')
                time.sleep(300)

        except requests.exceptions.Timeout:
            print('Tempo limite de conexão excedido, realizando pausa de 5 minutos')
            time.sleep(300)

        except requests.exceptions.RequestException as e:
            print('Erro ao fazer requisição:', e)
            time.sleep(300)
</code></pre>

<h3>Exibição do Horário de Finalização</h3>

<pre><code>print(response.text)
print("-" * 50)
print(f"\nHorario da finalização ", datetime.now().isoformat())
print('\nComando enviado com SUCESSO, FINALIZADO!')
</code></pre>

<h2>Execução do Script</h2>

<ol>
  <li>Execute o script:
    <pre><code>python main.py</code></pre>
  </li>
  <li>Monitore a execução:
    <ul>
      <li>O script exibirá mensagens no console indicando o progresso e pausas realizadas durante a execução.</li>
    </ul>
  </li>
</ol>

<h2>Contribuições</h2>

Sinta-se à vontade para contribuir com melhorias e novas funcionalidades. Abra uma issue ou envie um pull request com suas sugestões.

<h2>Licença</h2>

Este projeto está licenciado sob a licença MIT. Veja o arquivo <a href="LICENSE">LICENSE</a> para mais detalhes.
