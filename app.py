from email.mime.base import MIMEBase
import shutil
from flask import Flask, request, render_template, send_file
import pandas as pd
import openai
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from io import BytesIO
from dotenv import load_dotenv
import tempfile
from flask import session

# Carrega as variáveis de ambiente a partir do arquivo .env
load_dotenv()

app = Flask(__name__)

# Configura a chave API da OpenAI
openai.api_key = os.getenv("OPENAI_API_KEY")
EMAIL = os.environ.get('EMAIL')
PASSWORD = os.environ.get('PASSWORD')

@app.route("/home")
def index():
  return render_template("home.html")

@app.route("/sobremim")
def sobremim():
  return render_template("sobremim.html")

@app.route ("/portfolio")
def porfolio():
  return render_template("portfolio.html")

@app.route("/contato")
def contato():
  return render_template("contato.html")

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    uploaded_file = request.files['file']
    if uploaded_file.filename != '':

            # Cria um diretório temporário
            temp_dir = tempfile.mkdtemp()

            # Salva o arquivo enviado pelo usuário no diretório temporário
            caminho_arquivo = os.path.join(temp_dir, uploaded_file.filename)
            uploaded_file.save(caminho_arquivo)

            # Carrega o DataFrame do arquivo Excel
            df = pd.read_excel(caminho_arquivo, engine='openpyxl')

            coluna_pergunta = 'PERGUNTA'
            colunas_resposta = [col for col in df.columns if col != coluna_pergunta]

            primeiras_linhas = df.head(3).to_html()

            # Preparar uma lista para armazenar os resultados
            analysis_results = []

            # Iterar sobre cada pergunta no DataFrame
            for index, row in df.iterrows():
                question = row[coluna_pergunta]
                # Compilar todas as respostas dos entrevistados em uma única string
                all_answers = ". ".join([str(row[col]) for col in colunas_resposta if pd.notna(row[col])])

                if all_answers:
                    # Criar o prompt para a análise
                    prompt = f"Por favor, forneça uma análise geral e: Essas foram perguntas realizadas numa pesquisa qualitativa de diagnóstico de diversidade: '{question}', e esse é um compilado das respostas dos entrevistados ouvidos: {all_answers}. Por favor, forneça uma análise:. Você deve agir como consultor de diversidade.Eu preciso que você avalie o registro dessas entrevistas. Existe uma coluna com as perguntas, e cada uma das colunas seguintes representam um entrevistado. Analise essas respostas para me trazer uma percepção média, por pergunta. Tenha em mente que essa análise irá compor um relatório que resume essa etapa. Eu quero que as análises integrem todas as respostas dos entrevistados,cerca de 100 palavras. e destaque o que demostra algum padrão. Não quero bullets. Você pode citar parte das falas para ilustrar, com aspas. Lembrando que tudo que está nas respostas é transcrição das entrevistas.Pontos importantes: Você não precisa de uma introdução da sua análise, pode ir direto ao ponto, foque em trazer informações relevantes da análise, você pode usar o padrão de texto: Analise, dois pontos, e sua analise a seguir, sem mencionar textos como:Analisando as respostas fornecidas pelos entrevistados, que torna o retorno prolixo e pouco produtivo. Você não precisa utilizar linguagem rebuscada."

                    # Realizar a chamada à API da OpenAI para obter o resumo analisado
                    try:
                        response = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",  # Utiliza o modelo otimizado para chat
                            messages=[
                                {"role": "system", "content": "Você é um assistente de IA."},
                                {"role": "user", "content": prompt}
                            ]
                        )
                        analysis = response['choices'][0]['message']['content']
                    except Exception as e:
                        print(f"Erro ao realizar chamada à API da OpenAI: {e}")
                        analysis = "Erro na análise."

                    # Adicionar o resumo analisado e a pergunta ao resultado
                    analysis_results.append({'Pergunta': question, 'Análise': analysis})
                            # Salvar os resultados em um arquivo temporário

            caminho_arquivo_resultados = os.path.join(temp_dir, 'resultados_analise.xlsx')
            df_resultados = pd.DataFrame(analysis_results)
            df_resultados.to_excel(caminho_arquivo_resultados, index=False)

            # Retorna os resultados para uma nova página HTML
            return render_template('resultados_analise.html', primeiras_linhas=primeiras_linhas, caminho_arquivo_resultados=caminho_arquivo_resultados, analysis_results=analysis_results)


    else:
        return render_template('upload.html')


@app.route('/envio_email', methods=['POST'])
def send_email():
    try:
        # Definir o diretório temporário para salvar o arquivo de resultados
        temp_dir = tempfile.mkdtemp()
        caminho_arquivo_resultados = request.form['caminho_arquivo_resultados']


        # Dados para o e-mail
        smtp_server = "smtp-relay.brevo.com"
        port = 587
        remetente = EMAIL
        destinatario = request.form['destinatario']
        titulo = "Resultados da Análise Qualitativa"
        texto = "Segue em anexo os resultados da análise."

        # Salvar os resultados da análise em um arquivo Excel no diretório temporário

        # Iniciar conexão com o servidor SMTP do Brevo
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()
        server.login(EMAIL, PASSWORD)

        # Preparar a mensagem de e-mail
        mensagem = MIMEMultipart()
        mensagem["From"] = remetente
        mensagem["To"] = destinatario
        mensagem["Subject"] = titulo
        mensagem.attach(MIMEText(texto, "plain"))

        # Anexar o arquivo de resultados
        with open(caminho_arquivo_resultados, 'rb') as arquivo:
            anexo = MIMEBase('application', 'octet-stream')
            anexo.set_payload(arquivo.read())
        encoders.encode_base64(anexo)
        anexo.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_arquivo_resultados)}')
        mensagem.attach(anexo)

        # Enviar e-mail
        server.sendmail(remetente, destinatario, mensagem.as_string())

        # Fechar conexão com o servidor SMTP
        server.quit()

        # Remover o diretório temporário após o envio do e-mail
        shutil.rmtree(temp_dir)

        # Redirecionar de volta para a página de resultados
        return render_template('upload.html', enviado=True)
    except Exception as e:
        return render_template('upload.html', erro=str(e))
if __name__ == '__main__':
    app.run(debug=True)