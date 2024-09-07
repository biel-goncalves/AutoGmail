import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Escopos necessários para leitura
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

def main():
    """Mostra como usar a API do Google Sheets para ler dados e enviar um alerta por e-mail."""
    creds = None
    # Verifica se o arquivo token.json existe
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # Se não há credenciais válidas, realiza o fluxo de autenticação
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "client_secret.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Salva as credenciais para o próximo uso
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        # Constrói o serviço da API
        service = build("sheets", "v4", credentials=creds)

        # Obtém os dados da planilha
        sheet = service.spreadsheets()
        result_estoque = (
            sheet.values()
            .get(spreadsheetId="1TRRcioYisDrhTVMgJV-evNAW8j-5u9QWjYT5t2U9zDg", range="2024!E2:E55")  # Coluna de estoque
            .execute()
        )
        result_nomes = (
            sheet.values()
            .get(spreadsheetId="1TRRcioYisDrhTVMgJV-evNAW8j-5u9QWjYT5t2U9zDg", range="2024!A2:A55")  # Coluna de nomes
            .execute()
        )

        valores_estoque = result_estoque.get("values", [])
        valores_nomes = result_nomes.get("values", [])

        produtos_baixos = []

        # Usa o comprimento mínimo das listas para evitar IndexError
        num_estoque = len(valores_estoque)
        num_nomes = len(valores_nomes)
        for i in range(min(num_estoque, num_nomes)):
            if valores_estoque[i] and valores_nomes[i]:
                try:
                    nome_produto = valores_nomes[i][0].strip()
                    estoque_str = valores_estoque[i][0].strip()
                    
                    # Remove caracteres não numéricos e converte para inteiro
                    estoque = int(''.join(filter(str.isdigit, estoque_str)))

                    if estoque < 3:
                        produtos_baixos.append(f"Produto: {nome_produto}, Estoque: {estoque}")
                except ValueError:
                    print(f"Valor '{valores_estoque[i][0]}' não é um número válido.")
        
        # Cria o corpo do e-mail com ou sem produtos baixos
        corpo_email = f"""<div style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #000000;">Aviso Urgente: <span style="color: #000000;">Estoque Baixo</span></h2>
        <hr style="border: 0; border-top: 1px solid #000000;">
        
        <p style="font-size: 16px;">Prezado(a),</p>

        <p style="font-size: 16px;">
        <em>Espero que esteja bem.</em>
        </p>

        <p style="font-size: 16px;">
            Gostaria de alertá-lo sobre a situação atual do estoque. 
            Identificamos que os seguintes produtos estão com níveis de estoque <strong style="color: #FF0000;">críticos</strong>:
        </p>

        <ul style="font-size: 16px;">
        {"".join([f"<li><strong>{produto}</strong></li>" for produto in produtos_baixos]) or "<li>Nenhum produto com estoque baixo.</li>"}
        </ul>

        <p style="font-size: 16px;">
            <strong>Recomendamos</strong> que uma verificação seja feita o quanto antes e que sejam tomadas as medidas necessárias para evitar a falta desses itens.
        </p>

        <p style="font-size: 16px;">
            Por favor, não hesite em me contatar caso precise de mais informações ou assistência.
        </p>

        <hr style="border: 0; border-top: 1px solid #000000;">

        <p style="font-size: 16px;">
            Atenciosamente,<br>
            <strong>Gabriel Gonçalves Santana</strong><br>
            Setor de Compras<br>
            <span style="color: #228B22;">UNIFAG</span><br>
            <a href="mailto:gabriel.goncalves@unifag.com.br" style="color: #003366;">gabriel.goncalves@unifag.com.br</a><br>
            Telefone: (11) 91572-4833
        </p>
        </div>
        """

        # Configura o e-mail
        msg = MIMEMultipart()
        msg['Subject'] = "Aviso Urgente: Estoque Baixo"
        msg['From'] = ""
        msg['To'] = ""
        password = ""  # Usar variável de ambiente para a senha

        # Adiciona o corpo do e-mail
        msg.attach(MIMEText(corpo_email, 'html'))

        try:
            # Conecta ao servidor SMTP
            with smtplib.SMTP('smtp.gmail.com', 587) as s:
                s.starttls()
                # Login Credentials for sending the mail
                s.login(msg['From'], password)
                s.sendmail(msg['From'], [msg['To']], msg.as_string())
            print("Email Enviado")
        except Exception as e:
            print(f"Falha ao enviar o e-mail: {e}")

    except HttpError as err:
        print(f"Erro ao acessar a API: {err}")

if __name__ == "__main__":
    main()
