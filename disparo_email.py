import os
from email.message import EmailMessage
import ssl
import smtplib
import pandas as pd 
import time 


#ler excel
planilha_excel = 'local da sua planilha excel'
df = pd.read_excel(planilha_excel)

#anexo se tiver
def obter_anexos(caminhos_arquivos):
    anexos = []
    for caminho_arquivo in caminhos_arquivos:
        if os.path.exists(caminho_arquivo):  # Verificar se o arquivo existe antes de abri-lo
            with open(caminho_arquivo, 'rb') as arquivo:
                nome_arquivo = os.path.basename(caminho_arquivo)
                anexos.append((nome_arquivo, arquivo.read()))
        else:
            print(f"O arquivo '{caminho_arquivo}' não foi encontrado. O email será enviado sem anexos.")
    return anexos

#Lista para armazenar destinatarios que receberam o email
destinatarios_enviados = []

#Configurar servidor 
smtp_server = 'smtp.email.com'
smtp_port = numerodaporta
smtp_email = 'seuemail@email.com'
smtp_password = 'sua senha'

# Conectar ao servidor SMTP
server = smtplib.SMTP_SSL(smtp_server, smtp_port)

#Login no SMTP
server.login(smtp_email, smtp_password)

#Iterar por cada linha do DataFrame e criar os emails com anexos, se houver
for indice, linha in df.iterrows():
    if not pd.isnull(linha['destinatario']) and pd.isnull(linha['enviado']):
        destinatario = linha ['destinatario']
        print(f"Preparando email para: {destinatario}")
        assunto = linha['assunto']
        mensagem = linha['mensagem']
        anexos = ['local dos arquivos com a extensão do tipo do arquivo']

        em = EmailMessage()
        em['From'] = smtp_email
        em['To'] = destinatario
        em['Subject'] = assunto
        em.set_content(mensagem)

#Adicionar anexos se disponiveis
        if anexos:
            for nome_arquivo, conteudo_arquivo in obter_anexos(anexos):
                em.add_attachment(conteudo_arquivo, maintype='application', subtype='octet-stream', filename=nome_arquivo)

        try:
# Enviar o email
            server.sendmail(smtp_email, destinatario, em.as_string())
            destinatarios_enviados.append(destinatario)  # Adiciona o destinatário à lista de enviados
# Atualizar o DataFrame para indicar que o email foi enviado
            df.at[indice, 'enviado'] = 'Sim' #Marcar os contatos com enviados no DataFrame
        except Exception as e:
            print(f"Erro ao enviar email para {destinatario}: {e}")

        time.sleep(1)

server.quit()

#Salvar o DatafFrame no Excel
df.to_excel(planilha_excel, index=False)


print("Destinatarios que receberam o email com sucesso:")
print(destinatarios_enviados)










    


