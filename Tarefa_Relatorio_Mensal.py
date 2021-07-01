#Nós tinhamos um problema em relação um relatório que mensalmente deveria ser retirado do banco de dados atráves de uma
#query e com esse data frame montava um arquivo excel.
#Para automatizar essa tarefa foi desenvolvido o código abaixo e com a ajuda do gerenciador de tarefas rodamos um bat
#que executa esse procedimento em cada começo de mês, esse script contém configuração do envio de email para

#Importação de Biblioteca
import pyodbc
import pandas as pd
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Função de conexão
def retornar_conexao_sql():
    server = "localhost" #Aqui você indica o servidor que contém o banco de dados
    database = "DBTeste" #Aqui você indica o nome do banco de dados
    string_conexao = 'Driver={SQL Server Native Client 11.0};Server='+server+';Database='+database+';Trusted_Connection=yes;'
    #Nesse tipo de string de conexão é essencial que o usuário que estará executando o arquivo Python
    conexao = pyodbc.connect(string_conexao)
    return conexao

#Variável do cursor
conn = retornar_conexao_sql()

#Consultando com função
sql = 'SELECT * FROM TABELATESTE'
#Utilize a query que deve retornar o data set (recomendo utilizar uma procedure)

#Salva data set na variável
dados = pd.read_sql(sql, conn)

# Exporta data set para um arquivo excel na pasta do projeto
dados.to_excel('Planilha.xlsx', sheet_name='planilha', na_rep='#N/A', header=True, index=False)

#Monta e envia e-mail
mail_from = 'recebe@gmail.com'
mail_to = 'envia@gmail.com'

msg = MIMEMultipart()

msg['From'] = mail_from
msg['To'] = mail_to
msg['Subject'] = 'Relatório'

#Abaixo fica a variável que armazena o corpo do e-mail
body = '''
Esse é um e-mail automático.

Segue em anexo das planilha contendo as informações.

'''
msg.attach(MIMEText(body))

#Transforma anexo em base 64 pra inserir na variável
filemane = 'planilha.xlsx'
attachment = open(filemane, 'rb')

mimetype_anexo = mimetypes.guess_type(filemane)[0].split('/')
part = MIMEBase(mimetype_anexo[0], mimetype_anexo[1])

part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename = %s" % filemane)

msg.attach(part)

#Estrutura para encaminhar o e-mail
server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(mail_from, open('senha.txt').read().strip()) #Existe um arquivo chamado senha.txt onde consta a senha do email
                                                          #Esse arquivo deve esta em conjunto ao arquivo python.
text = msg.as_string()
server.sendmail(mail_from, mail_to, text)
server.quit()

