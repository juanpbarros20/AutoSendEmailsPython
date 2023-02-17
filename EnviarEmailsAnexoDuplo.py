import win32com.client as client
import pandas as pd
#Fazendo conexão com o Outlook
outlook = client.Dispatch('Outlook.Application')

#Inserindo o E-mail remetente
emissor = outlook.session.Accounts['seuemail@outlook.com']

#Lendo o arquivo excel que será atribuido a coluna como dados.
contatos = pd.read_excel('C:Aqui vai o caminho da pasta onde está o arquivo Excel.xlsx')
contatos.info()

#Criando tabela de envio para os contatos (aqui deve consta todas as colunas que estão na planilha)
dados = contatos[['NOME','DESTINATARIO','COLUNA2', 'COLUNA3', 'COLUNA4', 'ANEXO','ANEXO2','CONTADOR']].values.tolist()

for dado in dados:
    NOME = dado[0]
    DESTINATARIO = dado[1]
    COLUNA2 = dado[2]
    COLUNA3 = dado[3]
    COLUNA4 = dado[4]
    ANEXO = dado[5]
    ANEXO2 = dado[6]
    #Contador inserido apenas para identificar o endereçado, assim saber se foi enviado ou não.
    CONTADOR = dado [7]

    mensagem = outlook.CreateItem(0)
#Se for uma quantidade grande de emails é bom retirar o "mensagem.Display()"
    #mensagem.Display()
    mensagem.To = DESTINATARIO
#Digitar o Assunto aqui    
    mensagem.Subject = 'Aqui vai o Assunto do E-mail'
#Digitar o Texto aqui.
    mensagem.HTMLBody = f"""

    <p>Prezado(a) {NOME}, boa tarde! </p>
    <p> </p>
    <p> </p>
    <p> Corpo do Email paragrafo 1. </p>
    <p> Corpo do Email paragrafo 2</a> </p>
    <p> Corpo do Email paragrafo 3 </p>
    <p> </p>
    <p> </p>
    <p>Atenciosamente, </p>
    <p> </p>
    Seu nome <br />
    Seu cargo <br />
    Seus contatos <br />

    """
#Caso tenha anexo informar o caminho (Senão tiver anexo, COMENTAR)    
    ANEXO = "C:/ Aqui vai o caminho para encontrar o arquivo no seu PC, colocar extensão. Ex.: .pdf"
    ANEXO2 = "C:/ Aqui vai o caminho para encontrar o segundo arquivo no seu PC, colocar extensão. Ex.: .xlms"
    mensagem.Attachments.Add(ANEXO)
    mensagem.Attachments.Add(ANEXO2)

#Argunmento para selecionar o emissor
    mensagem._oleobj_.Invoke(*(64209,0,8,0,emissor))
# Salvar email
    mensagem.Save()
#enviar email
    mensagem.Send()
    print("E-mail ",CONTADOR," enviado!")