import imaplib
import email
import smtplib
from email.mime.text import MIMEText
import pandas as pd

def enviar_email_resposta(Notas_HHIB, pedido_compra):
    # Configurações do servidor SMTP
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    sender_email = 'lucas.degrande@bmchyundai.com.br'
    password = 'wssbhscozclgrphe'
    
    # Mensagem de resposta
    mensagem_base = """Prezados HCE-BR,

Por favor, podem verificar e enviar outra NF complementar para complementar o pedido {pedido_compra}?

Atenciosamente"""

    # Conectando ao servidor SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, password)
    
    # Conectando ao servidor IMAP
    imap_server = 'imap.gmail.com'
    imap_port = 993
    imap_user = sender_email
    imap_password = password
    
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(imap_user, imap_password)
    mail.select('inbox')
    
    print("Número de notas para verificar:", len(Notas_HHIB))

    for numero_nota, pedido in zip(Notas_HHIB, pedido_compra):
        # Buscando e-mails na caixa de entrada para o número da nota fiscal
        result, data = mail.search(None, '(OR (SUBJECT "Envio da NFe - Emitente: HYUNDAI CONSTRUCTION EQUIPMENT BRASIL N.Doc: {0}") (CC "Envio da NFe - Emitente: HYUNDAI CONSTRUCTION EQUIPMENT BRASIL N.Doc: {0}"))'.format(numero_nota))
        
        # Verificar se a busca retornou resultados
        if data[0]:
            email_ids = data[0].split()
            print(f"E-mails encontrados para a nota fiscal {numero_nota}")

            for email_id in email_ids:
                result, email_data = mail.fetch(email_id, '(RFC822)')
                raw_email = email_data[0][1]
                msg = email.message_from_bytes(raw_email)
                remetente = msg['From']
                destinatarios_cc = msg.get('CC', '').split(', ')  # Obter destinatários em cópia (CC)
                destinatarios = [remetente] + destinatarios_cc  # Lista de destinatários (remetente + CC)
                                
            # Enviando resposta
            subject = f'Re: {msg["Subject"]}'
            mensagem = mensagem_base.format(pedido_compra=pedido)
            msg = MIMEText(mensagem)
            msg['Subject'] = subject
            msg['From'] = sender_email
            msg['To'] = remetente  # Apenas o remetente no campo "To"
            msg['Cc'] = ', '.join(destinatarios_cc)  # Adicionar destinatários em cópia ao cabeçalho do e-mail
            server.sendmail(sender_email, ', '.join(destinatarios), msg.as_string())  # Enviar e-mail para todos os destinatários
            print(f'Resposta enviada para: {destinatarios}')

        else:
            print(f"Nenhum e-mail encontrado para a nota fiscal {numero_nota}")

    # Finalizando conexões
    server.quit()
    mail.close()
    mail.logout()

# Carregando os números de nota fiscal da planilha
planilha = pd.read_excel('FaltaItemPedido.xlsx')
Notas_HHIB = planilha['Nota fiscal'].tolist()
pedido_compra = planilha['PEDIDO COMPRA'].tolist()

# Chamada da função com os números de notas fiscal da planilha
enviar_email_resposta(Notas_HHIB, pedido_compra)






