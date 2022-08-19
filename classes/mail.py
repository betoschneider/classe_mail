# -*- coding: utf-8 -*-
"""
@author: Roberto Schneider
"""


class Mail:
    def __init__(self):
        self.usuario = 'meu_email_outlook@servidor.com.br'
        self.__senha = 'minha_senha' #__ antes do atributo impede a visualização desse atributo ao instanciar um objeto
        self.servidor_smtp = 'SMTP.office365.com: 587'
        self.servidor_imap = 'outlook.office365.com'
        
        
    def envia(self, remetente, para, assunto, mensagem, anexo = [],
                 cc = '', cco = '', html = False):
                
        #importando bibliotecas necessárias
        import smtplib
        import email # Modulos para manipulação de email
        import email.mime.application # Modulos para manipulação de email
        import email.mime.text
        import email.mime.multipart #import MIMEMultipart
        
        #Dados de Login e envio
        usuario = self.usuario
        senha = self.__senha
        servidor = self.servidor_smtp
        
        # Cabeçalho da mensagem do email
        msg = email.mime.multipart.MIMEMultipart()
        msg['Subject'] = assunto
        msg['From'] = remetente
        msg['To'] = para
        msg['Cc'] = cc
        msg['Bcc'] = cco
            
        # Corpo principal do email (tb um anexo)
        body = email.mime.text.MIMEText(mensagem)
        
        # Para inserir tags html no corpo do e-mail
        if html:
            body = email.mime.text.MIMEText(mensagem, 'html')
        msg.attach(body)
        
        
        if str(type(anexo)) != "<class 'list'>":
            anexo = [anexo]
            
        try:
            # Anexando o arquivo
            # Percorre a lista de caminhos adicionando o conteúdo a msg
            for i in anexo:
                nome_arquivo_anexo = i.split('/')[-1]
                fp=open(i,'rb')
                anexos = email.mime.application.MIMEApplication(fp.read(),_subtype = i.split('.')[-1])
                fp.close()
                anexos.add_header('Content-Disposition','attachment',filename=nome_arquivo_anexo)
                msg.attach(anexos)
        except:
            anexo = []
        
        try:
            # Enviando via Outlook server 
            s = smtplib.SMTP(servidor)
            s.starttls()
            s.login(usuario, senha)
            s.sendmail(msg['From'],msg['To'].split(',') + msg['Cc'].split(',') + msg['Bcc'].split(','), msg.as_string())
        except:
            s.quit()
            return False
        else:
            s.quit()
            return True


if __name__ == '__main__':
    
    remetente = 'meu_email_outlook@servidor.com.br'
    destinatario = 'destinatario@servidor.com.br'
    cc = 'cc@servidor.com.br'
    cco = 'cco@servidor.com.br'
    assunto = 'Assunto'
    mensagem = '<p>Corpo do e-mail</p>'
    anexo = 'caminho_completo_do_arquivo'
    mail = Mail()
    
    resultado_envio = mail.envia(remetente, destinatario, assunto, mensagem, anexo, cc, cco, html = True)
    
    print(resultado_envio)
    

