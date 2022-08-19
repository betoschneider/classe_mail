# -*- coding: utf-8 -*-
"""
Created on Fri Aug 19 16:33:52 2022

@author: Roberto Schneider
"""


#importando a classe
from classes.mail import Mail

remetente = 'meu_email_outlook@servidor.com.br'
destinatario = 'destinatario@servidor.com.br'
cc = 'cc@servidor.com.br'
cco = 'cco@servidor.com.br'
assunto = 'Assunto'
mensagem = '<p>Corpo do e-mail</p>' #para inserir tags html no corpo do e-mail, html = True
anexo = 'caminho_completo_do_arquivo'
mail = Mail()

resultado_envio = mail.envia(remetente, destinatario, assunto, mensagem, anexo, cc, cco, html = True)

print(resultado_envio)

#será criada uma exceção ao tentar exibir o valor do atributo __senha
try:
    print(mail.__senha)
    print(mail.usuario)
except:
    print(mail.usuario)