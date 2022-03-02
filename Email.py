#importação da biblioteca para a conexão com e-mail
import win32com.client as win32

#Integração do python com e-mail
outlook = win32.Dispatch("outlook.application")

#Configurações do e-mail
email = outlook.CreateItem(0)

#Parâmetros de envio do e-mail
email.To = "vinicius2004gab@gmail.com; isa.vini@hotmail.com" #Destinatário
email.Subject = "Teste Com Python" #Assunto do e-mail

#Mensagem do e-mail
email.HTMLbody = """<p>Boa Noite</p><br><br>
<p>Testando um código python para enviar um e-mail</p><br><br>
<p>Atenciosamente</p>
<p>Seu Código Python</p>
"""

#Envio do e-mail
email.Send()
print("Email Enviado com Sucesso")
