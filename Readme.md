# Envio de E-mail com Python: 


Este projeto é para aprendizagem do envio de e-mails usando Python. Logo ele utiliza a seguinte biblioteca que possibilita o funcionamento da aplicação:

* **pywin32**

*Esta biblioteca realiza a conexão com o e-mail e a elaboração da mensagem que será enviada.*

# Códigos Usados:

* **win32.Dispatch**
  * Cria a integração do Python com o e-mail podendo ser gmail, hotmail, outlook, etc
  

* **outlook.CreateItem**
  * Configura o email que será usado. No nosso será o outlook
  

* **email.To**
  * Configura o destinatário da mensagem  
  

* **email.Subject**
  * Configura o assunto da mensagem 


* **email.HTMLbody**
  * Configura a mensagem usando HTML 
 
 
* **email.Send()**
  * Envia o email 
