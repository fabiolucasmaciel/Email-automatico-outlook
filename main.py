import win32com.client as win32

# integração com o outlook
outlook = win32.Dispatch('outlook.application')

# cria um email
email = outlook.CreateItem(0)

# configurar as informações do seu email
email.To = "fabio.lucas1403@gmail.com"
email.Subject = "Teste"
email.HTMLBody = """
<h1>Teste</h1>
<hr>
<p>Boa noite</p>
"""

email.Send()
print("Email Enviado")
