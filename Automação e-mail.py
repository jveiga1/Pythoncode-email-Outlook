import win32com.client as win32
# criar a integração com o outlook
outlook = win32.Dispatch('Outlook.Application')

# criar um email
email = outlook.CreateItem(0)
email.Display()

# configurar as informações do seu e-mail
email.To = "teste@teste.com"
email.Subject = "Relatório Financeiro"
email.HTMLBody = "Bom dia"

anexo = 'http://localhost:8888/edit/arquivo.xls

email.Send()
print("Email Enviado")




