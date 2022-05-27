#!/usr/bin/env python
# coding: utf-8

# In[74]:


import win32com.client as win32
# criar a integração com o outlook
outlook = win32.Dispatch('Outlook.Application')


# In[75]:


# criar um email
email = outlook.CreateItem(0)


# In[76]:


email.Display()


# In[77]:


# configurar as informações do seu e-mail
email.To = "jveigabsb@gmail.com"
email.Subject = "Relatório Financeiro"
email.HTMLBody = "Bom dia"


# In[73]:


anexo = 'http://localhost:8888/edit/arquivo.xls


# In[68]:


email.Send()
print("Email Enviado")


# In[ ]:




