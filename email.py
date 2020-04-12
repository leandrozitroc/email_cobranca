import openpyxl
import smtplib
import sys

planilha = openpyxl.load_workbook("duesRecords.xlsx")
plan1 = planilha['Sheet1']
coluna = list(plan1.values)

devedores = {}
for _ in coluna:

    for n in _:
        if n == None:
            name = _[0]
            email = _[1]
            devedores[name] = email
print(devedores)

smtp1 = smtplib.SMTP('smtp.gmail.com', 587)
smtp1.ehlo()
smtp1.starttls()
smtp1.login('my_email_address@gmail.com', sys.argv[1])

for name , email in devedores.items():
    body = f"Assunto: Parcelas em Aberto.n\ Ola {name}, encontramos parcelas do seu contrato em aberto\n" \
           f"Por favor entre em contato:"
    print(f'Enviando email para {email}')
    status_envio = smtp1.sendmail('myemail@gamil.com', email , body)
    if status_envio != {}:
        print(f"Problema no envio do email {email} , {status_envio}")
    smtp1.quit()
