import imaplib

EMAIL = "ti.ahcmaq@gmail.com"
SENHA_APP = "fsno grka ifsq jmzm"

imap = imaplib.IMAP4_SSL("imap.gmail.com")
imap.login(EMAIL, SENHA_APP)

# pasta de enviados (PT-BR)
status, _ = imap.select('"[Gmail]/E-mails enviados"')
print("SELECT STATUS:", status)

# busca TODOS os emails
status, data = imap.search(None, 'ALL')

ids = data[0].split()
print(f"Encontrados {len(ids)} emails para apagar")

for num in ids:
    imap.store(num, '+FLAGS', '\\Deleted')
    print(num)

imap.expunge()
imap.logout()

print("✅ Todos os e-mails enviados foram apagados")