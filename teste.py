import requests

url = "https://graph.facebook.com/v19.0/YOUR_PHONE_NUMBER_ID/messages"

headers = {
    "Authorization": "Bearer YOUR_TEMP_TOKEN",
    "Content-Type": "application/json"
}

payload = {
    "messaging_product": "whatsapp",
    "to": "SEU_NUMERO_VERIFICADO",  # formato internacional, ex: 5511999999999
    "type": "text",
    "text": {"body": "Olá! Isso é um teste da API do WhatsApp da Meta com Python!"}
}

res = requests.post(url, headers=headers, json=payload)

print(res.status_code)
print(res.json())
