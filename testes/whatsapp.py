import requests

TOKEN = "EAATz4ltDrooBSVmzxPqV6g2XNJUE02OzB9zmNfPGuO8Oe4vAtjSe0ybSFS7HotiKxxfTsWOhFgvzRCVdb6uX0XCoTWSZA1lGKhXjQvAr31xf7eXZCxFNnYFWUAuwegOfeuwnZCV5BNNiE6XSdvx3mOUPMiyxHZC4Qhw2nE9fnHb7qWZCFcd5hCZCuOVbAnweVrzlNIU0focq4BWzFEvTeCsBwUWP9nFoReTvn5tAF7XguqAGU8qvBKklQ2jrZCYhuDuZAyRmxdZCnZA553onq20m8q"
PHONE_NUMBER_ID = "1256537490880143"
NUMERO_DESTINO = "5551981158315"

url = f"https://graph.facebook.com/v26.0/{PHONE_NUMBER_ID}/messages"

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
}

dados = {
    "messaging_product": "whatsapp",
    "to": NUMERO_DESTINO,
    "type": "template",
    "template": {
        "name": "hello_world",
        "language": {
            "code": "en_US"
        }
    }
}

resposta = requests.post(url, headers=headers, json=dados)

print("Status:", resposta.status_code)
print("Resposta:", resposta.json())