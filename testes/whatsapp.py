import os
from twilio.rest import Client

ACCOUNT_SID = os.environ["TWILIO_ACCOUNT_SID"]
AUTH_TOKEN = os.environ["TWILIO_AUTH_TOKEN"]

client = Client(ACCOUNT_SID, AUTH_TOKEN)

message = client.messages.create(
    from_="whatsapp:+14155238886",
    to="whatsapp:+55SEUNUMERO",
    body="Olá! 🚀 Esta mensagem foi enviada pelo Python."
)

print("Mensagem enviada!")
print("SID:", message.sid)