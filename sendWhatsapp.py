from twilio.rest import Client
from flask import Flask

def app():
    app = Flask(__name__)

    @app.route('/')
    def send_whatsapp():
        to = 'RECIPIENT_PHONE_NUMBER'  # Replace with recipient's phone number in WhatsApp format
        message = 'YOUR_MESSAGE'
        
        account_sid = 'YOUR_ACCOUNT_SID'
        auth_token = 'YOUR_AUTH_TOKEN'
        client = Client(account_sid, auth_token)
        
        message = client.messages.create(
            body=message,
            from_='whatsapp:+14155238886',  # Replace with your Twilio WhatsApp number
            to=f'whatsapp:{to}'
        )
        
        return 'Message sent!'
    
    if __name__ == '__main__':
        app.run(debug=True)

app()