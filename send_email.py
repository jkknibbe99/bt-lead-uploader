import smtplib, ssl
from config import DataCategories, get_config_data

# Send an email
def sendEmail(subject: str, message: str, receiver_emails: list):
    # Create a secure SSL context
    port = 465  # For SSL
    context = ssl.create_default_context()

    # Create message string
    message_str = 'Subject: {}\n\n{}'.format(subject, message)

    # Send email
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        sender_email = get_config_data(DataCategories.EMAIL_DATA, 'status_sender_email')
        password = get_config_data(DataCategories.EMAIL_DATA, 'status_email_password')
        try:
            server.login(sender_email, password)
        except smtplib.SMTPAuthenticationError as e:
            print(e)
            print(sender_email, password)
        for receiver_email in receiver_emails:
            server.sendmail(sender_email, receiver_email, message_str)
