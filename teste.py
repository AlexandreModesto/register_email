import pandas as pd
import PySimpleGUI as sg
import imaplib,email,os
from email.header import decode_header
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# load_dotenv()
# IMAP_SERVER = os.getenv("IMAP_SERVER")
# EMAIL_ACCOUNT = os.getenv("IMAP_EMAIL")
# PASSWORD = os.getenv("IMAP_APP_PASSWORD")
sg.theme("Dark Purple 4")
def import_sheet():
    sheet=sg.popup_get_file("Importe o relatório")
    global df
    df=pd.DataFrame(sheet)


def clean_text(text):
    if isinstance(text,bytes):
        text= text.decode("utf-8",errors="replace")
    return text

def extract_text_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    return soup.get_text()

def fetch_emails():
    mail = imaplib.IMAP4_SSL("IMAP_SERVER")
    mail.login("EMAIL_ACCOUNT", "PASSWORD")
    mail.select()  # Escolhe a caixa de entrada

    status,messages = mail.search(None,'UNSEEN')
    email_ids=messages[0].split()

    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id,'(RFC822)')

        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg=email.message_from_bytes(response_part[1])
                subject,enconding=decode_header(msg['Subject'])[0]
                if isinstance(subject,bytes):
                    subject=subject.decode(enconding if enconding else 'utf-8')
                if subject == "Assunto":
                    # Verifica se o e-mail tem partes
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get('Content-Disposition'))

                            if content_type == 'text/plain' and 'attachment' not in content_disposition:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    body = payload.decode()
                                    register_email(clean_text(body))
                            # elif content_type == 'text/html' and 'attachment' not in content_disposition:
                            #     payload = part.get_payload(decode=True)
                            #     if payload:
                            #         html_content = payload.decode()
                            #         body = extract_text_from_html(html_content)
                            #         print(f"Body (HTML): {clean_text(body)}")
                    else:
                            # Caso o e-mail não seja multipart
                            content_type = msg.get_content_type()
                            payload = msg.get_payload(decode=True)
                            if payload:
                                if content_type == 'text/plain':
                                    body = payload.decode()
                                    register_email(clean_text(body))
                                # elif content_type == 'text/html':
                                #     html_content = payload.decode()
                                #     body = extract_text_from_html(html_content)
                                #     print(f"Body (HTML): {clean_text(body)}")
        break
    mail.logout()

def register_email(cleaned_body):
    splited_text=cleaned_body.split(":")
    for t in splited_text:
        if t.lower() == "nome":
            nome=splited_text[splited_text.index(t)+1]

    df['Nome Cliente']=df['Nome Cliente']+nome

    df.to_excel("realtório.xlsx",index=False)

if __name__ == '__main__':
    import_sheet()
    fetch_emails()