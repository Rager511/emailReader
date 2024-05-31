import imapclient
import email
from email.header import decode_header
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

def parse_email_body(body):
    parsed_data = {
        "Nom complet": "",
        "Courte description du projet": "",
        "Numéro de téléphone": "",
        "Adresse courriel": "",
        "Ville": "",
        "Code postal": "",
        "Choix de mon École": "",
        "Meilleurs moments pour joindre": ""
    }

    lines = body.split('\n')

    current_key = None
    for line in lines:
        line = line.strip()
        if line.startswith("Nom complet :"):
            current_key = "Nom complet"
            parsed_data[current_key] = line.split(":", 1)[1].strip()
        elif line.startswith("Courte description du projet :"):
            current_key = "Courte description du projet"
            parsed_data[current_key] = line.split(":", 1)[1].strip()
        elif line.startswith("Numéro de téléphone :"):
            current_key = "Numéro de téléphone"
            parsed_data[current_key] = line.split(":", 1)[1].strip()
        elif line.startswith("Adresse courriel :"):
            current_key = "Adresse courriel"
            parsed_data[current_key] = line.split(":", 1)[1].strip()
        elif line.startswith("Ville :"):
            current_key = "Ville"
            parsed_data[current_key] = line.split(":", 1)[1].strip()
        elif line.startswith("Code postal :"):
            current_key = "Code postal"
            parsed_data[current_key] = line.split(":", 1)[1].strip()
        elif line.startswith("Choix de mon École :") and "Quels sont les meilleurs moments pour vous joindre par téléphone ? :" in line:
            parts = line.split("Quels sont les meilleurs moments pour vous joindre par téléphone ? :")
            parsed_data["Choix de mon École"] = parts[0].split(":", 1)[1].strip()
            parsed_data["Meilleurs moments pour joindre"] = parts[1].strip() if len(parts) > 1 else ""
        elif line.startswith("Choix de mon École :"):
            current_key = "Choix de mon École"
            parsed_data[current_key] = line.split(":", 1)[1].strip()
        elif line.startswith("Quels sont les meilleurs moments pour vous joindre par téléphone ? :"):
            current_key = "Meilleurs moments pour joindre"
            parsed_data[current_key] = line.split(":", 1)[1].strip()
        elif current_key:
            if current_key == "Meilleurs moments pour joindre" and (
                "Merci de bien vouloir contacter" in line or "Notez que nous enverrons" in line or "Merci de votre collaboration" in line or "L'équipe Entrepreneuriat Québec" in line):
                current_key = None
            elif current_key:
                parsed_data[current_key] += f" {line.strip()}"

    for key in parsed_data:
        parsed_data[key] = parsed_data[key].strip()

    return parsed_data

def fetch_emails(username, password, sender_email, since_date):
    import imaplib
    import email

    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(username, password)
    imap.select("inbox")

    result, data = imap.search(None, f'(FROM "{sender_email}" SINCE "{since_date}")')

    email_data = []

    if result != 'OK':
        print("Error searching for emails")
        return email_data

    email_ids = data[0].split()
    for email_id in email_ids:
        result, message_data = imap.fetch(email_id, "(RFC822)")

        if result != 'OK':
            print(f"Error fetching email ID {email_id}")
            continue

        raw_email = message_data[0][1]
        msg = email.message_from_bytes(raw_email)

        email_subject = msg["Subject"]
        email_from = msg["From"]
        email_date = msg["Date"]

        body_content = ""

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    try:
                        body_content += part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except:
                        pass
        else:
            try:
                body_content = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
            except:
                pass

        parsed_body = parse_email_body(body_content)
        parsed_body.update({
            "Subject": email_subject,
            "From": email_from,
            "Date": email_date
        })

        print(parsed_body)
        email_data.append(parsed_body)

    imap.close()
    imap.logout()

    return email_data

def save_to_excel(email_data, excel_file):
    headers = [
        "Nom complet",
        "Courte description du projet",
        "Numéro de téléphone",
        "Adresse courriel",
        "Ville",
        "Code postal",
        "Choix de mon École",
        "Meilleurs moments pour joindre",
        "Subject",
        "From",
        "Date"
    ]

    df = pd.DataFrame(email_data, columns=headers)
    
    wb = Workbook()
    ws = wb.active
    
    ws.append(headers)
    
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
    
    try:
        wb.save(excel_file)
        print("Email data saved to", excel_file)
    except Exception as e:
        print("Error occurred while saving to Excel:", e)

def main():
    username = input("Enter your email username: ")
    password = input("Enter your email password: ")
    sender_email = input("Enter the sender email to filter: ")
    since_date = input("Enter the date since (e.g., 01-Jan-2022): ")
    excel_file = input("Enter the name of the Excel file to save data (e.g., emails.xlsx): ")

    emails = fetch_emails(username, password, sender_email, since_date)
    for email_data in emails:
        print(email_data)

    save_to_excel(emails, excel_file)

if __name__ == "__main__":
    main()
