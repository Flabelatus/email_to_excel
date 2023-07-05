import time

import flet
from flet import FilledButton, TextField, Text, Row, Page, FilePicker
import imaplib
import email
import pandas as pd
import datetime
from flet_core import FilePickerResultEvent

IMAP_SERVER = 'imap.gmail.com'
USERNAME = 'email.example.com'
PASSWORD = 'password'
ENCODINGS = ["utf-8", "cp1252", "ascii", "latin-1", "iso-8859-1", "iso-8859-15", "utf-16", "utf-32"]
WIDTH = 300

SENDER = "sender@example.com"
SENDER_TEST = "sender@example.com"
MAIL_SERVER_NAME = "mail.zxcs.nl"


def email_to_excel(
        # mail_server_text_filed: TextField,
        username_text_filed: TextField,
        password_text_filed: TextField,
        sender_address_text_field: TextField,
        limit_text_filed: TextField = 10,

        file_path: str = None

):
    # Email credentials and server details
    # Connect to the IMAP server
    imap = imaplib.IMAP4_SSL(IMAP_SERVER)
    imap.login(username_text_filed.value, password_text_filed.value)
    imap.select('INBOX')

    today = datetime.datetime.today()
    since_date = today - datetime.timedelta(days=20)
    before_date = today

    since_date_str = since_date.strftime('%d-%b-%Y')
    before_date_str = before_date.strftime('%d-%b-%Y')

    # Search for all emails
    print(sender_address_text_field.value)
    status, response = imap.search(None, 'FROM', sender_address_text_field.value, "SINCE", since_date_str, "BEFORE",
                                   before_date_str)
    email_ids = response[0].split()

    # Reverse the email IDs list
    email_ids = email_ids[::-1]
    print(email_ids)
    email_ids = email_ids[:int(limit_text_filed.value)]
    print(email_ids)

    # Initialize empty lists to store email data
    sender_list = []
    subject_list = []
    date_list = []
    body_list = []

    body = None

    # Iterate over each email and extract relevant information
    count = 0  # Counter for limiting the number of emails
    for email_id in email_ids:
        if count == int(limit_text_filed.value):
            break

        status, response = imap.fetch(email_id, '(RFC822)')
        raw_email = response[0][1]
        email_message = email.message_from_bytes(raw_email)

        sender = email_message['From']
        subject = email_message['Subject']
        date = email_message['Date']

        # if sender_email is not None and sender != sender_email:
        #     continue

        if email_message.is_multipart():
            for part in email_message.walk():
                content_type = part.get_content_type()
                if content_type == 'html':
                    continue
                if content_type == 'text/plain':
                    content = part.get_payload(decode=True)
                    for encoding in ENCODINGS:
                        try:
                            body = content.decode(encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    break
        else:
            content = email_message.get_payload(decode=True)
            for encoding in ENCODINGS:
                try:
                    body = content.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue

        sender_list.append(sender)
        subject_list.append(subject)
        date_list.append(date)
        body_list.append(body)

        count += 1

    # Create a pandas DataFrame with the email data
    email_data = pd.DataFrame({
        'Sender': sender_list,
        'Subject': subject_list,
        'Date': date_list,
        'Body': body_list
    })

    # Save the DataFrame to an Excel file
    email_data.to_excel(f'{file_path}received_emails.xlsx', index=False)
    # Close the IMAP connection
    imap.logout()

    print(f"Successfully saved {count} emails to 'received_emails.xlsx'.")
    return f"Successfully saved {count} emails to 'received_emails.xlsx'."


def main(page: Page):
    page.theme_mode = "dark"

    saving_destination = []

    mail_server = ""
    username = ""
    password = ""
    limit = 0
    sender_address = ""

    page.title = "Email to Excel"
    page.description = "A simple tool to convert emails to Excel files."
    page.window_width = 400
    page.window_height = 500

    # mail_server_text_filed = TextField(label="Enter your mail server", width=WIDTH,
    #                                    on_submit=lambda e: get_mail_server(e))

    username_text_filed = TextField(label="Enter your email address", width=WIDTH, on_submit=lambda e: get_username(e))
    password_text_filed = TextField(label="Enter your password", width=WIDTH, on_submit=lambda e: get_password(e),
                                    password=True)
    limit_text_filed = TextField(label="Enter the limit", width=WIDTH, on_submit=lambda e: get_limit(e))
    sender_address_text_field = TextField(label="Enter sender's address", width=WIDTH,
                                          on_submit=lambda e: get_sender(e))

    final_msg = Text("emails successfully saved to 'received_emails.xlsx'", visible=False)
    starting_msg = Text("Exporting...", visible=False)

    #
    # def get_mail_server(event: TextField):
    #     nonlocal mail_server
    #     mail_server += event.value
    #     return event.value

    def get_username(event: TextField):
        nonlocal username
        username += event.value
        return event.value

    def get_password(event: TextField):
        nonlocal password
        password += event.value
        return event.value

    def get_limit(event: TextField):
        nonlocal limit
        limit += int(event.value)
        return int(event.value)

    def get_sender(event: TextField):
        nonlocal sender_address
        sender_address += event.value
        return event.value

    def execute(event: FilledButton):
        starting_msg.visible = True
        page.update()

        email_to_excel(
            # mail_server_text_filed,
            username_text_filed,
            password_text_filed,
            sender_address_text_field,
            limit_text_filed,
            file_path=saving_destination[0]
        )
        final_msg.visible = True
        starting_msg.visible = False
        page.update()
        time.sleep(2)

    def get_path(event: FilePickerResultEvent):
        try:
            print(event.path)
            saving_destination.append(event.path + '/')
            return event.path
        except TypeError as e:
            print(e)
            return None

    path = FilePicker(on_result=get_path)
    page.overlay.append(path)
    page.update()

    def save_path(event: FilledButton):
        path.get_directory_path()
        page.update()

    page.views.append(
        Row(
            controls=[
                Text("Email to Excel", size=30),
                # mail_server_text_filed,
                username_text_filed,
                password_text_filed,
                limit_text_filed,
                sender_address_text_field,
                FilledButton("Choose file destination",
                             on_click=save_path),
                FilledButton("Submit", on_click=execute),
                starting_msg,
                final_msg,
            ]
        )
    )

    page.update()


flet.app(target=main)
