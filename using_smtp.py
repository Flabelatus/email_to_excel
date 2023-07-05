import datetime
import imaplib
import email
import pandas as pd

# IMAP server settings
IMAP_SERVER = 'mail.zxcs.nl'
IMAP_PORT = 993
USERNAME = 'fractie@pvdarotterdam.nl'
PASSWORD = ''

# Sender's email address
SENDER_EMAIL = 'rotterdam@pvda.nl'

# Excel file path to save the emails
EXCEL_FILE_PATH = 'emails.xlsx'


def save_emails_to_excel():
    # Connect to the IMAP server
    server = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    server.login(USERNAME, PASSWORD)
    server.select("inbox")

    # Calculate the date range for the last 30 days
    end_date = datetime.datetime.now()
    start_date = end_date - datetime.timedelta(days=60)
    start_date_str = start_date.strftime("%d-%b-%Y")
    end_date_str = end_date.strftime("%d-%b-%Y")

    # Search for emails from the specific sender within the date range
    search_criteria = f'(FROM "{SENDER_EMAIL}") SINCE "{start_date_str}" BEFORE "{end_date_str}"'
    status, email_ids = server.search(None, search_criteria)
    email_ids = email_ids[0].split()

    # Initialize a list to store email data
    emails = []

    for email_id in email_ids:

        _, data = server.fetch(email_id, '(RFC822)')
        raw_email = data[0][1]
        parsed_email = email.message_from_bytes(raw_email)

        # Extract relevant information from the email
        email_subject = parsed_email['Subject']
        email_date = parsed_email['Date']
        email_body = ''

        if parsed_email.is_multipart():
            for part in parsed_email.get_payload():
                if part.get_content_type() == 'text/plain':
                    email_body = part.get_payload(decode=True).decode('utf-8')
        else:
            email_body = parsed_email.get_payload(decode=True).decode('utf-8')

        # Append email data to the list
        emails.append({'Subject': email_subject, 'Date': email_date, 'Body': email_body})

    # Close the IMAP server connection
    server.close()
    server.logout()

    # Save the emails to an Excel file
    df = pd.DataFrame(emails)
    df.to_excel(EXCEL_FILE_PATH, index=False)

    print(f"Emails from {SENDER_EMAIL} within the last 30 days saved to {EXCEL_FILE_PATH}.")


if __name__ == '__main__':
    # Call the function to save emails to Excel
    save_emails_to_excel()
