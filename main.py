import imaplib
import email
import os
from datetime import datetime

def connect_to_email_server(imap_server, email, pwd):
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(email, pwd)
    return imap


def select_inbox(imap):
    imap.select("Inbox")

def get_current_year():
    full_yr = datetime.today().year
    short_yr = str(full_yr)[-2:]

    return full_yr, short_yr

def get_user_date():
    usr_date = input("Select the date you want to print out in this format MM/DD: ")
    trans_date = datetime.strptime(usr_date, "%m/%d")
    
    return trans_date

def format_dates(date, short_yr, full_yr):
    download_folder = datetime.strftime(date, f"%m-%d-{short_yr}")
    imap_date = datetime.strftime(date, f"%d-%b-{full_yr}")

    return download_folder, imap_date

def search_emails(imap, imap_date):
    return imap.search(None, f'SUBJECT "Invoice" ON {imap_date} FROM "customerservice@highcountrylumber.com"')

def save_attachements(msgnums, imap, download_folder):
    saved_files = 0
    for msgnum in msgnums[0].split():
        _, data = imap.fetch(msgnum, "RFC822")

        message = email.message_from_bytes(data[0][1])

        filename = message.get('Subject').split()[1] + ".pdf"
        #print(filename)

        att_path = "No attachment found."
        for part in message.walk():

            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue
            
            os.makedirs(download_folder, exist_ok=True)
            att_path = os.path.join(download_folder, filename)
            
            #print(att_path)
            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        saved_files += 1

    return saved_files

def main():
    imap_server = "imap.gmail.com"
    email = input("Please enter the email: ")
    pwd = input("Please enter the password: ") 

    imap = connect_to_email_server(imap_server, email, pwd)
    select_inbox(imap)

    full_yr, short_yr = get_current_year()
    date = get_user_date()
    download_folder, imap_date = format_dates(date, short_yr, full_yr)

    _, msgnums = search_emails(imap, imap_date)

    saved_files_count = save_attachements(msgnums, imap, download_folder)
    print(f"{saved_files_count} files have been saved.")

    imap.close()


if __name__ == "__main__":
    main()