from time import sleep
import subprocess
import imaplib
import csv
import email
import json
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
import os
import re


class Autoprinter:
    def __init__(self, config="config.json"):
        with open(config) as json_file:
            data = json.load(json_file)

        self.username = data["username"]
        self.password = data["password"]
        imap_server = data["imap_server"]
        self.inbox_folder = data["inbox_folder"]
        self.printer_name = data["printer_name"]  # find the name with lpoptions
        smtp_server = data["smtp_server"]
        self.imap = imaplib.IMAP4_SSL(imap_server)

    def login(self):
        self.imap.login(self.username, self.password)  # TODO: CHECK

    def print_pdf(self, filename, mail_id):
        try:
            subprocess.run(["lp", "-d", self.printer_name, filename], check=True)
            print(f"Printing {filename} completed successfully.")
            status1, response1 = self.imap.uid("copy", str(mail_id), self.inbox_folder)
            status2, response2 = self.imap.uid(
                "store", str(mail_id), "+FLAGS", "(\Deleted)"
            )
            status3, response3 = self.imap.expunge()
            print(status1, status2, status3)
            if os.path.exists(filename):
                os.remove(filename)
                print("File removed successfully.")
            else:
                print("File does not exist.")
        except subprocess.CalledProcessError as e:
            sn
            print(f"Printing {filename} failed. Error: {e}")  # TODO EMAIL ALERT
        print(filename, mail_id)

    def save_pdf(self, msg, msg_id):
        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = str(part.get("Content-Disposition"))
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename.endswith(".pdf"):
                        tmp_folder_name = "invoices"
                        if not os.path.isdir(tmp_folder_name):
                            os.mkdir(tmp_folder_name)
                        filepath = os.path.join(tmp_folder_name, filename)
                        open(filepath, "wb").write(part.get_payload(decode=True))
                        self.print_pdf(filepath, msg_id)

    def select_folder(self):
        status, messages = self.imap.select("AUTOMATIC")
        if status != "OK":
            raise Exception(status)
        return messages

    def print_all(self):
        messages = self.select_folder()
        statuts, raw_msg_ids = self.imap.uid("SEARCH", None, "ALL")
        msg_ids = [int(x) for x in raw_msg_ids[0].split()]
        for uid in msg_ids:
            res, msg = self.imap.uid("fetch", str(uid), "(RFC822)")
            print(res, msg)
            for response in msg:
                if isinstance(response, tuple):
                    msg = email.message_from_bytes(response[1])
                    self.save_pdf(msg, uid)

    def close_session(self):
        self.imap.close()
        self.imap.logout()

    def run(self):
        self.login()
        while True:
            self.print_all()
            self.imap.noop()
            sleep(60)
        self.close_session()
