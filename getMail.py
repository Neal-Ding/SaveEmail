#! /usr/bin/env python3

import os
import json
import imaplib
import pyzmail
import shutil
from datetime import datetime
import time

config = {}


def get_mails(M, UIDs):
    for UID in UIDs:
        tye, raw_message_header = M.fetch(UID, 'BODY[HEADER]')
        message_header = pyzmail.PyzMessage.factory(raw_message_header[0][1])
        mail_time = message_header.get_decoded_header('date')[0:31].strip()
        mail_time = datetime.strptime(mail_time, '%a, %d %b %Y %H:%M:%S %z')
        sender = message_header.get_addresses('from')[0][1]
        subject = message_header.get_subject()

        if sender in config['sender_list'] and (config['last_timestamp'] == 0 or mail_time.timestamp() > config['last_timestamp']):
            print("Receive a New Email from {sender}".format(sender=sender))
            typ, raw_message = M.fetch(UID, 'BODY[]')
            message = pyzmail.PyzMessage.factory(raw_message[0][1])

            folder_prefix = datetime.strftime(mail_time, '%Y%m%d')
            config['last_date'] = datetime.strftime(mail_time, '%d-%b-%Y')
            config['last_timestamp'] = mail_time.timestamp()

            folder_name = os.path.join(config['mail_folder'], folder_prefix + subject)

            if os.path.exists(folder_name):
                shutil.rmtree(folder_name)
            os.makedirs(folder_name)

            print(folder_prefix + subject)

            for mailpart in message.mailparts:
                filePath = os.path.join(os.path.dirname(os.path.abspath(__file__)), folder_name, mailpart.sanitized_filename)
                open(filePath, 'wb').write(mailpart.get_payload())

            with open('config.json', 'w') as f:
                json.dump(config, f)


def login_imap(server, username, password, folder):
    M = imaplib.IMAP4_SSL(host=server)
    typ, data = M.login(username, password)
    if typ == 'OK':
        print(data[0].decode())

    typ, data = M.select(mailbox=folder)
    if typ == 'OK':
        print('In {folder} folder, you have {mail_sum} emails'.format(folder=folder, mail_sum=data[0].decode()))
    return M


def search_mail(M):
    print("Search for new email...")
    if config['last_timestamp'] == 0:
        typ, data = M.search(None, 'ALL')
        print('Receive All Mail')
    else:
        typ, data = M.search(None, 'SINCE', config['last_date'])
        print('Receive Mail from {date}'.format(date=config['last_date']))

    if typ == 'OK':
        return M, data[0].split()
    else:
        return M, []


def logout_imap(M):
    M.close()
    M.logout()
    print("LOGOUT completed!")


def read_config():
    global config

    print("Update Config Info")
    with open('config.json', encoding='utf-8') as json_data:
        config = json.load(json_data)


def main():
    global config

    read_config()
    M = login_imap(**config['user_info'])
    try:
        while True:
            time.sleep(config['interval_time'])
            read_config()
            M, UIDs = search_mail(M)
            get_mails(M, UIDs)

    finally:
        logout_imap(M)

if __name__ == "__main__":
    main()
