import imaplib
import time
import logging
import json
import sys
import os
from socket import gaierror

imaplib._MAXLINE = 10_000_000  # Higher limit for imaplib

def setup_logger(log_path):
    logging.basicConfig(
        filename=log_path,
        level=logging.DEBUG,
        format='%(asctime)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    return logging.getLogger()

def handle_exception():
    return logging.Formatter().formatException(sys.exc_info())

def ensure_directory_exists(directory):
    os.makedirs(directory, exist_ok=True)

def initialize_imap(imap_host, imap_port, username, password, mailbox, logger):
    try:
        logger.info('Attempting IMAP login')
        
        # mail = imaplib.IMAP4_SSL(imap_host, imap_port) # SSL
        #STARTTLS
        mail = imaplib.IMAP4(imap_host, imap_port)
        mail.starttls()

        mail.login(username, password)
        logger.info('Successfully logged in')
        mail.select(f'"{mailbox}"')
        return mail
    except imaplib.IMAP4.error:
        logger.error(f'Failed to log in. Please check user credentials.\n{handle_exception()}')
        sys.exit(1)
    except gaierror:
        logger.error(f'Failed to log in. Please check host name.\n{handle_exception()}')
        sys.exit(1)

def fetch_email_header(mail, uid, retry_count, timeout_limit, timeout_wait, logger):
    while retry_count <= timeout_limit:
        if retry_count > 0:
            logger.info(f'Retrying ({retry_count}/{timeout_limit})...')
            time.sleep(timeout_wait)

        try:
            status, response = mail.uid('FETCH', uid, '(BODY.PEEK[HEADER])')
            if status == 'OK':
                logger.info(f'Successfully fetched header for UID {uid}')
                return response[0][1]
            else:
                logger.warning(f'Problem fetching header for UID {uid}. Response: {response}')
        except imaplib.IMAP4.abort:
            logger.error(f'Connection closed by server while fetching UID {uid}.\n{handle_exception()}')
        except Exception:
            logger.error(f'Unexpected error while fetching UID {uid}.\n{handle_exception()}')

        retry_count += 1

    logger.error(f'Failed to fetch UID {uid} after {timeout_limit} retries.')
    return None

def load_store(data_path, logger):
    if os.path.isfile(data_path):
        with open(data_path, 'r') as file:
            store = json.load(file)
        logger.info(f'Loaded {len(store)} emails from existing store')
        return store
    else:
        logger.info('No existing store found. Starting fresh.')
        return {}

def save_store(store, data_path):
    with open(data_path, 'w') as file:
        json.dump(store, file, sort_keys=True)

def is_valid_port(value):
    try:
        port = int(value)
        return 1 <= port <= 65535
    except ValueError:
        return False

def parse_arguments():
    if len(sys.argv) < 4:
        print('Usage: python fetch.py <imap_host> <username> <password> [mailbox] [imap_port]')
        sys.exit(1)

    imap_host = sys.argv[1]
    username = sys.argv[2]
    password = sys.argv[3]
    mailbox = 'Inbox'
    imap_port = 993

    if len(sys.argv) > 4:
        if is_valid_port(sys.argv[4]):
            imap_port = int(sys.argv[4])
        else:
            mailbox = sys.argv[4]

    if len(sys.argv) > 5:
        if is_valid_port(sys.argv[5]):
            imap_port = int(sys.argv[5])
        else:
            print('Error: Invalid port number provided.')
            sys.exit(1)

    return imap_host, username, password, mailbox, imap_port

def main():
    imap_host, username, password, mailbox, imap_port = parse_arguments()

    data_path = os.path.join(username, 'data_unseen.json')
    log_path = os.path.join(username, 'fetch_unseen.log')
    timeout_wait = 30
    timeout_limit = 3

    ensure_directory_exists(username)
    logger = setup_logger(log_path)

    store = load_store(data_path, logger)
    previous_store_count = len(store)

    mail = initialize_imap(imap_host, imap_port, username, password, mailbox, logger)

    logger.info(f'Fetching unread emails from mailbox: {mailbox}...')
    _, data = mail.uid('SEARCH', None, '(SEEN)')
    unread_uids = set(data[0].split())
    new_uids = unread_uids - store.keys()

    logger.info(f'{len(unread_uids)} unread emails found, {len(new_uids)} new to fetch.')

    start_time = time.time()

    for uid in new_uids:
        header = fetch_email_header(mail, uid, 0, timeout_limit, timeout_wait, logger)
        if header:
            store[uid.decode()] = header.decode(errors='ignore')

    save_store(store, data_path)

    elapsed_time = time.time() - start_time
    logger.info(f'Completed fetching. Fetched {len(new_uids)} emails in {elapsed_time:.1f} seconds.')

if __name__ == '__main__':
    main()
