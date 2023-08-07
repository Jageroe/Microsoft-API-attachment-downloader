import requests
import json
import os
import argparse
import sys, os, functools
from datetime import date, timedelta
import logging
import msal
from sendmail import sending_mail


def parse_args(args:str):

    """
    Parse command-line arguments.

    Args:
        args (str): Command-line arguments as a string.

    Returns:
        argparse.Namespace: Parsed command-line arguments.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--name",
        type=str,
        required=True,
        help="The name of the running instance. This has only purpose for logging. "
    )


    parser.add_argument(
        "--mailsearch",
        type=str,
        required=True,
        help="You need to pass a KQL query to search specific messages. This use the Micosoft's KQL syntax"
    )

    # parser.add_argument(
    #     "--numofattachments",
    #     type=int,
    #     required=True,
    #     help="The expected number of attachments"
    # )

    parser.add_argument(
        "--savedir",
        type=str,
        required=True,
        help="The path of the folder where the attachments will be saved"
    )


    return parser.parse_args(args)


def set_logger(name:str) -> logging.Logger:

    """
    Method to setup the logger object. 
    It also does some basic configuration, such as define the level and create handlers

    Args:
        name (str): Name of the logger. This will contained in the log file's name

    Returns:
        The configured logger object.
    
    """

    # setting up the logging object
    logger = logging.getLogger()

    # setting up a handler for the terminal messages
    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)

    # setting up a handler for the logfile
    file_handler = logging.FileHandler(f'log/{name}_log.log')
    file_handler.setLevel(logging.INFO)

    # The level = logging.DEBUG is just for the basic config. The handlers
    # will overwrite this. Therefore we can change the logging levels
    # in the handler objects!
    logging.basicConfig(
        level=logging.DEBUG,
        format='[%(asctime)s] [%(module)s] [%(levelname)s] - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=[
            stream_handler,
            file_handler
        ]
    )

    return logger

def log(func):
    """
    Decorator for logging
    """

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            logger.info(f"{func.__name__} method ran successfully!")

            return result
        except Exception as ex:
            logger.exception(f"Exception raised in {func.__name__}. exception: {str(ex)}")

            raise ex
    return wrapper


def error_mail(func):
    """
    Decorator for sending error emails.
    """

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            return result
        
        except Exception as ex:
            sending_mail(
                subject='ERROR - Attachment downloader',
                message=f"Exception raised in {func.__name__}. exception: {str(ex)}"
                )
            
            raise ex
    return wrapper

@error_mail
@log
def read_config(path:str) -> dict:

    """
    Read the config file and return a dictionary containing the config data.

    Args:
        path (str): Path to the config file.

    Returns:
        dict: Config data as a dictionary.
    """

    with open(path, 'r', encoding='utf-8') as file:
        return json.loads(file.read())


class MicrosoftGraphApiConnection:

    """
    Class that acquires the token and stores it for the API requests.
    """

    @error_mail
    @log
    def __init__(self, client_id: str, authority: str, 
                 endpoint: str, scope: list, user_to_read: str,
                 user_password:str)-> None:
        
        self.endpoint = endpoint
        app = msal.ConfidentialClientApplication(client_id, authority=authority)

        token = app.acquire_token_by_username_password(
            username=user_to_read,
            password=user_password,
            scopes=scope
        )

        self.token = token

        # if the acquired token doesn't have the acces_token key the try statement will raise an error
        self.access_token = self.token['access_token']

        logger.info("Token has been successfully acquired!")
        self.headers={'Authorization': 'Bearer ' + self.access_token}


    @error_mail
    @log
    def get_mails(self, mail_query:str = None)-> list:

        """
        Returns a dict with the found messages containing the id, the subject, and the sender's address. 
        parameters:
        search_query: 
            Use this, if you wanna search specific messages. 
            It uses the microsoft KQL syntax.
            Further information : https://learn.microsoft.com/en-us/graph/search-query-parameter?tabs=http
        """

        
        response = requests.get(
            url=f'{self.endpoint}/me/messages{mail_query if mail_query else ""}',
            headers=self.headers
        )
        response.raise_for_status()
        messages = []

        for i, mail in enumerate(response.json()['value']):

            messages.append({})
            messages[i]['id'] = mail['id']
            messages[i]['subject'] = mail['subject']
            messages[i]['from'] = mail['from']['emailAddress']['address']

        logger.info(f'{len(messages)} messages has been found: {[message for message in messages]}')
        return messages


        
    
    @error_mail
    @log
    def download_attachments(self, message_id: str, save_path: str) ->int:

        """
        It downloads the attachments of the given mail
        parameters:
        message_id:
            the id of the message from which we want to download the attachments
        save_path:
            here will be the attachments downloaded

        returns the number of the attachments which have been downloaded
        
        """

        # getting the ids of the attachments
        response_mail = requests.get(
            url=f'{self.endpoint}/me/messages/{message_id}/attachments', 
            headers=self.headers
        )
        response_mail.raise_for_status()


    # requesting the attachments with a loop
        for attachment in response_mail.json()['value']:
            attachment_id = attachment['id']
            attachment_name = attachment['name']

            response_attachment = requests.get(
                f'{self.endpoint}/me/messages/{message_id}/attachments/{attachment_id}/$value', 
                headers=self.headers
            )
            response_attachment.raise_for_status()

            # saving the file
            with open(os.path.join(save_path,attachment_name),'wb') as f:
                f.write(response_attachment.content)

            logger.info(f'{attachment_name} has been saved succesfully!')

        # Number of attachments
        return len(response_mail.json()['value'])

@error_mail
@log
def main():

    args = parse_args(sys.argv[1:])

    global logger
    logger = set_logger(args.name)

    # reading the config file
    config = read_config('main_config.json')

    CLIENT_ID = config['client_id']
    TENANT_ID = config['tenant_id']
    AUTHORITY = 'https://login.microsoftonline.com/' + TENANT_ID
    ENDPOINT = 'https://graph.microsoft.com/v1.0'
    SCOPE = ['https://graph.microsoft.com/.default']
    USER_TO_READ = config['user_to_read']
    USER_PASSWORD = config['user_password']

    # initialize the connection object.
    connection = MicrosoftGraphApiConnection(
            client_id = CLIENT_ID,
            authority = AUTHORITY,
            endpoint = ENDPOINT,
            scope = SCOPE,
            user_to_read = USER_TO_READ,
            user_password = USER_PASSWORD
        )


    mails = connection.get_mails(args.mailsearch)

    # loop trough the mails list, and download all the attachments
    for mail_id in [mail['id'] for mail in mails]:
        connection.download_attachments(
            message_id = mail_id,
            save_path = args.savedir
        )

if __name__ == '__main__':
    main()

