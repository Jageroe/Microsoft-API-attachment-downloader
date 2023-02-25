import requests
import json
import os
from datetime import date, timedelta
import logging
import msal
from sendmail import sending_mail

# setting up the logging object
logger = logging.getLogger()

# setting up a handler for the terminal messages
stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.INFO)

# setting up a handler for the logfile
file_handler = logging.FileHandler('log.log')
file_handler.setLevel(logging.INFO)

# The level = logging.DEBUG is just for the basic config. The handlers
# will overwrite this. Therefore we can change the logging levels
# in the handler objects!
logging.basicConfig(
    level = logging.DEBUG,
    format = '[%(asctime)s] [%(module)s] [%(levelname)s] - %(message)s',
    datefmt = '%Y-%m-%d %H:%M:%S',
    handlers = [
        stream_handler,
        file_handler
    ]
)

class MicrosoftGraphApiConnection:

    """
    Class that acquires the token and stores it for the API requests
    """

    def __init__(self, client_id: str, authority: str, 
                 endpoint: str, scope: list, user_to_read: str,
                 user_password:str)-> None:
        
        self.endpoint = endpoint
        app = msal.ConfidentialClientApplication(client_id, authority = authority)

        try:

            token = app.acquire_token_by_username_password(
                username = user_to_read,
                password = user_password,
                scopes = scope
            )
        
            self.token = token

            # if the acquired token doesn't have acces_token key the try statement will raise an error
            self.access_token = self.token['access_token']

            logger.info("Token has been successfully acquired!")
            self.headers={'Authorization': 'Bearer ' + self.access_token}

        except:
            logger.error(f'An error has occurred during acquiring the acces token:\n{token}')
            sending_mail(
                subject = 'ERROR - Logicort attachment donwloader',
                message = f'An error has occurred during acquiring the acces token:\n{token}'
            )
            raise Exception


    def get_mails(self, search_query: str = None)-> list:

        """
        Returns a dict with the found messages containing the id, the subject, and the sender's address. 

        parameters:

        search_query: 
            Use this, if you wanna search specific messages. 
            It uses the microsoft KQL syntax.
            Further information : https://learn.microsoft.com/en-us/graph/search-query-parameter?tabs=http

        """

        try:
            response = requests.get(
                url = f'{self.endpoint}/me/messages{search_query if search_query else ""}',
                headers = self.headers
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

        except Exception as exception:
            
            logger.error(f'An error has occurred:\n{exception}\nThe content of the response:\n{response.json()}')
            sending_mail(
                subject = 'ERROR - Attachment donwloader',
                message = f'An error has occurred:\n{exception}\nThe content of the response:\n{response.json()}'
            )
            raise Exception
        
    

    def download_attachments(self, message_id: str, save_path: str):

        """
        It downloads the attachments of the given mail

        parameters:

        message_id:
            the id of the message from which we want to download the attachments

        save_path:
            here will be downloaded the attachments
        
        """

        # getting the ids of the attachments
        try:
            response_mail = requests.get(
                url = f'{self.endpoint}/me/messages/{message_id}/attachments', 
                headers = self.headers
            )
            response_mail.raise_for_status()

        except Exception as exeption:
            logger.error(f'An error has occurred: \n {exeption}\n The content of the response:\n {response_mail.json()}')
            raise Exception
        

        # requesting the attachments with a loop
        try:
            for attachment in response_mail.json()['value']:
                attachment_id = attachment['id']
                attachment_name = attachment['name']

                response_attachment = requests.get(
                    f'{self.endpoint}/me/messages/{message_id}/attachments/{attachment_id}/$value', 
                    headers = self.headers
                )
                response_attachment.raise_for_status()

                # saving the file
                with open(os.path.join(save_path,attachment_name),'wb') as f:
                    f.write(response_attachment.content)

                logger.info(f'{attachment_name} has been saved succesfully!')

        except Exception as exeption:
            logger.error(
                f"""An error has occurred with the following attachment:\n 
                {attachment_name} - id: {attachment_id}\n {exeption}"""
            )
            
            sending_mail(
                subject = 'ERROR - Logicort attachment donwloader',
                message = f"""An error has occurred with the following attachment:\n 
                        {attachment_name} - id: {attachment_id}\n {exeption}"""
            )

            raise Exception

def main():

    # reading the config file
    with open('main_config.json', 'r', encoding='utf-8') as file:
        config = json.loads(file.read())

    CLIENT_ID = config['client_id']
    TENANT_ID = config['tenant_id']
    AUTHORITY = 'https://login.microsoftonline.com/' + TENANT_ID
    ENDPOINT = 'https://graph.microsoft.com/v1.0'
    SCOPE = ['https://graph.microsoft.com/.default']
    USER_TO_READ = config['user_to_read']
    USER_PASSWORD = config['user_password']
    SAVE_PATH = config['save_path']

    # initialize the connection object. 
    connection = MicrosoftGraphApiConnection(
            client_id = CLIENT_ID, 
            authority = AUTHORITY,
            endpoint = ENDPOINT,
            scope = SCOPE, 
            user_to_read = USER_TO_READ, 
            user_password = USER_PASSWORD
        )
    

    # Here's an example below (commented out) how to use the search query parameter
    # in this example I request mails with three condition:
    # their subject have to contain the folowing: Sample-subject <yesterday's date> 
    # and they have to have attachments
    # and they are from jageroeee@gmail.com

    # yesterday = date.today() - timedelta(days = 1) 
    # yesterday = yesterday.strftime('%Y-%m-%d') 
    # search_query = f'?$search="subject:Sample-subject {yesterday} AND hasAttachments:true AND from:jageroeee@gmail.com"'


    search_query = None
    mails = connection.get_mails(search_query)

    # loop trough the mails list, and download all the attachments
    for mail_id in [mail['id'] for mail in mails]:
        connection.download_attachments(
            message_id = mail_id,
            save_path = SAVE_PATH
        )

if __name__ == '__main__':
    main()

