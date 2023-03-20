# Microsoft-API-attachment-downloader


This is my simple solution for downloading email attachments from Microsoft cloud using the Microsoft Graph REST API. t is designed to run daily and automatically download attachments from specific emails identified for example by subject and sender. Since it's intented to be run daily, it also has sophisticated logging feature and email notification function in case of any errors. The mail notification feature works with SMTP, so to make it work you need to provide your SMTP parameters in the mail_config file.

To use this solution, you must first register an app on Azure using your Microsoft account. Once you have registered your app, you can use its credentials, and paste them in the main_config.json file. 
To save the config files correctly, follow the instructions below:

- Rename the mail_config_EMPTY.json file to mail_config.json.
- Rename the main_config_EMPTY.json file to main_config.json.

The main.py script can be run from the terminal using the following command:

        python3 main.py --mailsearch MAILSEARCH --savedir SAVEDIR

The following are the command-line arguments:

 **--mailsearch:**  You need to pass a KQL query to search specific messages. This use the Micosoft's KQL syntax

**--savedir:** The path of the folder where the attachments will be saved

I also created the **`run.sh`** bash script that allows you to run main.py with dinamically changing arguments such as dates.

