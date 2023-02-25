# Microsoft-API-attachment-downloader


This is my simple solution to download mail attachments from Microsoft cloud using the Microsoft Graph REST API. 

This was basically designed to run daily, and to automatically download attachments from specific emails identified by subject and sender. 

Since it's intented to be run daily, it has also has sophisticated logging feature, and email notification function in case of any errors. It works with SMTP. To make it work you need to give your SMTP parameters in the mail_config file.

To use this solution, you must first register an app on Azure using your Microsoft account. Once you have registered your app, you can use its credentials, and paste them in the main_config.json file. 
To save the config files correctly, follow the instructions below:

- Rename the mail_config_EMPTY.json file to mail_config.json.
- Rename the main_config_EMPTY.json file to main_config.json.

