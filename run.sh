#!/bin/bash

# yesterday's date
yesterday="$(date --date "-1 days" +'%Y-%m-%d')"

# activating the virtual env
source /home/oli/python/Microsoft-API-attachment-downloader/venv-Microsoft-API-attachment-downloader/bin/activate

# it runs the python file with the necessary arguments
# In this example it will search the message with 
#   - Sample-subject <Yesterday'se dat in YYYY-mm-dd format>
#   - Having an attachment
#   - The sender is: jageroee@gmail.com

python3 main.py \
    --mailsearch "?\$search=\"subject:Sample-subject $yesterday AND hasAttachments:true AND from:jageroee@gmail.com\"" \
    --name "Sample" \
    --savedir "/home/oli/python/Microsoft-API-attachment-downloader/downloaded-attachments/" \
    --numofattachments 1
