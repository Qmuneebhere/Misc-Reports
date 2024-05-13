from extract_msg import Message
import os
import glob

def read_msg_file(msg_file_path, count):

    # Open the MSG file
    
    msg = Message(msg_file_path)

    # Extract information from the MSG file

    subject = msg.subject
    sender = msg.sender
    recipients = [recipients.email for recipients in msg.recipients]
    date = msg.date.date()
    body = msg.body
    attachments = msg.attachments

    # Print or process the extracted information as needed

    print('\n----------------------------------------------\n')

    print("Subject:", subject)
    print("Sender:", sender)
    print("Recipients:", recipients)
    print("Date:", date)
    print("Body:", body)

    folderPath = f'{output_dir}{date}-Email-{count}'
    imagePath = f'{folderPath}\\Images\\'

    os.makedirs(folderPath)
    os.makedirs(imagePath)

    for attachment in attachments:

        print(attachment.longFilename)

        # Save the image attachment to a file
        
        attachment.save(customPath=imagePath)


pattern = 'C:\\Users\\MUNEEB\\Outlook\\Files\\*.msg'
output_dir = 'C:\\Users\\MUNEEB\\Outlook\\Exports\\'

fileList = glob.glob(pattern)

count = 0

for file in fileList:

    read_msg_file(file, count)

    count += 1

