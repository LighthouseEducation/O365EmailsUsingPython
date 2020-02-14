# pip install O365
from  O365 import Account, MSOffice365Protocol
# use csv file for names, email addresses and pdf attachments
import csv
# protocol may not be needed. If you delete it use: account = Account(credentials)
protocol = MSOffice365Protocol()
# register app with Azure, get client ID and client secret
credentials = ( 'clientId','clientSecret')
# create account, with credentials
account = Account(credentials, protocol=protocol)
# authenticate your account
if account.authenticate(scopes=['basic','message_all']):
    # confirm authentication with message
    print('authenticated')
# confirm new message is working
print('new message beginning')
# open your csv file
with open('example.csv', newline='') as csvfile:
    # confirm it opened
    print('opened csv')
    # read csv file
    readCSV = csv.reader(csvfile)
    # define the header row
    header = next(readCSV)
    # loop through each csv row, indicate which data you will need
    for name, email1, email2, pdf in readCSV:
        # establish which mailbox you will send from
        # this uses browser authentication from the console -> copy and past authenticated link
        mailbox = account.mailbox()
        # start new message
        message = mailbox.new_message()
        # attach email signature image for inline image
        message.attachments.add('test-logo.png')
        # assign it as the first attachment item
        logo = message.attachments[0]
        # define the image as inline
        logo.is_inline = True
        # define what type of image file it is
        logo.content_id = 'image.png'
        # print the name and emails to send to and the path the attached pdf
        print(name, email1,email2,pdf)
        # add the emails you will send the message to
        message.to.add([email1,email2])
        # enter your message subject
        message.subject = name + ''' - Semester 1 Report'''
        # create your message body. you can use text. or html
        message.body = '''
            <p>Email Greeting</p>
            <p>1st paragraph</p>
            <p>Sincerely,</p>
            <p>Name Here</p>
            <p>
                <img src="cid:image.png">
            </p>
        '''
        # attach your pdf file
        message.attachments.add(pdf)
        # send your message
        message.send()
        # print to confirm your message sent
        print('message sent')
