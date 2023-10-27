import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
from colorama import Fore, Style
import openai

# Setup port number and server name
smtp_port = 587  # Standard secure SMTP port
smtp_server = "smtp.gmail.com"  # Google SMTP Server

# OpenAI key to access OpenAI account
openai.api_key = "" #your API key

# Password to access Gmail account
pswd = '' #Password to your gmail account

# Email id from which the mails have to be sent
email_from = "" #Your Email ID

# prints the intro of the program
def intro():
    he = '''###
    Hey want to send an EMAIL?
    Dont worry we got you!!!
    Just type the required fields below.
    '''
    for char in he:
        print(Fore.BLUE + Style.BRIGHT + char, end='', flush=True)
    print()
intro()

# asks the user for the number of recipients to send the mail.
def no_of_recepients():
    while True:
        try:
            x = input(Fore.GREEN + Style.BRIGHT + 'No of recipients you want to send - ')
            if x.isdigit():
                pass
            else:
                raise ValueError
            return int(x)
        except ValueError:
            print(Fore.RED + Style.BRIGHT + 'ERROR, please enter again')
recipient_count = no_of_recepients()


# asks the user for the emailid's of the respective recipients
def taking_recepient_emailid():
    time.sleep(2)
    d = '@gmail.com'
    email_list = []

    # for single recipient
    if recipient_count == 1:
        for i in range(1, recipient_count + 1):
            def tried1():
                while True:
                    try:
                        t = input(Fore.GREEN + Style.BRIGHT + f'Recipient - ') + d
                        for j in t:
                            if j.isdigit() or j.isalpha() or j == '@' or j == '.' :
                                pass
                            else:
                                raise ValueError
                        return t
                    except ValueError:
                        print(Fore.RED + Style.BRIGHT + 'ERROR, please enter again')
            t1 = tried1()
            email_list.append(t1)

    # for multiple recipients
    else:
        for i in range(1, recipient_count + 1):
            def tried2():
                while True:
                    try:
                        g = input(Fore.GREEN + Style.BRIGHT + f'Recipient {i} - ') + d
                        for v in g:
                            if v.isdigit() or v.isalpha() or v == '@' or v == '.':
                                pass
                            else:
                                raise ValueError
                        return g
                    except ValueError:
                        print(Fore.RED + Style.BRIGHT + 'ERROR, please enter again')
            t2 = tried2()
            email_list.append(t2)
    return email_list
recipient_list = taking_recepient_emailid()


print()
print(Fore.RED + "Connecting to server...")
print(Fore.RED + "Successfully connected to server....", time.sleep(2))
print()



# creates the subject for the mail.
def mail_subject():
    subject = input(Fore.CYAN + Style.BRIGHT + 'write the subject of the email - ')
    return subject
subject_of_mail = mail_subject()


# creates the body for the mail
def mail_body():
    while True:
        try:
            print()
            input1 = input(Fore.CYAN + Style.BRIGHT + '''Do you want to write email on your own or use chat gpt
(say yes to not use gpt and say gpt to use chatgpt) - ''')
            print()
            if input1 == 'yes' or input1 == 'YES' or input1 == 'Yes':
                response1 = input(Fore.GREEN + Style.BRIGHT + 'Write your body - ')
                return response1
            elif input1 == 'gpt' or input1 == 'GPT' or input1 == 'Gpt':
                def chatgpt():
                    while True:
                        try:
                            print()

                            # telling chatgpt to generate the body of the email
                            prompt = input(Fore.CYAN + Style.BRIGHT + 'Write your prompt - ')
                            completion = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=[{"role": "user", "content": prompt}], n=3, stop=None, temperature=0)
                            response2 = completion.choices[0].message.content

                            print()
                            print(response2)
                            print()
                            x = input('Are you satisfied with the result(Yes/No) - ')
                            if x == 'Yes' or x == 'YES' or x == 'yes':
                                pass
                            elif x == 'NO' or x == 'No' or x == 'no':
                                raise ValueError
                            else:
                                raise TypeError
                            return response2
                        except TypeError:
                            print(Fore.RED + Style.BRIGHT + 'ERROR, please enter again')
                        except ValueError:
                            print(Fore.RED + Style.BRIGHT + 'Sorry for the inconvenience caused. please type your prompt again')
                response2 = chatgpt()
                return response2
            else:
                raise ValueError
        except ValueError:
            print(Fore.RED + Style.BRIGHT + 'ERROR, please enter again')
body_of_mail = mail_body()
print()


# Attaching files to the email
fileattachment = input(Fore.MAGENTA + Style.BRIGHT + 'Do you want to send an attachment(Yes/No) - ')
if fileattachment == 'yes':
    # shows the files available to send on the console
    def show_files():
        print()
        print(Fore.MAGENTA + Style.BRIGHT + 'These are the available files you can send:- ')
        l = 1
        file_list = {}
        for i in os.listdir(r''):  # add your own directory from which the files are being displayed
            file_list[l] = i
            l = l + 1
        time.sleep(2)
        po = 1
        for g in file_list:
            print(f'{po}.', file_list[g])
            po = po + 1
        print()
        return file_list
    filelist = show_files()
    time.sleep(2)

    # asks the user the files he/she wants to send from the list shown on the console
    def file_you_want_to_send():
        while True:
            try:
                fileinput = input(Fore.GREEN + Style.BRIGHT + '''which files do you want to send
        (input in the form - fileno,fileno,....) - ''')
                file_inp = fileinput.split(',')
                if len(file_inp) <= len(filelist):
                    for i in file_inp:
                        if i.isdigit() and int(i) > 0:
                            pass
                        else:
                            raise ValueError
                        return file_inp
                else:
                    raise ValueError
            except ValueError:
                print(Fore.RED + Style.BRIGHT + 'ERROR, please enter again')
    file_input = file_you_want_to_send()
else:
    file_input = 'false'


# this function sends email to the respective recipients in a for loop.
def send_emails(email_list,subject):
    for person in email_list:
        # Making the structure of the body of the mail
        body = f"""
Dear: {person}

      {body_of_mail}"""

        # making a MIME object to define parts of the email
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = person
        msg['Subject'] = subject

        # Attaching the body to the mail
        msg.attach(MIMEText(body, 'plain'))
        print()

        # if the user has said 'yes' to attaching of files
        if file_input != 'false':
            # attaching multiple files
            def attaching_the_file():
                FILELIST = []
                for j in file_input:
                    FILELIST.append(filelist[int(j)])
                for v in FILELIST:
                    filename = f'' # file path in this format (Keeep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{v}

                    # Open the file in python as a binary
                    attachment = open(filename, 'rb')  # r for read and b for binary

                    # Encode as base 64
                    attachment_package = MIMEBase('application', 'octet-stream')
                    attachment_package.set_payload((attachment).read())
                    encoders.encode_base64(attachment_package)
                    attachment_package.add_header('Content-Disposition', "attachment; filename= " + v)
                    msg.attach(attachment_package)

                # Cast as string
                text = msg.as_string()
                # msg is a MIMEMultipart object that i have created to compose the email.
                # It includes the sender, recipient, subject, body, and any attachments.
                # However, for the email to be sent, it needs to be converted into a specific
                # format that adheres to the email standards.
                # This is where the as_string() method comes into play.
                return text

            text = attaching_the_file()

        # if the user has said 'no' to attaching of files
        else:
            # we are rewriting the same body cus giving text variable to a string
            # rewrites the whole code without the subject and the body.

            # Making the outline of the body of the email
            body = f"""
            Dear: {person}

            {body_of_mail}"""

            # Making a MIME object to define parts of the email
            msg = MIMEMultipart()
            msg['From'] = email_from
            msg['To'] = person
            msg['Subject'] = subject

            # Attaching the body of the message
            msg.attach(MIMEText(body, 'plain'))
            text = msg.as_string()

        # Connecting to the server
        TIE_server = smtplib.SMTP(smtp_server, smtp_port)
        TIE_server.starttls()
        TIE_server.login(email_from, pswd)

        # Sending emails to the respective "person" as list is iterated
        print(Fore.RED + Style.BRIGHT + f"Sending email to: {person}")
        TIE_server.sendmail(email_from, person, text)
        print(Fore.RED + Style.BRIGHT + f"Email sent to: {person}")

    # Closing the port
    TIE_server.quit()
send_emails(recipient_list, subject_of_mail)


print()
print(Fore.RED + Style.BRIGHT + 'Finished')