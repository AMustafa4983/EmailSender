import smtplib
import pandas
from email.mime.text import MIMEText
from email.mime.multipart  import MIMEMultipart
from email.mime.application import MIMEApplication
import time



li = pandas.read_excel('emails.xlsx')

# Change this
email = ''  # your email
password = ''  # generated password from google account settings (Thirdparty access)
username = ''  # your name
files = ["","",""] # Set your Attachments file names with extention


failed_to_reach = []

def send_email(subject, body, sender, recipients, password, files):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = f"{username} <{sender}>"
    msg['To'] = recipients
    msg.attach(MIMEText(body,"plain"))
    
    for file in files:
        with open(file, "rb") as f:
            attach = MIMEApplication(f.read(),_subtype="pdf")
            attach.add_header('Content-Disposition','attachment',filename=str(file))
            msg.attach(attach)
    
    smtp_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    smtp_server.login(sender, password)
    smtp_server.sendmail(sender, recipients, msg.as_string())
    smtp_server.quit()

print(li['emails'])

for i in range(len(li)):
    # change message title from here
    subject = 'Enthusiastic Data Scientist Seeking New Opportunities – Immediate Availability'    
    # message here and don't change the text between curly brackets
    message = f'''
Dear HR Team,

(Details about you for example:
I am Abdelrahman Mohamed, a Data Scientist with a Computer Science degree from Zagazig University and specialized certifications in Machine Learning and Data Analysis from Udacity. My experience includes impactful roles at Harvard University and Johns Hopkins University, where I honed my skills in Python, TensorFlow, and advanced machine learning.

Please review my LinkedIn profile for a detailed overview of my background: (Your Linkedin account)

I am currently in Dubai on a 2-month visit visa and can start a new role with a 7-day notice period. I am confident in my ability to contribute effectively to your team.
)

Enclosed is my resume for your consideration. I look forward to the possibility of working with your esteemed organization.

Thank you for your time and consideration.

Best Regards,

Your Name
Contact information
Email

'''
    dest = li['emails'][i]
    try:
        send_email(subject, message, email, dest, password,files)
        print(f'Mail Sent to {dest}')
    except:
        print(f'the mail didn\'t reach {dest}')
        failed_to_reach.append(dest)
        continue

    # Check if 50 emails have been sent, then wait for 5 minutes
    if (i + 1) % 50 == 0:
        print('Waiting for 5 minutes...')
        time.sleep(300)  # 5 minutes


if failed_to_reach:
    pandas.DataFrame({'emails':failed_to_reach}).to_csv('failed_to_reach.csv')
