#Import Library
import xlrd
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#After that, we need a path of our Excel file as well as all information that we need from that file 
# (name, email, city, paid, amount).
path = "clients.xlsx"
openFile = xlrd.open_workbook(path)
sheet = openFile.sheet_by_name('clients')
#I put the email, amount, and the name of clients that owe money in three different lists. 
# And from that I check if cllient is paid or not.
mail_list = []
amount = []
name = []
for k in range(sheet.nrows-1):
    client = sheet.cell_value(k+1,0)
    email = sheet.cell_value(k+1,1)
    paid = sheet.cell_value(k+1,3)
    count_amount = sheet.cell_value(k+1,4)
    if paid == 'No':
        mail_list.append(email) 
        amount.append(count_amount)
        name.append(client)
#After that, we need to focus on sending emails.

email = 'some@gmail.com' 
password = 'pass' 
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(email, password)
#We need to get the index so then we can find the name of the person.
for mail_to in mail_list:
    send_to_email = mail_to
    find_des = mail_list.index(send_to_email) 
    clientName = name[find_des] 
    subject = f'{clientName} you have a new email'
    message = f'Dear {clientName}, \n' \
              f'we inform you that you owe ${amount[find_des]}. \n'\
              '\n' \
              'Best Regards' 

    msg = MIMEMultipart()
    msg['From '] = send_to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
    text = msg.as_string()
    print(f'Sending email to {clientName}... ') 
    server.sendmail(email, send_to_email, text)
#And last we need to be sure that be sure everything it's ok.
server.quit()
print('Process is finished!')
time.sleep(10) 