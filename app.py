import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

while True:
    directory = single_input('Enter directory: ')
    if not directory.endswith('\\'):
        directory+='\\'
    filenames = os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        break
    else:
        show_message('Error', 'Directory does not contain required files')



product_descriptions = {}
product_prices = {}

description = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
for index,row in description.iterrows():
    product_descriptions[row[1]] = row[0]

prices = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')
for index,row in prices.iterrows():
    product_prices[row[0]] = row[1]


def get_order_data(invoice_num,filename,directory):
    order_data = {}
    wb = openpyxl.load_workbook(directory+filename)
    sheet = wb['OrderForm']
    date = sheet['C9'].value.strftime('%m/%d/%Y')
    name = sheet['C14'].value
    email = sheet['C13'].value

    quantity_cell = 24
    product_data = {}
    while sheet[f'C{quantity_cell}'].value != None:
        product_data[sheet[f'B{quantity_cell}'].value] = sheet[f'C{quantity_cell}'].value
        quantity_cell+=1


    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    total = 0
    product_index = 0
    
    for un_num, quant in product_data.items():
        for un, des in product_descriptions.items():
            if un == un_num:
                order_data[f'Description{product_index}'] = des
                product_index+=1
    
    product_index = 0
    for un_num, quant in product_data.items():
        for un,price in product_prices.items():
            if un == un_num:
                order_data[f'Price{product_index}'] = f'${price:.2f}'
                order_data[f'Quantity{product_index}'] = str(quant)
                order_data[f'Amount{product_index}'] = f'${(price*quant):,.2f}'
                product_index+=1
                total += price*quant
                
    order_data['Total'] = f'${total:,.2f}'
    order_data['CustomerCompany'] = sheet['C15'].value
    order_data['CustomerAddress1'] = sheet['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerAddress2'] = f'{sheet['C17'].value}, {sheet['C18'].value}, {sheet['C19'].value}'
    
    return order_data,name,email,date


def create_invoice(invoice_num,order_data,directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)


gmail_user = "peggysnowfan@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num,name,email,date,directory):
    

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},

Attached is the invoice dated {date}.\n
Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


filenames = os.listdir(directory)
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data,name,email,date = get_order_data(invoice_num,filename,directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num+=1
