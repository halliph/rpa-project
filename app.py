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
    directory = single_input('Please input the directory for ProductData.xlsx, MasterInvoice.pdf, and Order')
    if directory[-1] != '\\':
        directory = (f'{directory}\\')
    files_in_dir = os.listdir(directory)
    if 'ProductData.xlsx' in files_in_dir and 'MasterInvoice.pdf' in files_in_dir and 'Order1.xlsx' in files_in_dir and 'Order2.xlsx' in files_in_dir and 'Order3.xlsx' in files_in_dir and 'Order4.xlsx' in files_in_dir:
        break

    else:
        show_message('Error','Please re-enter the directory')

print(directory)



product_descriptions = {}
product_prices = {}
prod_descrips = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Descriptions')
for index,row in prod_descrips.iterrows():
    descrip = row['Description']
    unnum = row['Unit number']
    product_descriptions[unnum] = descrip 
product_descriptions


prod_price = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Prices')
for index,row in prod_price.iterrows():
    price = row['Price per unit']
    unnum = row['Unit number']
    product_prices[unnum] = price
product_prices



def get_order_data(invoice_num,filename,directory):
    total = 0
    order = openpyxl.load_workbook(directory + filename)
    selected_sheet = order['OrderForm']
    unit0 = selected_sheet['B24'].value
    unit1 = selected_sheet['B25'].value
    unit2 = selected_sheet['B26'].value
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value

    order_data = {} 
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = date
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = (f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {str(selected_sheet['C19'].value)}')

    if unit0 is not None:
        order_data['Description0'] = product_descriptions.get(unit0)
        price0 = product_prices.get(unit0)
        order_data['Price0'] = f'${price0:,.2f}'
        quantity0 = selected_sheet['C24'].value
        order_data['Quantity0'] = quantity0
        amount0 = price0 * quantity0
        order_data['Amount0'] = f'${amount0:,.2f}'
        total = total + amount0
    if unit1 is not None:
        order_data['Description1'] = product_descriptions.get(unit1)
        price1 = product_prices.get(unit1)
        order_data['Price1'] = f'${price1:,.2f}'
        quantity1 = selected_sheet['C25'].value
        order_data['Quantity1'] = quantity1
        amount1 = price1 * quantity1
        order_data['Amount1'] = f'${amount1:,.2f}'
        total = total + amount1
    if unit2 is not None:
        order_data['Description2'] = product_descriptions.get(unit2)
        price2 = product_prices.get(unit2)
        order_data['Price0'] = f'${price2:,.2f}'
        quantity2 = selected_sheet['C26'].value
        order_data['Quantity2'] = quantity2
        amount2 = price2 * quantity2
        order_data['Amount2'] = f'${amount2:,.2f}'
        total = total + amount2
    total = f'${total:,.2f}'
    order_data['Total'] = total

    return order_data, name, email, date 



def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf',directory + (f'Invoice {invoice_num}.pdf'), order_data)


gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()


def send_email(invoice_num, name, email, date, directory):
    gmail_user = "Input your email"
    

    # List of recipients

    recipients = [email]

    # Email subect

    subject = (f'Your Order from PythonTech Co. - Invoice {invoice_num}')

    # Body of the email

    name = name
    date = date

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    Sincerely,
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com
    """

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()



filenames = os.listdir(directory)
invoicenum = 8515
for xlsx in filenames:
    if 'Order' in xlsx and 'xlsx' in xlsx:
        order_data, name, email, date = get_order_data(invoicenum,xlsx,directory)
        create_invoice(invoicenum,order_data,directory)
        send_email(invoicenum, name, email, date, directory)
        invoicenum = invoicenum + 1