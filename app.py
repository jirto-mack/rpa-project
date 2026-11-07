# IMPORT PACKAGES IN THIS CELL

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
    directory = single_input("Please paste the directory where the Product Data, Invoices, and Master File are located.")
    
    if (directory[-1:] != "\\"):
        directory = directory + "\\"

    try:
        files = os.listdir(directory)
        if(('ProductData.xlsx' in files) and ('MasterInvoice.pdf' in files) and ('Order1.xlsx' in files) and ('Order2.xlsx' in files) and ('Order3.xlsx' in files) and ('Order4.xlsx' in files) ):
            break
        else: 
            show_message("Error", "The directory listed does not contain the proper files. Please re-enter the directory.")
    except:
        show_message("Error", "The directory provided is faulty, please try again.")

product_description = {}
product_prices = {}

df_desc = pd.read_excel(directory+"ProductData.xlsx", sheet_name = 'Descriptions')
for index, row in df_desc.iterrows():
    product_description[row["Unit number"]] = row.Description

df_prices = pd.read_excel(directory+"ProductData.xlsx", sheet_name = 'Prices')
for index, row in df_prices.iterrows():
    product_prices[row["Unit number"]] = row["Price per unit"]

def get_order_data(invoice_num, filename, directory):
    order_dict = {}
    total = 0

    order = openpyxl.load_workbook(directory+filename)
    selected_sheet = order["OrderForm"]
    name = selected_sheet["C14"].value
    email = selected_sheet["C13"].value
    date = selected_sheet["C9"].value
    date = date.strftime("%m/%d/%Y")

    
    order_dict["InvoiceNum"] = str(invoice_num)
    order_dict["Date"]  = date
    order_dict["CustomerAddress1"] = selected_sheet["C16"].value
    order_dict['CustomerName'] = name 
    order_dict['CustomerCompany'] = selected_sheet['C15'].value
    order_dict['CustomerAddress2'] = selected_sheet['C17'].value + ", " + selected_sheet['C18'].value + " " + str(selected_sheet["C19"].value)
    
    if selected_sheet['B24'].value is not None:
        order_dict['Price0'] = f'${product_prices[selected_sheet['B24'].value]:,.2f}'
        order_dict['Quantity0'] = str(selected_sheet["C24"].value)
        order_dict['Amount0'] = f'${product_prices[selected_sheet['B24'].value]*selected_sheet["C24"].value:,.2f}'
        order_dict['Description0']= product_description[selected_sheet['B24'].value]
        total += product_prices[selected_sheet['B24'].value]*selected_sheet["C24"].value
    if selected_sheet['B25'].value is not None:
        order_dict['Price1'] = f'${product_prices[selected_sheet['B25'].value]:,.2f}'
        order_dict['Quantity1'] = str(selected_sheet["C25"].value)
        order_dict['Amount1'] = f'${product_prices[selected_sheet['B25'].value]*selected_sheet["C25"].value:,.2f}'
        order_dict['Description1']= product_description[selected_sheet['B25'].value]
        total += product_prices[selected_sheet['B25'].value]*selected_sheet["C25"].value
    if selected_sheet['B26'].value is not None:
        order_dict['Price2'] = f'${product_prices[selected_sheet['B26'].value]:,.2f}'
        order_dict['Quantity2'] = str(selected_sheet["C26"].value)
        order_dict['Amount2'] = f'${product_prices[selected_sheet['B26'].value]*selected_sheet["C26"].value:,.2f}'
        order_dict['Description2']= product_description[selected_sheet['B25'].value]
        total += product_prices[selected_sheet['B26'].value]*selected_sheet["C26"].value

    order_dict['Total'] = f'${total:,.2f}'
    return order_dict, name, email, date

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice (invoice_num, order_data, directory):

    create_pdf(directory+'MasterInvoice.pdf', directory+f'Inovice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)
# INSERT ANSWER FOR QUESTION 5 HERE
def send_email (invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "themichaelwinward@gmail.com"
    gmail_password = open('gmail_password.txt').read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f""" 
    Dear {name},
    
    Attached is the invoice dated {date}.

    Thank you for your business.
    
    Python TechCo.
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

    path_to_file1 = directory+f'Inovice {invoice_num}.pdf'

    messsage = add_attachment(path_to_file1, message)
        
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, name, email, date, directory)
filenames = os.listdir(directory)
invoice_num = 8515
for filename in filenames:
    if "Order" in filename and "xlsx" in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num += 1 


