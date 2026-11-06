while True:
    directory = single_input('Please type out the directory to get to your files:')
    directory = directory.strip('"')
    if os.path.exists(directory) == True:
        if directory[-1] != '\\':
            directory = directory + '\\'
        if 'ProductData.xlsx' in os.listdir(directory) and 'MasterInvoice.pdf' in os.listdir(directory) and 'Order1.xlsx' in os.listdir(directory) and 'Order2.xlsx' in os.listdir(directory) and 'Order3.xlsx' in os.listdir(directory) and 'Order4.xlsx' in os.listdir(directory):
            break
        else:
            show_message('Error','Re-enter the directory.')
    else:
        show_message('Error!','The directory you entered does not exist.')

product_descriptions = {}
product_prices = {}

df = pd.read_excel(directory+'ProductData.xlsx')
df1 = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')

for info, row in df.iterrows():
    unit_num = row['Unit number']
    product_descriptions[unit_num] = row['Description']

for info, row in df1.iterrows():
    unit_num = row['Unit number']
    product_prices[unit_num] = row['Price per unit']

def get_order_data(invoice_num, filename, directory):
    order_data = {}
    template = openpyxl.load_workbook(directory+filename, data_only = True)
    data = template['OrderForm']
    name = data['C14'].value
    email = data['C13'].value
    date = data['C9'].value.strftime('%m/%d/%Y')
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    
    order_data['CustomerAddress1'] = data['C16'].value
    order_data['CustomerName'] = data['C14'].value
    order_data['CustomerCompany'] = data['C15'].value
    order_data['CustomerAddress2'] = f'{data['C17'].value}, {data['C18'].value} {data['C19'].value}'
    total = 0

    order_data['Description0'] = product_descriptions[data['B24'].value]
    price0 = product_prices[data['B24'].value]
    order_data['Price0'] = f'${price0:,.2f}'
    quantity0 = data['C24'].value
    order_data['Quantity0'] = str(quantity0)
    amount0 = price0*quantity0
    order_data['Amount0'] = f'${amount0:,.2f}'
    total += amount0

    if data['B25'].value is not None:
        order_data['Description1'] = product_descriptions[data['B25'].value]
        price1 = product_prices[data['B25'].value]
        order_data['Price1'] = f'${price1:,.2f}'
        quantity1 = data['C25'].value
        order_data['Quantity1'] = str(quantity1)
        amount1 = price1*quantity1
        order_data['Amount1'] = f'${amount1:,.2f}'
        order_data['Total'] = f'${amount1+amount0:,.2f}'
        total += amount1

    if data['B26'].value is not None:
        order_data['Description2'] = product_descriptions[data['B26'].value]
        price2 = product_prices[data['B26'].value]
        order_data['Price2'] = f'${price2:,.2f}'
        quantity2 = data['C26'].value
        order_data['Quantity2'] = str(quantity2)
        amount2 = price2*quantity2
        order_data['Amount2'] = f'${amount2:,.2f}'
        total += amount2

    order_data['Total'] = f'${total:,.2f}'
    
    return order_data, name, email, date

def create_invoices(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "thomas.apke@gmail.com"

    # List of recipients

    recipients = email

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

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

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()
send_email(8515, name ='Thomas Holt', email = 'thomas.apke@gmail.com', date = '11/13/2023', directory = directory)

filenames = os.listdir(directory)
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoices(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num += 1
    else:
        pass