import requests
import pandas as pd
import io
from datetime import datetime
from exchangelib import Credentials, Account, Message, Mailbox, HTMLBody, FileAttachment
from dotenv import load_dotenv
from pathlib import Path
import os

load_dotenv(Path('.') / '.env')
my_user = os.getenv('user_name')
my_pass = os.getenv('password')
my_account = os.getenv('account')

credentials = Credentials(username=my_user, password=my_pass)
account = Account(my_account, credentials=credentials, autodiscover=True)


url = "https://shafafiyaportal.doh.gov.ae/dictionary/DrugCoding/Drugs.xlsx"

r = requests.get(url, allow_redirects=True)
xl = io.BytesIO(r.content)
df = pd.read_excel(xl, sheet_name='Drugs')
date_today = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
df['Today Date'] = date_today
df['Today Date'] = pd.to_datetime(df['Today Date'])
df['delta'] = (df['Today Date'] - df['Delete Effective Date']).dt.days
pd.to_numeric(df['delta'])
df = df[df['delta'].notna()]
df = df.drop(df[(df.delta < -30.00)].index)
df = df.drop(df[(df.delta >= 0.00)].index)
df = df.drop(columns=['Today Date', 'delta'])

df.to_excel('Expiring Grace Codes.xlsx', sheet_name='Drugs', index=False)


# m = Message(
#     account=account,
#     subject='Shafafiya Codes Expiring in 30 Days',
#     body=HTMLBody('''
#     <html>
#         <body style="font-family:Consolas; color:#0b3a5a">
#             <p>Dear Colleagues,</p>
#             <p>Kindly find the attached Shafafiya code expiring in the coming 30 days.</p>
#             <p>Regards,</p>
#         </body>
#     </html>
#     '''),
#     to_recipients=[Mailbox(email_address='email@org.ae')
#     ],

#     cc_recipients=['email@org.ae', 'email@org.ae'], 
# )
# my_file_1 = FileAttachment(name='Expiring Grace Codes.xlsx', content=open('Expiring Grace Codes.xlsx', 'rb').read())
# m.attach(my_file_1)
# m.send()
# os.remove('Expiring Grace Codes.xlsx')
