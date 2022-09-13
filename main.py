from datetime import *
from datetime import timedelta
import pandas as pd
import smtplib
import numpy as np
from email.message import EmailMessage
import openpyxl
import time

while(True):


    df = pd.read_excel(r'C:\Users\User\Downloads\Recruiters Comments Consolidate 19th July 2022.xlsx',sheet_name='Overall',engine='openpyxl')
    df['Req'] = df['Req'].fillna(0)

    df = df.loc[df['Req'] != 0]

    df['Email'] = df['Email'].fillna(0)
    df = df.loc[df['Email'] != 0]


    dat = np.datetime64('today')

    req = np.int64(df['Req'].values)
    Date = df['Date'].values
    Email = df['Email'].values
    recruiter = df['Recruiter Name '].values
    job_req_id = df['Job Req ID'].values
    job_title = df['Job Title'].values
    job_grade = df['Job Grade'].values
    hm = df['Hm Name'].values
    ageing = df['Revised Reqs Approved Ageing (Actual - Hold days)'].values
    zipped = zip(req,Date,Email,recruiter,job_req_id,job_title,job_grade,hm,ageing)

    #creating an smtp server 
    server = smtplib.SMTP_SSL('smtp.gmail.com',465) 
    server.starttls
    server.login('robin.kiliyilathu@flexability.in','Robin$1999')

    for (a,b,c,d,e,f,g,h,i) in zipped:
    #   a = np.int64(a)   
        expiry = np.timedelta64(a,"D")
        tot = b + expiry
        delta = dat-tot
        days = delta.astype('timedelta64[D]')
        msg = EmailMessage()
        msg['Subject'] = 'Req Closing Status'
        msg['From'] = 'robin.kiliyilathu@flexability.in'
        msg['To'] = c
        msg['Cc'] = 'robinabrahamra99@gmail.com'
        if days > 0:      
            msg.set_content('Hello {},\n\nThis is just to remind you that for Req ID-{}, Job Title-{} is still Open and not in Offer Stage.\n\nPlease plan accordingly to make sure things turns to Offer.\n\nMore Info:\n\nAgeing: {}, HM: {}, Job Grade: {}'.format(d,e,f,i,h,g))
            server.send_message(msg)
        elif days < 0:
            msg.set_content('You have {} days remaining for req Closing'.format(days*-1))
            server.send_message(msg)
    server.quit()
    print('Executed Successfully!!')
    time.sleep(30)
    
wb.save(r'C:\Users\User\Downloads\POFU -Candidate Status (Excl.Weekend).xlsx') # Path to save the file

									
									
									
									
									
					
									
									
									
									
									
									
									
