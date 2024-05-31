# -*- coding: utf-8 -*-
"""
Created on Tue May  7 13:29:01 2024

@author: NXP
"""
import ssl
import requests
import pandas as pd
import gspread as gs
import df2gspread as d2g
import gspread_dataframe as gd
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import smtplib
import time
from datetime import datetime

credentials ={

}



gc = gs.service_account_from_dict(credentials)

def export_to_sheets(file_name, sheet_name):
    ws = gc.open(file_name).worksheet(sheet_name)
    data = ws.get_all_values()
    headers = data.pop(1)
    return pd.DataFrame(data, columns=headers)
data = export_to_sheets("Copy of WorkSheet - Verifikasi Reg/UnReg Shipper Mitra","Main WorkSheet")
data = data[1:]

current_date = datetime.now().strftime("%d-%b-%Y")

data = data[data.iloc[:, 29] == current_date]
print(data.info())

this_is_list_cc = 'andi.darmawan@ninjavan.co'
this_is_list_bcc = 'ahmad.kamal@ninjavan.co'

    
def send_mail(recipient,cc,bcc,sender,app_password,subject,mitra_name,
shipper_name,global_id,sample_trid,onboard_date):

    cc1 = [cc]
    bcc1 = [bcc]
    list_cc = [elem.strip().split(',') for elem in cc1]
    list_bcc =  [elem.strip().split(',') for elem in bcc1]
    recipients = [recipient]+list_cc[0]+list_bcc[0]
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['To'] = recipient
    msg['From'] = sender
    msg['Cc'] = cc
    msg['Bcc'] = bcc
    
    
    html = """\
    <html>
    <head>
    
    </head>
    <body>
    <div style = "display: grid;
      grid-template-columns: auto;
      gap: 0px;
      background-color: white;
      padding-left: 250px;padding-right: 250px;">
      
      <div style="  background-color: #CEF6CE;
      border: 1px solid black;
      justify-content: left;
      text-align: left;
      font-size: 30px;">
      
      <img src="https://drive.google.com/thumbnail?id=1FTRA3zJklF25IKOZ4PNwHQaR5dq82mwX" alt="Image Description" style="margin: 10px;display: block;max-width: 160px"></img>
      
      </div>
      
      
      
    
      <div style="  background-color:white;
      border: 1px solid black;
      text-align: center;
      padding: 10px;">
      
    
      <b style="text-align: center;font-size:120%;">Hi {0},</b>
      
      <img src="https://drive.google.com/thumbnail?id=1ht91nsS3IiHff8QSecmO9y6UY302Y9oF" alt="Image Description" style="margin: 10px auto;display: block;max-width: 90px"></img>
     
      <p style="text-align: left;font-size:120%;">Shipper yang kamu daftarkan sudah masuk ke dalam reservasi di Ninja Driver kamu ya, berikut detailnya:</p>
      
            <div style="padding-left: 20px; padding-right: 20px;">
                <table style="width: 100%; border-collapse: collapse; background-color: white;">
                    <tr>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px; width: 30%;">
                            <b>Nama Shipper</b>
                        </td>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px; width: 70%;">
                            <b>{1}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Global Shipper ID</b>
                        </td>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{2}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Sample Tracking ID</b>
                        </td>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{3}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Tanggal Pick Up</b>
                        </td>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{4}</b>
                        </td>
                    </tr>
                </table>
            </div>


      <p style="text-align: left;font-size:120%;">Kamu wajib Pick Up paket shipper tersebut sesuai dengan reservasi pada aplikasi Ninja Driver kamu ya!</p>
      

      <br>
      <p style="text-align: left;font-size:120%;">Terima Kasih!</p>
      <p style="text-align: left;font-size:120%;">Ninja Xpress</p>
      </div>
    </div>
    
    
    </body>
    </html>
    """.format(mitra_name,shipper_name,global_id,sample_trid,onboard_date)
    
    part1 = MIMEText(html, 'html')
    msg.attach(part1)
    
    
    
    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo() 
        server.starttls(context=context)
        server.ehlo() 
        server.login(sender, app_password)
        server.sendmail(sender,recipients, msg.as_string()) 
    
    status = "Email "+ recipient +" Sent..."
    print(status) 
    return status

def send_mail1(recipient,cc,bcc,sender,app_password,subject,mitra_name,
shipper_name,global_id,sample_trid,alasan):
    

    cc1 = [cc]
    bcc1 = [bcc]
    list_cc = [elem.strip().split(',') for elem in cc1]
    list_bcc =  [elem.strip().split(',') for elem in bcc1]
    print(recipient)
    print(list_cc[0])
    print(list_bcc[0])
    recipients = [recipient]+list_cc[0]+list_bcc[0]
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['To'] = recipient
    msg['From'] = sender
    msg['Cc'] = cc
    msg['Bcc'] = bcc
    
    
    html = """\
    <html>
    <head>
    
    </head>
    <body>
    <div style = "display: grid;
      grid-template-columns: auto;
      gap: 0px;
      background-color: white;
      padding-left: 250px;padding-right: 250px;">
      
      <div style="  background-color: #F5A9A9;
      border: 1px solid black;
      justify-content: left;
      text-align: left;
      font-size: 30px;">
      
      <img src="https://drive.google.com/thumbnail?id=1FTRA3zJklF25IKOZ4PNwHQaR5dq82mwX" alt="Image Description" style="margin: 10px;display: block;max-width: 160px"></img>
      
      </div>
      
      
      
    
      <div style="  background-color:white;
      border: 1px solid black;
      text-align: center;
      padding: 10px;">
      
    
      <b style="text-align: center;font-size:120%;">Hi {0},</b>
      
      <img src="https://drive.google.com/thumbnail?id=10fZGONb89qjeGvT3BjaHru-YYqpGKLaC" alt="Image Description" style="margin: 10px auto;display: block;max-width: 90px"></img>
     
      <p style="text-align: left;font-size:120%;">Mohon maaf shipper yang kamu daftarkan tidak dapat kami proses untuk dimasukkan ke dalam Ninja Driver kamu berikut detailnya:</p>
      
            <div style="padding-left: 20px; padding-right: 20px;">
                <table style="width: 100%; border-collapse: collapse; background-color: white;">
                    <tr>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px; width: 30%;">
                            <b>Nama Shipper</b>
                        </td>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px; width: 70%;">
                            <b>{1}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Global Shipper ID</b>
                        </td>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{2}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Sample Tracking ID</b>
                        </td>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{3}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Alasan</b>
                        </td>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{4}</b>
                        </td>
                    </tr>
                </table>
            </div>



      

      <br>
      <p style="text-align: left;font-size:120%;">Terima Kasih!</p>
      <p style="text-align: left;font-size:120%;">Ninja Xpress</p>
      </div>
    </div>
    
    
    </body>
    </html>
    """.format(mitra_name,shipper_name,global_id,sample_trid,alasan)
    
    part1 = MIMEText(html, 'html')
    msg.attach(part1)
    
    
    
    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo() 
        server.starttls(context=context)
        server.ehlo() 
        server.login(sender, app_password)
        server.sendmail(sender,recipients, msg.as_string()) 
    
    status = "Email "+ recipient +" Sent..."
    print(status) 
    return status

def send_mail2(recipient,cc,bcc,sender,app_password,subject,mitra_name,
shipper_name,global_id,sample_trid,onboard_date):
    

    cc1 = [cc]
    bcc1 = [bcc]
    list_cc = [elem.strip().split(',') for elem in cc1]
    list_bcc =  [elem.strip().split(',') for elem in bcc1]
    recipients = [recipient]+list_cc[0]+list_bcc[0]
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['To'] = recipient
    msg['From'] = sender
    msg['Cc'] = cc
    msg['Bcc'] = bcc
    
    
    html = """\
    <html>
    <head>
    
    </head>
    <body>
    <div style = "display: grid;
      grid-template-columns: auto;
      gap: 0px;
      background-color: white;
      padding-left: 250px;padding-right: 250px;">
      
      <div style="  background-color: #CEF6CE;
      border: 1px solid black;
      justify-content: left;
      text-align: left;
      font-size: 30px;">
      
      <img src="https://drive.google.com/thumbnail?id=1FTRA3zJklF25IKOZ4PNwHQaR5dq82mwX" alt="Image Description" style="margin: 10px;display: block;max-width: 160px"></img>
      
      </div>
      
      
      
    
      <div style="  background-color:white;
      border: 1px solid black;
      text-align: center;
      padding: 10px;">
      
    
      <b style="text-align: center;font-size:120%;">Hi {0},</b>
      
      <img src="https://drive.google.com/thumbnail?id=1ht91nsS3IiHff8QSecmO9y6UY302Y9oF" alt="Image Description" style="margin: 10px auto;display: block;max-width: 90px"></img>
     
      <p style="text-align: left;font-size:120%;">Kami ingin menginformasikan bahwa:</p>
      
            <div style="padding-left: 20px; padding-right: 20px;">
                <table style="width: 100%; border-collapse: collapse; background-color: white;">
                    <tr>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px; width: 30%;">
                            <b>Nama Shipper</b>
                        </td>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px; width: 70%;">
                            <b>{1}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Global Shipper ID</b>
                        </td>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{2}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>RSVN ID</b>
                        </td>
                        <td style="background-color: #CEF6CE; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{3}</b>
                        </td>
                    </tr>
                </table>
            </div>


      <p style="text-align: left;font-size:120%;">Shipper tersebut akan di take out dari reservasi kamu per tanggal {4} ya!</p>
      

      <br>
      <p style="text-align: left;font-size:120%;">Terima Kasih!</p>
      <p style="text-align: left;font-size:120%;">Ninja Xpress</p>
      </div>
    </div>
    
    
    </body>
    </html>
    """.format(mitra_name,shipper_name,global_id,sample_trid,onboard_date)
    

    
    part1 = MIMEText(html, 'html')
    msg.attach(part1)
    
    
    
    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo() 
        server.starttls(context=context)
        server.ehlo() 
        server.login(sender, app_password)
        server.sendmail(sender, recipients, msg.as_string()) 
    
    status = "Email "+ recipient +" Sent..."
    print(status) 
    return status

def send_mail3(recipient,cc,bcc,sender,app_password,subject,mitra_name,
shipper_name,global_id,sample_trid,alasan):
    

    cc1 = [cc]
    bcc1 = [bcc]
    list_cc = [elem.strip().split(',') for elem in cc1]
    list_bcc =  [elem.strip().split(',') for elem in bcc1]
    recipients = [recipient]+list_cc[0]+list_bcc[0]
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['To'] = recipient
    msg['From'] = sender
    msg['Cc'] = cc
    msg['Bcc'] = bcc
    
    
    html = """\
    <html>
    <head>
    
    </head>
    <body>
    <div style = "display: grid;
      grid-template-columns: auto;
      gap: 0px;
      background-color: white;
      padding-left: 250px;padding-right: 250px;">
      
      <div style="  background-color: #F5A9A9;
      border: 1px solid black;
      justify-content: left;
      text-align: left;
      font-size: 30px;">
      
      <img src="https://drive.google.com/thumbnail?id=1FTRA3zJklF25IKOZ4PNwHQaR5dq82mwX" alt="Image Description" style="margin: 10px;display: block;max-width: 160px"></img>
      
      </div>
      
      
      
    
      <div style="  background-color:white;
      border: 1px solid black;
      text-align: center;
      padding: 10px;">
      
    
      <b style="text-align: center;font-size:120%;">Hi {0},</b>
      
      <img src="https://drive.google.com/thumbnail?id=10fZGONb89qjeGvT3BjaHru-YYqpGKLaC" alt="Image Description" style="margin: 10px auto;display: block;max-width: 90px"></img>
     
      <p style="text-align: left;font-size:120%;">Kami ingin menginformasikan bahwa:</p>
      
            <div style="padding-left: 20px; padding-right: 20px;">
                <table style="width: 100%; border-collapse: collapse; background-color: white;">
                    <tr>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px; width: 30%;">
                            <b>Nama Shipper</b>
                        </td>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px; width: 70%;">
                            <b>{1}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Global Shipper ID</b>
                        </td>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{2}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>RSVN ID</b>
                        </td>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{3}</b>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>Alasan</b>
                        </td>
                        <td style="background-color: #F5A9A9; border: 1px solid black; text-align: left; height: 50px; font-size: 120%; padding-left: 10px;">
                            <b>{4}</b>
                        </td>
                    </tr>
                </table>
            </div>


<p style="text-align: left;font-size:120%;">Kamu tetap wajib Pick Up paket shipper tersebut sesuai dengan reservasi pada aplikasi Ninja Driver kamu ya!</p>

      

      <br>
      <p style="text-align: left;font-size:120%;">Terima Kasih!</p>
      <p style="text-align: left;font-size:120%;">Ninja Xpress</p>
      </div>
    </div>
    
    
    </body>
    </html>
    """.format(mitra_name,shipper_name,global_id,sample_trid,alasan)
    
 
    
    part1 = MIMEText(html, 'html')
    msg.attach(part1)
    
    
    
    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo() 
        server.starttls(context=context)
        server.ehlo() 
        server.login(sender, app_password)
        server.sendmail(sender,recipients, msg.as_string()) 
    
    status = "Email "+ recipient +" Sent..."
    print(status) 
    return status

for index,row in data.iterrows():
    print(row[5],row[26])
    if (row[5] == "REG") & (row[26]=="APPROVED"):
        send_mail(row[1],this_is_list_cc,this_is_list_bcc,'your email','your apppassword',"Notifikasi: Shipper "+row[8]+" Berhasil Terdaftar",row[15],row[8],row[7],row[3],row[31])
    elif (row[5] == "REG") & (row[26]=="REJECT"):
        send_mail1(row[1],this_is_list_cc,this_is_list_bcc,'your email','your apppassword',"Notifikasi: Shipper "+row[8]+" Tidak dapat Diproses",row[15],row[8],row[7],row[3],row[27])
    elif (row[5] == "UNREG") & (row[26]=="APPROVED"):
        send_mail2(row[1],this_is_list_cc,this_is_list_bcc,'your email','your apppassword',"Notifikasi: Shipper "+row[8]+" Sudah Berhasil di Take Out",row[15],row[8],row[7],row[3],row[31])
    elif (row[5] == "UNREG") & (row[26]=="REJECT"):
        send_mail3(row[1],this_is_list_cc,this_is_list_bcc,'your email','your apppassword',"Notifikasi: Shipper "+row[8]+" Tidak Berhasil di Take Out",row[15],row[8],row[7],row[3],row[27])
    time.sleep(1)
        
        
print("Done")
   
    
