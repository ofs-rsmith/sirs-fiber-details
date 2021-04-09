import email
import os
import shutil
import time
import pandas as pd
import numpy as np
import requests
import traceback
from imapclient import IMAPClient
import pytz
#import datetime
import time
from datetime import datetime, timedelta, tzinfo

def store_attachment(part):
    """ Store attached files as they are and with the same name. """
    this_dir = os.path.dirname(os.path.realpath(__file__))
    TARGET = os.path.join(this_dir, 'attachments')
    
    filename = part.get_filename()
    att_path = os.path.join(TARGET, filename)
 
    if not os.path.isfile(att_path) :
        fp = open(att_path, 'wb')
        fp.write(part.get_payload(decode=True))
        fp.close()
    print("Successfully stored attachment!")

def check_attachment(mail):
    """ Examine if the email has the requested subject and store the attachment if so. """
    #print(mail)
    this_dir = os.path.dirname(os.path.realpath(__file__))
    attachments_dir = os.path.join(this_dir, 'attachments')
    
    print("attach: ["+mail["From"]+"] :" + mail["Subject"] + " : date at: " + mail["Date"])
    mail_date = ""
    do_parse = False
    for part in mail.walk():
        #print("content = " + part.get_content_maintype())
        #print("Content-Disposition = " + part.get('Content-Disposition'))
        if part.get_content_maintype() == 'multipart':
            if do_parse:
               do_parse = False 
               for filename in os.listdir(attachments_dir):
                    try:
                        filepath = os.path.join(attachments_dir, filename)
                        filename_no_ext, file_ext = os.path.splitext(filepath)
                        if file_ext == ".xls":
                            try:
                                if mail_date == "" :
                                    date_s = mail["Date"]
                                else:
                                    date_s = mail_date
                                print("exe write to sirs ***" + date_s)
                                upload_file_to_sirs(filepath, date_s)
                            except:
                                traceback.print_exc()
                                pass
                        else:
                            #DELETE NON-EXCEL FILES
                            os.unlink(filepath) 
                            print("***NON-EXCEL FILE: " + filepath)                             
                    except:
                        traceback.print_exc()
                        pass
            lines = str(part.get_payload(1))
            s = lines.find('creation-date=')
            l = len("creation-date=")
            if s > 0 :
                e = lines[(s+l+1):].find(";")
                #print ("string1 = " + lines[(s+l+1):])
                mail_date = ((lines[(s+l+1):])[:(e-1)]).replace('GMT', '+0000')

        #if part.get('Content-Disposition') is None:
        #    continue
        if part.get_content_maintype() == 'application':
            store_attachment(part)
            time.sleep(3)
            #print("Successfully stored attachment!")
            do_parse = True
    if do_parse:
        do_parse = False 
        for filename in os.listdir(attachments_dir):
            try:
                filepath = os.path.join(attachments_dir, filename)
                filename_no_ext, file_ext = os.path.splitext(filepath)
                if file_ext == ".xls":
                    try:
                        if mail_date == "" :
                            date_s = mail["Date"]
                        else:
                            date_s = mail_date
                        print("exe write to sirs ***" + date_s)
                        upload_file_to_sirs(filepath, date_s)
                    except:
                        traceback.print_exc()
                        pass
                else:
                    #DELETE NON-EXCEL FILES
                    os.unlink(filepath)  
                    print("***NON-EXCEL FILE: " + filepath)
            except:
                traceback.print_exc()
                pass
def main(data):
    # run with run_test.py
    #email = 'somfiberlabelgen@ofsoptics.com'
    HOST = 'mail.ofsoptics.com'
    USERNAME = 'somfiberlabelgen'
    PASSWORD = 'S0merset!'
    print("...")
    
    server = IMAPClient(HOST, ssl=False)
    print("Logging In to Mail Server...")
    server.login(USERNAME, PASSWORD)
    print("Selecting Inbox...")
    server.select_folder('INBOX')
    print("Searching...")
    messages = server.search('UNSEEN')
    print(messages)
    all_unread = server.fetch(messages, ['RFC822'])
    for msg_id in all_unread:
        try:
            new_email = email.message_from_string(all_unread[msg_id][b'RFC822'].decode("utf-8"))
            print("["+new_email["From"]+"] :" + new_email["Subject"])
            print("date= "+ new_email["Date"])
            check_attachment(new_email)
        except:
            traceback.print_exc()
            pass

    server.logout()
    print("Logging Out of Mail Server...")
    
def upload_file_to_sirs(filepath, maildate):
    print("Reading Excel file: {}".format(filepath))
    print ("datetime = " + maildate)
    df = pd.read_excel(filepath).iloc[0].dropna()

    sirs = {"Header" : {
    "Action" : "UploadData ",
    #"TableName" : "fibers_fiber_details"},
    "TableName" : "fibers_dk_fiber_details"},
    "Body" : {
        "Comments": "Data entered by SOMAPP1 server"
    }
    }
    col_lookup = {}
    col_lookup["Product Name"] = "ProductName"
    col_lookup["Part Number"] = "PartNumber"
    col_lookup["Serial Number"] = "SerialNumber"
    col_lookup["Lot number"] = "LotNumber"
    col_lookup["Fiber Length (m)"] = "FiberLength"
    col_lookup["X-pos OE (m)"] = "XPositionOSE"   
    col_lookup["X-pos IE (m)"] = "XPositionISE"     
    col_lookup["Aeff 1070 nm (um)^2"] = "Aeff1070nm"    
    col_lookup["MFD 1070 nm (um)"] = "MFD"    # for different wavelength
    col_lookup["Cladding abs. 920 nm (dB/m)"] = "CladdingAbsorption"    # for different wavelength
    col_lookup["Core attenuation 1150 nm (dB/km)"] = "CoreAttenuation1150nm"
    col_lookup["MFD 1070 nm Tapered (um)"] = "MFD1070nmTapered"
    col_lookup["HOM bend loss 24 cm 1070 nm (dB/m)"] = "HOMBendLoss24cm1070nm"
    col_lookup["Bend loss 12 cm 1070 nm (dB/m)"] = "BendLoss12cm1070nm"
    col_lookup["Cladding diameter flat-flat (um)"] = "CladdingDiameterFlatFlat"
    col_lookup["Cladding diameter peak-peak (um)"] = "CladdingDiameterPeakPeak"
    col_lookup["Coating diameter (um)"] = "CoatingDiameter"
    col_lookup["Core absorption@1535 nm (dB/m)"] = "CoreAbsorption1535nm" # new
    col_lookup["CoreDiameter no record"] = "CoreDiameter"  # new
    col_lookup["Core NA"] = "CoreNA"   # new
    col_lookup["Aeff 1070 nm Tapered (um)^2"] = "AeffTapered"  # new
    col_lookup["Aeff 1070 nm Tapered"] = "AeffTapered" # handle alternative

    types = {}
    types["ProductName"] = "VARCHAR"
    types["PartNumber"] = "VARCHAR"
    #types["Description"] = "VARCHAR"    
    types["SerialNumber"] = "VARCHAR"
    types["LotNumber"] = "VARCHAR"
    types["FiberLength"] = "DECIMAL"
    types["XPositionOSE"] = "DECIMAL"
    types["XPositionISE"] = "DECIMAL"
    types["Aeff1070nm"] = "DECIMAL"
    types["MFD"] = "DECIMAL"
    types["CladdingAbsorption"] = "DECIMAL"
    types["CoreAttenuation1150nm"] = "DECIMAL"
    types["MFD1070nmTapered"] = "DECIMAL"
    types["HOMBendLoss24cm1070nm"] = "DECIMAL"
    types["BendLoss12cm1070nm"] = "DECIMAL"
    types["CladdingDiameterFlatFlat"] = "DECIMAL"
    types["CladdingDiameterPeakPeak"] = "DECIMAL"
    types["CoatingDiameter"] = "DECIMAL"
    types["CoreAbsorption1535nm"] = "DECIMAL"
    types["CoreDiameter"] = "DECIMAL"
    types["CoreNA"] = "DECIMAL"
    types["AeffTapered"] = "DECIMAL"
    types["Result"] = "VARCHAR"      
    types["Comments"] = "VARCHAR"
    #types["ReceivedDateTime"] = "DATETIME2"
    
    for key, val in df.iteritems():
        try:
            key = key.strip()
            if (types[col_lookup[key]] == "DECIMAL"):
                sirs["Body"][col_lookup[key]] = "{:.7f}".format(float(val))
            else:
                sirs["Body"][col_lookup[key]] = val
        except:
            pass
    tz = maildate[-5:]
    
    date_str = maildate[:-6]
    dt_utc = datetime.strptime(date_str, "%a, %d %b %Y %H:%M:%S")
    
    utc_est = -500    
    if bool(pytz.timezone('US/Eastern').dst(dt_utc, is_dst=None)):
        utc_est = -500 + 100
    delta = utc_est - int(tz)  #local timezone GMT/UTC = -0500
    
    dt_utc = dt_utc + timedelta(hours = int(delta/100)) 
    dt = dt_utc.replace(tzinfo=FixedOffset(tz))

    #dateparsed = datetime.datetime.strptime(maildate, '%a, %d %b %Y %H:%M:%S %z')
    datesaved = dt.strftime("%m/%d/%Y %H:%M:%S")    #datetime.utcnow().strftime("%m/%d/%Y %H:%M:%S %p")
    
    print('datesaved = {}'.format(datesaved))
    sirs['Body']["ReceivedDateTime"] = datesaved
    
    print('sirs DB = {}'.format(sirs))
    sirs_url = 'http://webapps.ofsoptics.com/corp/somerset/SIRS-PTS-Lite-WebServices/TransferCommands/api/DataTransfer/'
    #sirs_url = 'http://devwebapps.ofsoptics.com/corp/somerset/SIRS-PTS-Lite-WebServices/TransferCommands/api/DataTransfer/'
    ###post data
    r = requests.post(sirs_url, json=sirs)
    resp = r.json()
    ##print('requests post = {}'.format(resp))
    if resp['Response']['Message'] == "OK":
        # IF SIRS UPLOAD WAS SUCCESSFUL, DELETE THE FILE
        os.unlink(filepath)

class FixedOffset(tzinfo):
    """offset_str: Fixed offset in str: e.g. '-0400'"""
    def __init__(self, offset_str):
        sign, hours, minutes = offset_str[0], offset_str[1:3], offset_str[3:]
        offset = (int(hours) * 60 + int(minutes)) * (-1 if sign == "-" else 1)
        self.__offset = timedelta(minutes=offset)
        # NOTE: the last part is to remind about deprecated POSIX GMT+h timezones
        # that have the opposite sign in the name;
        # the corresponding numeric value is not used e.g., no minutes
        '<%+03d%02d>%+d' % (int(hours), int(minutes), int(hours)*-1)
    def utcoffset(self, dt=None):
        return self.__offset
    def tzname(self, dt=None):
        return self.__name
    def dst(self, dt=None):
        return timedelta(0)
    def __repr__(self):
        return 'FixedOffset(%d)' % (self.utcoffset().total_seconds() / 60)    
