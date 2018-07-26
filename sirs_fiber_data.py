import email
import os
import shutil
import time
import pandas as pd
import numpy as np
import requests
import traceback
from imapclient import IMAPClient

def store_attachment(part):
    """ Store attached files as they are and with the same name. """
    TARGET = '/mnt/c/Users/robsmith/Documents/Projects/SIRS Fiber Data/attachments'  # this is the path where attachments are stored 

    filename = part.get_filename()
    att_path = os.path.join(TARGET, filename)
 
    if not os.path.isfile(att_path) :
        fp = open(att_path, 'wb')
        fp.write(part.get_payload(decode=True))
        fp.close()
    print("Successfully stored attachment!")
 

def check_attachment(mail):
    """ Examine if the email has the requested subject and store the attachment if so. """
    # print(mail)

    print("["+mail["From"]+"] :" + mail["Subject"])
 
    for part in mail.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
 
        store_attachment(part)
        time.sleep(3)


def upload_file_to_sirs(filepath):
    print("Reading Excel file: {}".format(filepath))
    df = pd.read_excel(filepath).iloc[0].dropna()

    sirs = {"Header" : {
    "Action" : "UploadData ",
    "TableName" : "fibers_fiber_details"},
    "Body" : {
        "Comments": "Data entered by SOMAPP1 server"
    }
    }

    col_lookup = {}
    col_lookup["Part Number"] = "PartNumber"
    col_lookup["Product Name"] = "ProductName"
    col_lookup["Serial Number"] = "SerialNumber"
    col_lookup["Lot number"] = "LotNumber"
    col_lookup["Fiber Length (m)"] = "FiberLength"
    col_lookup["X-pos OE (m)"] = "XPositionOE"
    col_lookup["X-pos IE (m)"] = "XPositionIE"
    col_lookup["Aeff 1070 nm (um)^2"] = "Aeff1070nm"
    col_lookup["MFD 1070 nm (um)"] = "MFD1070nm"
    col_lookup["Cladding abs. 920 nm (dB/m)"] = "CladdingAbs920nm"
    col_lookup["Core attenuation 1150 nm (dB/km)"] = "CoreAttenuation1150nm"
    col_lookup["MFD 1070 nm Tapered (um)"] = "Mfd1070nmTapered"
    col_lookup["HOM bend loss 24 cm 1070 nm (dB/m)"] = "HOMBendLoss24cm1070nm"
    col_lookup[" Bend loss 12 cm 1070 nm (dB/m)"] = "BendLoss12cm1070nm"
    col_lookup["Cladding diameter flat-flat (um)"] = "CladdingDiameterFlatFlat"
    col_lookup["Cladding diameter peak-peak (um)"] = "CladdingDiameterPeakPeak"
    col_lookup["Coating diameter (um)"] = "CoatingDiameter"

    types = {}
    types["ProductName"] = "VARCHAR"
    types["PartNumber"] = "VARCHAR"
    types["Description"] = "VARCHAR"
    types["SerialNumber"] = "VARCHAR"
    types["LotNumber"] = "VARCHAR"
    types["FiberLength"] = "DECIMAL"
    types["XPositionOE"] = "DECIMAL"
    types["XPositionIE"] = "DECIMAL"
    types["Aeff1070nm"] = "DECIMAL"
    types["MFD1070nm"] = "DECIMAL"
    types["CladdingAbs920nm"] = "DECIMAL"
    types["CoreAttenuation1150nm"] = "DECIMAL"
    types["MFD1070nmTapered"] = "DECIMAL"
    types["HOMBendLoss24cm1070nm"] = "DECIMAL"
    types["BendLoss12cm1070nm"] = "DECIMAL"
    types["CladdingDiameterFlatFlat"] = "DECIMAL"
    types["CladdingDiameterPeakPeak"] = "DECIMAL"
    types["CoatingDiameter"] = "DECIMAL"
    types["Result"] = "VARCHAR"
    types["Comments"] = "VARCHAR"

    for key, val in df.iteritems():
        try:
            if (types[col_lookup[key]] == "DECIMAL"):
                sirs["Body"][col_lookup[key]] = "{:.7f}".format(float(val))
            else:
                sirs["Body"][col_lookup[key]] = val
        except:
            pass

    print(sirs)
    sirs_url = 'http://devwebapps.ofsoptics.com/corp/somerset/SIRS-PTS-Lite-WebServices/TransferCommands/api/DataTransfer/'
    r = requests.post(sirs_url, json=sirs)
    resp = r.json()
    print(resp)
    if resp['Response']['Message'] == "OK":
        # IF SIRS UPLOAD WAS SUCCESSFUL, DELETE THE FILE
        os.unlink(filepath)


HOST = 'mail.ofsoptics.com'
USERNAME = 'sirsfiberdata'
PASSWORD = 'somersetsirs'

# Get attachments folder path
this_dir = os.path.dirname(os.path.realpath(__file__))
attachments_dir = os.path.join(this_dir, 'attachments')

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
        check_attachment(new_email)
    except:
        traceback.print_exc()
        pass

print("Logging Out of Mail Server...")
server.logout()

for filename in os.listdir(attachments_dir):
    try:
        filepath = os.path.join(attachments_dir, filename)
        filename_no_ext, file_ext = os.path.splitext(filepath)
        
        if file_ext == ".xls":
            try:
                upload_file_to_sirs(filepath)
            except:
                traceback.print_exc()
                pass
        else:
            # DELETE NON-EXCEL FILES
            os.unlink(filepath)    
    except:
        traceback.print_exc()
        pass
