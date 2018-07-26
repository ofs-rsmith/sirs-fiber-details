import email
import os
import shutil
import time
import pandas as pd
import numpy as np
import requests
import traceback
from imapclient import IMAPClient

class TestFiberDataSource():
    def __init__(self):
        self.fiber_data = []

    def load_data(self):
        data = {}
        data["ProductName"] = "PE14-125-2.2"
        data["PartNumber"] = "80853"
        data["Description"] = "PE14-125-2.2"
        data["SerialNumber"] = "18221062265569"
        data["LotNumber"] = "171212422-02"
        data["FiberLength"] = 499
        data["XPositionOE"] = 9315
        data["XPositionIE"] = 8816
        data["Aeff1070nm"] = 171.3
        data["MFD1070nm"] = 14.9
        data["MFD1070nmTapered"] = 14.2
        data["HOMBendLoss24cm1070nm"] = 190
        data["BendLoss12cm1070nm"] = 0
        data["CoatingDiameter"] = 255

        self.fiber_data = [data]

        return self.fiber_data

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