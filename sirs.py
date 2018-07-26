from decimal import *
import requests
import json
json.encoder.FLOAT_REPR = lambda x: format(x, '.8f')

body = {'ProductName': 'DCT14-125LI',
        'BendLoss12cm1070nm': 0.016,
        'MFD1070nm': 14.6,
        'CoatingDiameter': 252.0,
        'LotNumber': '180119421-10',
        'Comments': 'Data entered by SOMAPP1 server',
        'XPositionOE': 4e-6,
        'FiberLength': 993.0,
        'XPositionIE': 6409.0,
        'PartsNumber': 80845.0,
        'HOMBendLoss24cm1070nm': 120.0,
        'CoreAttenuation1150nm': 1.9,
        'SerialNumber': '18211062265351',
        'Aeff1070nm': 177.9
        }

for attr, value in body.iteritems():
    print type(value)
    if type(value) is not str:
        print value
        body[attr] = "{:.8f}".format(value)

sirs = {'Header':
        {'Action': 'UploadData ',
         'TableName': 'fibers_fiber_details'},
        'Body': body}


print(json.dumps(sirs))

sirs_url = 'http://devwebapps.ofsoptics.com/corp/somerset/SIRS-PTS-Lite-WebServices/TransferCommands/api/DataTransfer/'
r = requests.post(sirs_url, json=sirs)
resp = r.json()
print(resp)
