import pandas as pd

df = pd.read_excel("PE14-125-2.2LI_18221062265569.xls")

sirs = {"Header" : {
 "Action" : "UploadData ",
"TableName" : "fibers_fiber_details"},
"Body" : {
    "Comments": "Data entered by SOMAPP1 server"
}
}

col_lookup = {}
col_lookup["Part Number"] = "PartsNumber"
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

for col in df:
    try:
        sirs["Body"][col_lookup[col]] = df[col][0]
    except:
        pass
print(sirs)