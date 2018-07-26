def load_fiber_data_sources(fiber_data_sources):
    fiber_data = []

    for fiber_data_source in fiber_data_sources:
        for data in fiber_data_source.load_data():
            fiber_data.append(data)
    
    return fiber_data


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

fiber_data_sources = [TestFiberDataSource()]

fiber_data = load_fiber_data_sources(fiber_data_sources)
print(fiber_data)
