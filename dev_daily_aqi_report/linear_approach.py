from datetime import datetime, timedelta
import xlsxwriter, os, time
import requests
import logging


GIS_USERNAME = "Developer1"
GIS_PASSWORD = "qweRTY77**"
AERIS_ID = 'IJwChrr7utrp6BFWVnJ1A'
AERIS_SECRET = 'WmzZHc1PnJiAsr996tFNwfsPt7IqetRHHuB8Ldty'

def get_each_client(client_ids):
    for each_id in client_ids:  
        # client_json_data = aqi_clients(each_id)
        cid = 369
        name = "Home Depot Canada"
        aqithreshold = 150
        # states = client_json_data[0].get('aqistates')
        aqistates = 'BC'
        redflag = "F"
        fireweather = "F"
        # aqi_email = client_json_data[0].get('aqi_email')
        email_list = ['369_PrimaryDistribution@earlyalert.com', '369_West_Region@earlyalert.com']

        print(cid, name, aqithreshold, "states", aqistates, redflag, fireweather, "aqi_email", email_list)
        create_product1(cid, name, aqithreshold, aqistates, redflag, fireweather, email_list)

def create_product1(clientId, name, threshold, aqistates, redflag, fireweather, email_list):
    """Creates excel worksheet based on product data"""
    formats, workbook, workbook_filename = formats_workbook(name)
    # gis = GIS("https://maps.earlyalert.com/portal/home/", GIS_USERNAME, GIS_PASSWORD)
    # item = gis.content.get("3604f01a15274e139e47ba1fe00183aa")
    # feature_layer = item.layers[0]
    # query = 'ClientId = {} AND State IN ({})'.format(clientId, aqistates)
    # f = feature_layer.query(where=query)
    # print(query)

    f = [{"geometry": {"x": -13294370.3632, "y": 6427068.618199997, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696047, "clientid": 369, "name": "7032", "code": "CA 7032", "code2": None, "address": "2515 ENTERPRISE WAY", "city": "KELOWNA", "state": "BC", "country": "CA", "zipcode": "V1X 7K2", "lat": 49.88896456, "lon": -119.4253609, "phone": None, "districtcode": "342", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760166477, "creator": "EA_ClientServices", "editdate": 1681760166477, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13707999.928, "y": 6330238.202100001, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696050, "clientid": 369, "name": "7035", "code": "CA 7035", "code2": None, "address": "E1-840 MAIN STREET", "city": "WEST VANCOUVER", "state": "BC", "country": "CA", "zipcode": "V7T 2Z3", "lat": 49.32529335, "lon": -123.1410585, "phone": None, "districtcode": "342", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760166636, "creator": "EA_ClientServices", "editdate": 1681760166636, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13809146.798799999, "y": 6314337.4758, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696054, "clientid": 369, "name": "7040", "code": "CA 7040", "code2": None, "address": "6555 METRAL STREET", "city": "NANAIMO", "state": "BC", "country": "CA", "zipcode": "V9T 2L9", "lat": 49.23210828, "lon": -124.0496763, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760166843, "creator": "EA_ClientServices", "editdate": 1681760166843, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13655248.827300001, "y": 6295611.759599999, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696055, "clientid": 369, "name": "7041", "code": "CA 7041", "code2": None, "address": "6550 - 200TH STREET", "city": "LANGLELY", "state": "BC", "country": "CA", "zipcode": "V2Y 1P2", "lat": 49.12214173, "lon": -122.6671873, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760166894, "creator": "EA_ClientServices", "editdate": 1681760166894, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13701606.3487, "y": 6320605.890500002, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696056, "clientid": 369, "name": "7042", "code": "CA 7042", "code2": None, "address": "900 TERMINAL AVENUE", "city": "VANCOUVER", "state": "BC", "country": "CA", "zipcode": "V6A 4G4", "lat": 49.2688649, "lon": -123.083624, "phone": None, "districtcode": "342", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760166953, "creator": "EA_ClientServices", "editdate": 1681760166953, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13701332.2578, "y": 6307910.238499999, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696057, "clientid": 369, "name": "7043", "code": "CA 7043", "code2": None, "address": "2700 SWEDEN WAY", "city": "RICHMOND", "state": "BC", "country": "CA", "zipcode": "V6V 1K1", "lat": 49.19439194, "lon": -123.0811618, "phone": None, "districtcode": "342", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760167006, "creator": "EA_ClientServices", "editdate": 1681760167006, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13679752.2621, "y": 6298024.504699998, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696058, "clientid": 369, "name": "7044", "code": "CA 7044", "code2": None, "address": "7350-120TH STREET", "city": "SURREY", "state": "BC", "country": "CA", "zipcode": "V3W 3M9", "lat": 49.13632426, "lon": -122.8873054, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760167061, "creator": "EA_ClientServices", "editdate": 1681760167061, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13674719.6304, "y": 6313500.028700002, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696059, "clientid": 369, "name": "7045", "code": "CA 7045", "code2": None, "address": "1900 UNITED BOULEVARD (STORE)", "city": "COQUITLAM", "state": "BC", "country": "CA", "zipcode": "V3K 6Z1", "lat": 49.2271956, "lon": -122.8420965, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760167117, "creator": "EA_ClientServices", "editdate": 1681760167117, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13677689.322700001, "y": 6309533.2809000015, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696060, "clientid": 369, "name": "7046", "code": "CA 7046", "code2": None, "address": "12701 - 110TH AVENUE", "city": "SURREY", "state": "BC", "country": "CA", "zipcode": "V3V 3K7", "lat": 49.203919, "lon": -122.8687737, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760167170, "creator": "EA_ClientServices", "editdate": 1681760167170, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13694188.9195, "y": 6319608.813900001, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696061, "clientid": 369, "name": "7047", "code": "CA 7047", "code2": None, "address": "3950 HENNING DRIVE", "city": "BURNABY", "state": "BC", "country": "CA", "zipcode": "V5C 6M2", "lat": 49.26302009, "lon": -123.0169921, "phone": None, "districtcode": "342", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760167222, "creator": "EA_ClientServices", "editdate": 1681760167222, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13709183.6438, "y": 6397845.178000003, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696065, "clientid": 369, "name": "7053", "code": "CA 7053", "code2": None, "address": "39251 DISCOVERY WAY", "city": "SQUAMISH", "state": "BC", "country": "CA", "zipcode": "V8B 0M9", "lat": 49.719535, "lon": -123.151692, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760167437, "creator": "EA_ClientServices", "editdate": 1681760167437, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13729474.760200001, "y": 6185513.322400004, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696066, "clientid": 369, "name": "7055", "code": "CA 7055", "code2": None, "address": "3980 SHELBOURNE STREET", "city": "SAANICH", "state": "BC", "country": "CA", "zipcode": "V8N 3E3", "lat": 48.47064319, "lon": -123.3339702, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760167488, "creator": "EA_ClientServices", "editdate": 1681760167488, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13655717.7829, "y": 6293548.275899999, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696078, "clientid": 369, "name": "7072", "code": "CA 7072", "code2": None, "address": "19930 Fraser Highway", "city": "Langley", "state": "BC", "country": "CA", "zipcode": None, "lat": 49.110009, "lon": -122.6714, "phone": None, "districtcode": None, "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760168158, "creator": "EA_ClientServices", "editdate": 1681760168158, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13747919.676199999, "y": 6184039.851599999, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696080, "clientid": 369, "name": "7074", "code": "CA 7074", "code2": None, "address": "2400 MILLSTREAM ROAD", "city": "LANGFORD", "state": "BC", "country": "CA", "zipcode": "V9B 3R3", "lat": 48.46186664, "lon": -123.4996637, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760168265, "creator": "EA_ClientServices", "editdate": 1681760168265, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13277842.8586, "y": 6494951.796499997, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696089, "clientid": 369, "name": "7084", "code": "CA 7084", "code2": None, "address": "5501 ANDERSON WAY", "city": "VERNON", "state": "BC", "country": "CA", "zipcode": "V1T 9V1", "lat": 50.28024648, "lon": -119.2768918, "phone": None, "districtcode": "342", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760168760, "creator": "EA_ClientServices", "editdate": 1681760168760, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13670222.7125, "y": 6284997.312600002, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696112, "clientid": 369, "name": "7122", "code": "CA 7122", "code2": None, "address": "2527 160TH STREET", "city": "SURREY", "state": "BC", "country": "CA", "zipcode": "V3Z 0C8", "lat": 49.0597, "lon": -122.8017, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760169974, "creator": "EA_ClientServices", "editdate": 1681760169974, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13611051.438299999, "y": 6280783.532899998, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696129, "clientid": 369, "name": "7141", "code": "CA 7141", "code2": None, "address": "1956 VEDDER WAY", "city": "ABBOTSFORD", "state": "BC", "country": "CA", "zipcode": "V2S 8K1", "lat": 49.03488977, "lon": -122.2701554, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760170932, "creator": "EA_ClientServices", "editdate": 1681760170932, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13399495.0913, "y": 6561863.266900003, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696131, "clientid": 369, "name": "7144", "code": "CA 7144", "code2": None, "address": "1020 HILLSIDE DRIVE", "city": "KAMLOOPS", "state": "BC", "country": "CA", "zipcode": "V2E 2S5", "lat": 50.66280546, "lon": -120.3697124, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760171036, "creator": "EA_ClientServices", "editdate": 1681760171036, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13663658.9703, "y": 6318716.826399997, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696132, "clientid": 369, "name": "7145", "code": "CA 7145", "code2": None, "address": "1069 NICOLA DRIVE", "city": "PORT COQUITLAM", "state": "BC", "country": "CA", "zipcode": "V3B 8B2", "lat": 49.25779072, "lon": -122.7427369, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760171101, "creator": "EA_ClientServices", "editdate": 1681760171101, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13667887.463399999, "y": 7144746.3116, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696155, "clientid": 369, "name": "7171", "code": "CA 7171", "code2": None, "address": "5959 O'GRADY ROAD", "city": "PRINCE GEORGE", "state": "BC", "country": "CA", "zipcode": "V2N 6Z5", "lat": 53.86561478, "lon": -122.7807221, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760172325, "creator": "EA_ClientServices", "editdate": 1681760172325, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13910990.195700001, "y": 6396597.983599998, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696160, "clientid": 369, "name": "7177", "code": "CA 7177", "code2": None, "address": "388 LERWICK ROAD", "city": "COURTENAY", "state": "BC", "country": "CA", "zipcode": "V9N 9E5", "lat": 49.71229091, "lon": -124.9645511, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760172586, "creator": "EA_ClientServices", "editdate": 1681760172586, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13942835.9079, "y": 6452294.0254999995, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696173, "clientid": 369, "name": "7221", "code": "CA 7221", "code2": None, "address": "1482 ISLAND HIGHWAY", "city": "CAMPBELL RIVER", "state": "BC", "country": "CA", "zipcode": "V9W 8C9", "lat": 50.034738, "lon": -125.250626, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760173268, "creator": "EA_ClientServices", "editdate": 1681760173268, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13316995.5042, "y": 6416941.945600003, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696196, "clientid": 369, "name": "7252", "code": "CA 7252", "code2": None, "address": "3350 CARRINGTON RD", "city": "WESTBANK", "state": "BC", "country": "CA", "zipcode": "V4T 2Z1", "lat": 49.83032, "lon": -119.628606, "phone": None, "districtcode": "342", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760174509, "creator": "EA_ClientServices", "editdate": 1681760174509, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -12885409.1705, "y": 6365914.269400001, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696199, "clientid": 369, "name": "7255", "code": "CA 7255", "code2": None, "address": "2000 MCPHEE CT", "city": "CRANBROOK", "state": "BC", "country": "CA", "zipcode": "V1C 0A3", "lat": 49.53373, "lon": -115.7516, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760174665, "creator": "EA_ClientServices", "editdate": 1681760174665, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13705074.9527, "y": 6319804.2874, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696203, "clientid": 369, "name": "7259", "code": "CA 7259", "code2": None, "address": "2388 Cambie Street", "city": "Vancouver", "state": "BC", "country": "CA", "zipcode": "V5Z 2T8", "lat": 49.264166, "lon": -123.114783, "phone": None, "districtcode": "342", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760174889, "creator": "EA_ClientServices", "editdate": 1681760174889, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13772323.5024, "y": 6241524.989799999, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696212, "clientid": 369, "name": "7272", "code": "CA 7272", "code2": None, "address": "1-2980 Drink Water Road", "city": "Duncan", "state": "BC", "country": "CA", "zipcode": "V9L 6C6", "lat": 48.803145, "lon": -123.718887, "phone": None, "districtcode": "280", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760175370, "creator": "EA_ClientServices", "editdate": 1681760175370, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13578493.671, "y": 6300755.734099999, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696213, "clientid": 369, "name": "7273", "code": "CA 7273", "code2": None, "address": "100-8443 Eagle Landing", "city": "Chilliwack", "state": "BC", "country": "CA", "zipcode": "V2P 0E2", "lat": 49.152374, "lon": -121.977684, "phone": None, "districtcode": "126", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760175423, "creator": "EA_ClientServices", "editdate": 1681760175423, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13707325.4988, "y": 6309887.844400004, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696216, "clientid": 369, "name": "7283", "code": "CA 7283", "code2": None, "address": "7003 72nd Avenue", "city": "Vancouver", "state": "BC", "country": "CA", "zipcode": None, "lat": 45.67318, "lon": -122.599345, "phone": None, "districtcode": None, "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760175579, "creator": "EA_ClientServices", "editdate": 1684848155420, "editor": "Developer1", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}, {"geometry": {"x": -13657657.636300001, "y": 6291952.622199997, "spatialReference": {"wkid": 102100, "latestWkid": 3857}}, "attributes": {"objectid": 3696220, "clientid": 369, "name": "7348", "code": "CA 7348", "code2": None, "address": "19238 54th Ave", "city": "Surrey", "state": "BC", "country": "CA", "zipcode": "V3S 8E5", "lat": 49.10062495, "lon": -122.688826, "phone": None, "districtcode": "Surrey BDC", "divisioncode": None, "regioncode": "West", "facilitycategory": None, "level1email": None, "level2email": None, "level3email": None, "level4email": None, "level5email": None, "level6email": None, "level7email": None, "level8email": None, "level10email": None, "level11email": None, "level12email": None, "level13email": None, "level14email": None, "level15email": None, "level16email": None, "level17email": None, "level18email": None, "email": None, "creationdate": 1681760175812, "creator": "EA_ClientServices", "editdate": 1681760175812, "editor": "EA_ClientServices", "indicator1": None, "indicator2": None, "level1phone": None, "level9email": None, "level2phone": None, "level3phone": None, "level4phone": None, "level5phone": None, "code3": None, "code4": None, "code5": None, "code6": None, "code7": None, "county": None, "clientname": "Home Depot Canada", "code8": None, "code9": None, "code10": None, "code11": None, "code12": None, "code13": None, "code14": None, "code15": None}}]
    print(f)
    worksheet1, workbook, workbook_sheet_name_today, max_row_today  = aqi_today(formats, workbook, clientId, name, threshold, aqistates, f)
    # time.sleep(30)
    # workbook.close()
    worksheet2, workbook, workbook_sheet_name_tomorrow, max_row_tomorrow  = aqi_tomorrow(formats, workbook, clientId, name, threshold, aqistates, f, workbook_filename)
    workbook.close()
    send_aqi_email(clientId, workbook.filename, email_list, workbook_filename)


def generate_esri_token():
    url = "https://www.arcgis.com/sharing/generateToken"
    params = {
        "f": "json",
        "username": 'EA_Developer1',
        "password": GIS_PASSWORD,
        "client": "referer",
        "referer": "https://maps.earlyalert.com/portal/home/"
    }
    r = requests.post(url, data=params)
    token = r.json()["token"]
    return token

def aqi_clients(client_id):
    aq_cid_list = []
    url = 'https://services8.arcgis.com/lrWk3ELQFeb23nh1/arcgis/rest/services/service_d8f8bb17c9274da9940c1daaeedb833a/FeatureServer/0/query?'
    token = generate_esri_token()
    params = {
        'where': f'aqithreshold is not None AND cid = {client_id}',
        'geometryType': 'esriGeometryEnvelope',
        'spatialRel': 'esriSpatialRelIntersects',
        'relationParam': '',
        'outFields': '*',
        'returnGeometry': 'true',
        'geometryPrecision': '',
        'outSR': '',
        'returnIdsOnly': 'false',
        'returnCountOnly': 'false',
        'orderByFields': '',
        'groupByFieldsForStatistics': '',
        'returnZ': 'false',
        'returnM': 'false',
        'returnDistinctValues': 'false',
        'f': 'pjson',
        'token': token
    }
    r = requests.get(url, params=params)
    output_data = r.json()
    length = len(output_data.get('features'))
    for i in range(length):
        data = output_data.get('features')[i].get('attributes')
        aq_cid_list.append(data)
    return aq_cid_list

def formats_workbook(name):
    workbook_filename = f"{name} Air Quality Forecast {datetime.now():%Y-%m-%d}.xlsx"
    workbook = xlsxwriter.Workbook(os.path.join("/tmp",
        workbook_filename))
    # formats
    
    title_format = workbook.add_format({
        "bold": True,
        # "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "black",
        "font_color": "white",
        "font_size": 14,
        "text_wrap": True,
        # "border_color": "white",
    })
    date_format = workbook.add_format({
        "bold": True,
        "italic": True,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "black",
        "font_color": "white",
        "font_size": 12,
        "border_color": "white",
    })
    col_header_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "black",
        "font_color": "white",
        "font_size": 11,
    })
    green_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#9ff5b6",
        "font_size": 11,
        "text_wrap": True,
    })
    yellow_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#fcfa86",
        "font_size": 11,
        "text_wrap": True,
    })
    orange_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#fcd186",
        "font_size": 11,
        "text_wrap": True,
    })
    red_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#ff9191",
        "font_size": 11,
        "text_wrap": True,
    })
    purple_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#d791ff",
        "font_size": 11,
        "text_wrap": True,
    })
    maroon_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#b53c7a",
        "font_size": 11,
        "text_wrap": True,
    })
    no_loc_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "bold": True,
        "font_color": "red"
    })
    disclaimer_format = workbook.add_format({
        "italic": True,
        "align": "center",
        "valign": "vcenter",
        "font_size": 11,
        "text_wrap": True,
    })

    formats = [
        title_format,
        date_format,
        col_header_format,
        green_format,
        yellow_format,
        orange_format,
        red_format,
        purple_format,
        maroon_format,
        no_loc_format,
        disclaimer_format,
    ]
    return formats, workbook, workbook_filename

def observed_aqi(lat, lon):

    endpoint_url = "https://api.aerisapi.com/airquality/{},{}"
    forecast_r = requests.get(
        endpoint_url.format(lat, lon),
        params={
            "client_id": AERIS_ID,
            "client_secret": AERIS_SECRET
        },
    )
    response = forecast_r.json()
    if response["success"]:
        try:
            data = response["response"][0]
            period = data["periods"][0]
            for pollutant in period["pollutants"]:
                if pollutant["type"] == "pm2.5":
                    if not isinstance(pollutant["aqi"], int):
                        return 0
                    else:
                        return pollutant["aqi"]
                else: continue
        except:
            pass
    else: return 0

def forecast_aqi(lat, lon):
    endpoint_url = "https://api.aerisapi.com/airquality/forecasts/{},{}"
    forecast_r = requests.get(
        endpoint_url.format(lat, lon),
        params={
            "client_id": AERIS_ID,
            "client_secret": AERIS_SECRET
        },
    )
    response = forecast_r.json()
    if response["success"]:
        try:
            data = response["response"][0]
            period = data["periods"][0]
            for pollutant in period["pollutants"]:
                if pollutant["type"] == "pm2.5":
                    if not isinstance(pollutant["aqi"], int):
                        return 0
                    else:
                        return pollutant["aqi"]
                else: continue
        except: 
            pass
    else: return 0

def red_flag(lat, lon):
    endpoint_url = "https://api.aerisapi.com/alerts/"
    p = "{},{}".format(lat, lon)
    params = {
                "query": "type:FW.W",
                "client_id": AERIS_ID,
                "client_secret": AERIS_SECRET,
                "p": p
            }
    r = requests.get(endpoint_url.format(lat, lon), params=params)
    # print(r.content)
    # print(r.json().get('success'))
    return r

def fire_watch(lat, lon):
    endpoint_url = "https://api.aerisapi.com/alerts/"
    p = "{},{}".format(lat, lon)
    params = {
                "query": "type:FW.A",
                "client_id": AERIS_ID,
                "client_secret": AERIS_SECRET,
                "p": p
            }
    r = requests.get(endpoint_url.format(lat, lon), params=params)
    # print(r.content)
    # print(r.json().get('success'))
    return r

def aqi_today(formats, workbook, clientId, name, threshold, aqistates, f):
    print("aqi_today")
    bool_populate = False
    workbook_sheet_name = f"{datetime.now():%Y-%m-%d}"
    worksheet1 = workbook.add_worksheet(workbook_sheet_name)
    
    last_column = "F"
    title_format = formats[0]
    date_format = formats[1]
    col_header_format = formats[2]
    green_format = formats[3]
    yellow_format = formats[4]
    orange_format = formats[5]
    red_format = formats[6]
    purple_format = formats[7]
    maroon_format = formats[8]
    no_loc_format = formats[9]
    disclaimer_format = formats[10]
    date = datetime.now()

    worksheet1.set_row(0, 57)
    worksheet1.set_column(0, 0, 40)
    worksheet1.set_column(1, 1, 25)
    worksheet1.set_column(2, 2, 20)
    worksheet1.set_column(3, 4, 15)
    # ea_logo = os.path.join("eapytools", "static", "eapytools", "images",
    #                        "ealogo white.png")
    # worksheet1.insert_image("D1", ea_logo, {"x_offset": 0, "y_offset": 5})
    worksheet1.merge_range(
        f"A1:{last_column}1",
        f"{name} Locations with PM 2.5 AQI > {threshold}",
        title_format,
    )
    worksheet1.merge_range(f"A2:{last_column}2", f"Valid: {date:%a %b %d, %Y}", date_format)
    worksheet1.write(2, 0, "Name", col_header_format)
    worksheet1.write(2, 1, "Code", col_header_format)
    worksheet1.write(2, 2, "City", col_header_format)
    worksheet1.write(2, 3, "State", col_header_format)
    worksheet1.write(2, 4, "Zipcode", col_header_format)
    worksheet1.write(2, 5, "AQI", col_header_format)
    row = 2
    print("Today ----> ", len(f))
    hit_count = 0
    # for i in range(len(f.features)):
    for i in range(10):
        lat = f[i]["attributes"]['lat']
        lon = f[i]["attributes"]['lon']
        if isinstance(lat, float) and isinstance(lon, float):
            if hit_count == 100:
                print("Sleeping")
                time.sleep(60)
                hit_count = 0
            else:
                aqi = observed_aqi(lat, lon)
                if aqi > threshold:
                    print(f"Today: {i} & AQI: {aqi}")
                    row += 1
                    bool_populate = True
                    json_output = f.features[i].attributes
                    loc_name = json_output.get('name')
                    loc_code = json_output.get('code')
                    loc_city = json_output.get('city')
                    loc_state = json_output.get('state')
                    # if clientId == 219:
                    #     insert_loc_db(loc_name, clientId, loc_code, loc_city, loc_state, aqi)
                    if aqi < 50:
                        row_format = green_format
                    elif 51 <= aqi < 101:
                        row_format = yellow_format
                    elif 101 <= aqi < 151:
                        row_format = orange_format
                    elif 151 <= aqi < 200:
                        row_format = red_format
                    elif 200 <= aqi < 301:
                        row_format = purple_format
                    else:
                        row_format = maroon_format
                    worksheet1.write(row, 0, json_output.get('name'), row_format)
                    worksheet1.write(row, 1, json_output.get('code'), row_format)
                    worksheet1.write(row, 2, json_output.get('city'), row_format)
                    worksheet1.write(row, 3, json_output.get('state'), row_format)
                    worksheet1.write(row, 4, json_output.get('Zipcode'), row_format)
                    worksheet1.write(row, 5, aqi, row_format)
                    print(f"hit_count {hit_count}")
                    hit_count += 1
                else:
                    hit_count += 1
                    print(f"else hit_count {hit_count}")

            
        else:
            print("not float")
    if bool_populate == False:
        worksheet1.merge_range("A4:F4", "No locations exceeded the threshold.", no_loc_format)

    aqi_legend_row = row + 2
    worksheet1.set_row(aqi_legend_row, 15)
    worksheet1.write(aqi_legend_row, 0, "AQI Color", title_format)
    worksheet1.write(aqi_legend_row, 1, "Level of Concern", title_format)
    worksheet1.write(aqi_legend_row, 2, "Values of Index", title_format)
    worksheet1.merge_range(aqi_legend_row, 3, aqi_legend_row, 4, "Description", title_format)

    worksheet1.set_row(aqi_legend_row + 1, 80)
    worksheet1.write(aqi_legend_row + 1, 0, "Green", green_format)
    worksheet1.write(aqi_legend_row + 1, 1, "Good", green_format)
    worksheet1.write(aqi_legend_row + 1, 2, "0 to 50", green_format)
    text_string = (
        "Air quality is satisfactory, and air pollution poses little or no risk."
    )
    worksheet1.merge_range(aqi_legend_row + 1, 3, aqi_legend_row + 1, 4, text_string, green_format)

    worksheet1.set_row(aqi_legend_row + 2, 80)
    worksheet1.write(aqi_legend_row + 2, 0, "Yellow", yellow_format)
    worksheet1.write(aqi_legend_row + 2, 1, "Moderate", yellow_format)
    worksheet1.write(aqi_legend_row + 2, 2, "51 to 100", yellow_format)
    text_string = (
        "Air quality is acceptable. However, there may be a risk for some people, particularly those "
        "who are unusually sensitive to air pollution.")
    worksheet1.merge_range(aqi_legend_row + 2, 3, aqi_legend_row + 2, 4, text_string, yellow_format)

    worksheet1.set_row(aqi_legend_row + 3, 80)
    worksheet1.write(aqi_legend_row + 3, 0, "Orange", orange_format)
    worksheet1.write(aqi_legend_row + 3, 1, "Unhealthy for Sensitive Groups",
                    orange_format)
    worksheet1.write(aqi_legend_row + 3, 2, "101 to 150", orange_format)
    text_string = (
        "Members of sensitive groups may experience health effects. The general public is less likely "
        "to be affected.")
    worksheet1.merge_range(aqi_legend_row + 3, 3, aqi_legend_row + 3, 4, text_string, orange_format)

    worksheet1.set_row(aqi_legend_row + 4, 80)
    worksheet1.write(aqi_legend_row + 4, 0, "Red", red_format)
    worksheet1.write(aqi_legend_row + 4, 1, "Unhealthy", red_format)
    worksheet1.write(aqi_legend_row + 4, 2, "151 to 200", red_format)
    text_string = (
        "Some members of the general public may experience health effects; members of sensitive groups "
        "may experience more serious health effects.")
    worksheet1.merge_range(aqi_legend_row + 4, 3, aqi_legend_row + 4, 4, text_string, red_format)

    worksheet1.set_row(aqi_legend_row + 5, 80)
    worksheet1.write(aqi_legend_row + 5, 0, "Purple", purple_format)
    worksheet1.write(aqi_legend_row + 5, 1, "Very Unhealthy", purple_format)
    worksheet1.write(aqi_legend_row + 5, 2, "201 to 300", purple_format)
    text_string = (
        "Health alert: The risk of health effects is increased for everyone.")
    worksheet1.merge_range(aqi_legend_row + 5, 3, aqi_legend_row + 5, 4, text_string, purple_format)

    worksheet1.set_row(aqi_legend_row + 6, 80)
    worksheet1.write(aqi_legend_row + 6, 0, "Maroon", maroon_format)
    worksheet1.write(aqi_legend_row + 6, 1, "Hazardous", maroon_format)
    worksheet1.write(aqi_legend_row + 6, 2, "301 and higher", maroon_format)
    text_string = "Health warning of emergency conditions: everyone is more likely to be affected."
    worksheet1.merge_range(aqi_legend_row + 6, 3, aqi_legend_row + 6, 4, text_string, maroon_format)
    return worksheet1, workbook, workbook_sheet_name, row

def aqi_tomorrow(formats, workbook, clientId, name, threshold, aqistates,f, workbook_filename):
    print("aqi_tomorrow")
    # workbook = xlsxwriter.Workbook(os.path.join("/tmp", workbook_filename))
    bool_populate = False
    date = datetime.now() + timedelta(days=1)
    date2 = date.strftime('%Y-%m-%d') 
    worksheet1 = workbook.add_worksheet(f"{date2}")
    
    last_column = "F"
    title_format = formats[0]
    date_format = formats[1]
    col_header_format = formats[2]
    green_format = formats[3]
    yellow_format = formats[4]
    orange_format = formats[5]
    red_format = formats[6]
    purple_format = formats[7]
    maroon_format = formats[8]
    no_loc_format = formats[9]
    disclaimer_format = formats[10]
    

    worksheet1.set_row(0, 57)
    worksheet1.set_column(0, 0, 40)
    worksheet1.set_column(1, 1, 25)
    worksheet1.set_column(2, 2, 20)
    worksheet1.set_column(3, 4, 15)
    # ea_logo = os.path.join("eapytools", "static", "eapytools", "images",
    #                        "ealogo white.png")
    # worksheet1.insert_image("D1", ea_logo, {"x_offset": 0, "y_offset": 5})
    worksheet1.merge_range(
        f"A1:{last_column}1",
        f"{name} Locations with PM 2.5 AQI > {threshold}",
        title_format,
    )
    worksheet1.merge_range(f"A2:{last_column}2", f"Valid: {date:%a %b %d, %Y}",date_format)
    worksheet1.write(2, 0, "Name", col_header_format)
    worksheet1.write(2, 1, "Code", col_header_format)
    worksheet1.write(2, 2, "City", col_header_format)
    worksheet1.write(2, 3, "State", col_header_format)
    worksheet1.write(2, 4, "Zipcode", col_header_format)
    worksheet1.write(2, 5, "AQI", col_header_format)
    row = 2
    # print(f.features[0].attributes)
    hit_count = 0
    logging.info("Tomorrow ----> ", len(f))
    # for i in range(len(f.features)):
    for i in range(10):
        lat = f[i]["attributes"]["lat"]
        lon = f[i]["attributes"]["lon"]
        # print(str(lat) + " " + str(lon))
        if hit_count == 100:
            print("Sleeping")
            time.sleep(60)
            hit_count = 0
        else:
            aqi = forecast_aqi(lat, lon)
            if aqi > threshold:
                print(f"tomm : {i} and AQI: {aqi}")
                row += 1
                bool_populate = True
                json_output = f.features[i].attributes
                loc_name = json_output.get('name')
                loc_code = json_output.get('code')
                loc_city = json_output.get('city')
                loc_state = json_output.get('state')
                # if clientId == 219:
                    # insert_loc_db(loc_name, clientId, loc_code, loc_city, loc_state, aqi)
                if aqi < 50:
                    row_format = green_format
                elif 51 <= aqi < 101:
                    row_format = yellow_format
                elif 101 <= aqi < 151:
                    row_format = orange_format
                elif 151 <= aqi < 200:
                    row_format = red_format
                elif 200 <= aqi < 301:
                    row_format = purple_format
                else:
                    row_format = maroon_format
                worksheet1.write(row, 0, json_output.get('name'), row_format)
                worksheet1.write(row, 1, json_output.get('code'), row_format)
                worksheet1.write(row, 2, json_output.get('city'), row_format)
                worksheet1.write(row, 3, json_output.get('state'), row_format)
                worksheet1.write(row, 4, json_output.get('Zipcode'), row_format)
                worksheet1.write(row, 5, aqi, row_format)
                hit_count += 1
                print(f"hit_count {hit_count}")
            else:
                hit_count += 1
                print(f"else hit_count {hit_count}")




    if bool_populate == False:
        worksheet1.merge_range("A4:F4", "No locations forecast to exceed the threshold.", no_loc_format)

    aqi_legend_row = row + 2
    worksheet1.set_row(aqi_legend_row, 15)
    worksheet1.write(aqi_legend_row, 0, "AQI Color", title_format)
    worksheet1.write(aqi_legend_row, 1, "Level of Concern", title_format)
    worksheet1.write(aqi_legend_row, 2, "Values of Index", title_format)
    worksheet1.merge_range(aqi_legend_row, 3, aqi_legend_row, 4, "Description", title_format)

    worksheet1.set_row(aqi_legend_row + 1, 80)
    worksheet1.write(aqi_legend_row + 1, 0, "Green", green_format)
    worksheet1.write(aqi_legend_row + 1, 1, "Good", green_format)
    worksheet1.write(aqi_legend_row + 1, 2, "0 to 50", green_format)
    text_string = (
        "Air quality is satisfactory, and air pollution poses little or no risk."
    )
    worksheet1.merge_range(aqi_legend_row + 1, 3, aqi_legend_row + 1, 4, text_string, green_format)

    worksheet1.set_row(aqi_legend_row + 2, 80)
    worksheet1.write(aqi_legend_row + 2, 0, "Yellow", yellow_format)
    worksheet1.write(aqi_legend_row + 2, 1, "Moderate", yellow_format)
    worksheet1.write(aqi_legend_row + 2, 2, "51 to 100", yellow_format)
    text_string = (
        "Air quality is acceptable. However, there may be a risk for some people, particularly those "
        "who are unusually sensitive to air pollution.")
    worksheet1.merge_range(aqi_legend_row + 2, 3, aqi_legend_row + 2, 4, text_string, yellow_format)

    worksheet1.set_row(aqi_legend_row + 3, 80)
    worksheet1.write(aqi_legend_row + 3, 0, "Orange", orange_format)
    worksheet1.write(aqi_legend_row + 3, 1, "Unhealthy for Sensitive Groups",
                    orange_format)
    worksheet1.write(aqi_legend_row + 3, 2, "101 to 150", orange_format)
    text_string = (
        "Members of sensitive groups may experience health effects. The general public is less likely "
        "to be affected.")
    worksheet1.merge_range(aqi_legend_row + 3, 3, aqi_legend_row + 3, 4, text_string, orange_format)

    worksheet1.set_row(aqi_legend_row + 4, 80)
    worksheet1.write(aqi_legend_row + 4, 0, "Red", red_format)
    worksheet1.write(aqi_legend_row + 4, 1, "Unhealthy", red_format)
    worksheet1.write(aqi_legend_row + 4, 2, "151 to 200", red_format)
    text_string = (
        "Some members of the general public may experience health effects; members of sensitive groups "
        "may experience more serious health effects.")
    worksheet1.merge_range(aqi_legend_row + 4, 3, aqi_legend_row + 4, 4, text_string, red_format)

    worksheet1.set_row(aqi_legend_row + 5, 80)
    worksheet1.write(aqi_legend_row + 5, 0, "Purple", purple_format)
    worksheet1.write(aqi_legend_row + 5, 1, "Very Unhealthy", purple_format)
    worksheet1.write(aqi_legend_row + 5, 2, "201 to 300", purple_format)
    text_string = (
        "Health alert: The risk of health effects is increased for everyone.")
    worksheet1.merge_range(aqi_legend_row + 5, 3, aqi_legend_row + 5, 4, text_string, purple_format)

    worksheet1.set_row(aqi_legend_row + 6, 80)
    worksheet1.write(aqi_legend_row + 6, 0, "Maroon", maroon_format)
    worksheet1.write(aqi_legend_row + 6, 1, "Hazardous", maroon_format)
    worksheet1.write(aqi_legend_row + 6, 2, "301 and higher", maroon_format)
    text_string = "Health warning of emergency conditions: everyone is more likely to be affected."
    worksheet1.merge_range(aqi_legend_row + 6, 3, aqi_legend_row + 6, 4, text_string, maroon_format)
    return worksheet1, workbook, date2, row

def red_flag_warning(formats, workbook, clientId, name, threshold, aqistates,f):
    # print("red_flag_warning")
    bool_populate = False
    workbook_sheet_name = "Red Flag Warnings"
    worksheet = workbook.add_worksheet(workbook_sheet_name)
    last_column = "F"
    title_format = formats[0]
    date_format = formats[1]
    col_header_format = formats[2]
    green_format = formats[3]
    yellow_format = formats[4]
    orange_format = formats[5]
    red_format = formats[6]
    purple_format = formats[7]
    maroon_format = formats[8]
    no_loc_format = formats[9]
    disclaimer_format = formats[10]
    date = datetime.now()
    worksheet.set_row(0, 57)
    worksheet.set_column(0, 0, 40)
    worksheet.set_column(1, 1, 25)
    worksheet.set_column(2, 2, 20)
    worksheet.set_column(3, 3, 15)
    # ea_logo = os.path.join(
    #             "eapytools", "static", "eapytools", "images", "ealogo white.png"
    #         )
    # worksheet.insert_image(
    #             "D1",
    #             ea_logo,
    #             {"x_scale": 0.8, "y_scale": 0.8, "x_offset": -65, "y_offset": 11},
    #         )
    worksheet.merge_range(
                "A1:C1", f"{name} Locations in Red Flag Warnings", title_format
            )
    worksheet.write(0, 3, "", title_format)
    worksheet.merge_range(
                "A2:D2", f"Valid as of {date:%I:%M %p %Z: %a %b %d, %Y}", date_format
            )
    worksheet.write(2, 0, "Name", col_header_format)
    worksheet.write(2, 1, "Code", col_header_format)
    worksheet.write(2, 2, "City", col_header_format)
    worksheet.write(2, 3, "State", col_header_format)
            # write location data
    row = 2
    hit_count = 0
    for i in range(len(f.features)):
        lat = f.features[i].attributes.get('lat')
        lon = f.features[i].attributes.get('lon')
        if hit_count == 300:
            print("Sleeping")
            time.sleep(60)
            hit_count = 0
        else:
            r = red_flag(lat, lon) 
            if r.json().get('success') == True:
                if r.json().get('error') == None:
                    bool_populate = True
                    row += 1
                    json_output = f.features[i].attributes
                    worksheet.write(row, 0, json_output.get('name'), red_format)
                    worksheet.write(row, 1, json_output.get('code'), red_format)
                    worksheet.write(row, 2, json_output.get('city'), red_format)
                    worksheet.write(row, 3, json_output.get('state'), red_format)
                    hit_count += 1
            else:
                print("FAILED")
    if bool_populate == False:
        worksheet.merge_range(
                    "A4:D4", "No locations in Red Flag Warnings today.", no_loc_format
                )
    return worksheet, workbook, workbook_sheet_name, row

def fire_weather_watch(formats, workbook, clientId, name, threshold, aqistates,f):
    # print("fire_weather_watch")
    bool_populate = False
    workbook_sheet_name = "Fire Weather Watch"
    worksheet = workbook.add_worksheet(workbook_sheet_name)
    last_column = "F"
    title_format = formats[0]
    date_format = formats[1]
    col_header_format = formats[2]
    green_format = formats[3]
    yellow_format = formats[4]
    orange_format = formats[5]
    red_format = formats[6]
    purple_format = formats[7]
    maroon_format = formats[8]
    no_loc_format = formats[9]
    disclaimer_format = formats[10]
    date = datetime.now()
    worksheet.set_row(0, 57)
    worksheet.set_column(0, 0, 40)
    worksheet.set_column(1, 1, 25)
    worksheet.set_column(2, 2, 20)
    worksheet.set_column(3, 3, 15)
    # ea_logo = os.path.join(
    #             "eapytools", "static", "eapytools", "images", "ealogo white.png"
    #         )
    # worksheet.insert_image(
    #             "D1",
    #             ea_logo,
    #             {"x_scale": 0.8, "y_scale": 0.8, "x_offset": -65, "y_offset": 11},
    #         )
    worksheet.merge_range(
                "A1:C1",
                f"{name} Locations in Fire Weather Watches",
                title_format,
            )
    worksheet.write(0, 3, "", title_format)
    worksheet.merge_range(
                "A2:D2", f"Issued at {date:%I:%M %p %Z: %a %b %d, %Y}", date_format
            )
    worksheet.write(2, 0, "Name", col_header_format)
    worksheet.write(2, 1, "Code", col_header_format)
    worksheet.write(2, 2, "City", col_header_format)
    worksheet.write(2, 3, "State", col_header_format)
    row = 2
    hit_count = 0
    for i in range(len(f.features)):
        lat = f.features[i].attributes.get('lat')
        lon = f.features[i].attributes.get('lon')
        if hit_count == 300:
            print("Sleeping")
            time.sleep(60)
            hit_count = 0
        else:
            r = fire_watch(lat, lon) 
            if r.json().get('success') == True:
                if r.json().get('error') == None:
                    bool_populate = True
                    row += 1
                    json_output = f.features[i].attributes
                    worksheet.write(row, 0, json_output.get('name'), red_format)
                    worksheet.write(row, 1, json_output.get('code'), red_format)
                    worksheet.write(row, 2, json_output.get('city'), red_format)
                    worksheet.write(row, 3, json_output.get('state'), red_format)
                    hit_count += 1
            else:
                print("FAILED")
    if bool_populate == False:
        worksheet.merge_range("A4:D4", "No locations in Fire Weather Watch today.", no_loc_format)
    return worksheet, workbook, workbook_sheet_name, row

def send_aqi_email(clientId, filename, email_list, workbook_filename):
        """Sends daily email with spreadsheet to clients."""
        sub = email_list
        email_list = ['yasir.khan@cooperativecomputing.com']
        # email_list = ['daldret@earlyalert.com']
        
        send_email = requests.post("https://api.mailgun.net/v3/earlyalert.com/messages",
        auth=("api", "key-abc3ac7030c2113b91c27b6733ebe510"),
            files = [("attachment", (f"{filename}", open(f"{filename}", "rb").read()))],
            data={"from": "airquality@earlyalert.com",
                "to": email_list,
                "bcc" : [
                    "EAInformalGroupA@earlyalert.com", 
                "ykhan@earlyalert.com"],
                "subject": f"See attached for the Daily Air Quality Report.",
                "text": "Excel File for Daily Air Quality Report",
                # "template": template_name,
                # "h:X-Mailgun-Variables": dynamic,
                # "o:tag": tags_lists
                })
        print("********** Email Sent **********", filename, send_email.text)
        return send_email


get_each_client([369])