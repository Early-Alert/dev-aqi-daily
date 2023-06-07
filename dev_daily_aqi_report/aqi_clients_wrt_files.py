import json, os, xlsxwriter
from datetime import datetime, timedelta, timezone
import requests, logging
import multiprocessing
import azure.functions as func


GIS_USERNAME = "Developer1"
GIS_PASSWORD = "qweRTY77**"
AERIS_ID = 'IJwChrr7utrp6BFWVnJ1A'
AERIS_SECRET = 'WmzZHc1PnJiAsr996tFNwfsPt7IqetRHHuB8Ldty'

def get_features(client_id, aqi_states):
    """Commented Out in order to avoid exception of package conflict"""
    # gis = GIS("https://maps.earlyalert.com/portal/home/", GIS_USERNAME, GIS_PASSWORD)
    # item = gis.content.get("3604f01a15274e139e47ba1fe00183aa")
    # feature_layer = item.layers[0]
    # if len(aqi_states) == 2:
    #     query = "ClientId = {} AND State IN ('{}')".format(client_id, aqi_states)
    # else:
    #     query = "ClientId = {} AND State IN {}".format(client_id, aqi_states)

    # f = feature_layer.query(where=query)
    # return f

def formats_workbook(name):
    workbook_filename = f"{name} Air Quality Forecast {datetime.now():%Y-%m-%d}.xlsx"
    workbook = xlsxwriter.Workbook(os.path.join("/tmp",
        workbook_filename))
    
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

def add_static_footer(formats, workbook_one, workbook_two, workbook_one_row, workbook_two_row):
    workbooks = [[workbook_one, workbook_one_row], [workbook_two, workbook_two_row]]

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

    for each_workbook in workbooks:

        aqi_legend_row = each_workbook[1] + 2
        each_workbook[0].set_row(aqi_legend_row, 15)
        each_workbook[0].write(aqi_legend_row, 0, "AQI Color", title_format)
        each_workbook[0].write(aqi_legend_row, 1, "Level of Concern", title_format)
        each_workbook[0].write(aqi_legend_row, 2, "Values of Index", title_format)
        each_workbook[0].merge_range(aqi_legend_row, 3, aqi_legend_row, 4, "Description", title_format)

        each_workbook[0].set_row(aqi_legend_row + 1, 80)
        each_workbook[0].write(aqi_legend_row + 1, 0, "Green", green_format)
        each_workbook[0].write(aqi_legend_row + 1, 1, "Good", green_format)
        each_workbook[0].write(aqi_legend_row + 1, 2, "0 to 50", green_format)
        text_string = (
            "Air quality is satisfactory, and air pollution poses little or no risk."
        )
        each_workbook[0].merge_range(aqi_legend_row + 1, 3, aqi_legend_row + 1, 4, text_string, green_format)

        each_workbook[0].set_row(aqi_legend_row + 2, 80)
        each_workbook[0].write(aqi_legend_row + 2, 0, "Yellow", yellow_format)
        each_workbook[0].write(aqi_legend_row + 2, 1, "Moderate", yellow_format)
        each_workbook[0].write(aqi_legend_row + 2, 2, "51 to 100", yellow_format)
        text_string = (
            "Air quality is acceptable. However, there may be a risk for some people, particularly those "
            "who are unusually sensitive to air pollution.")
        each_workbook[0].merge_range(aqi_legend_row + 2, 3, aqi_legend_row + 2, 4, text_string, yellow_format)

        each_workbook[0].set_row(aqi_legend_row + 3, 80)
        each_workbook[0].write(aqi_legend_row + 3, 0, "Orange", orange_format)
        each_workbook[0].write(aqi_legend_row + 3, 1, "Unhealthy for Sensitive Groups",
                        orange_format)
        each_workbook[0].write(aqi_legend_row + 3, 2, "101 to 150", orange_format)
        text_string = (
            "Members of sensitive groups may experience health effects. The general public is less likely "
            "to be affected.")
        each_workbook[0].merge_range(aqi_legend_row + 3, 3, aqi_legend_row + 3, 4, text_string, orange_format)

        each_workbook[0].set_row(aqi_legend_row + 4, 80)
        each_workbook[0].write(aqi_legend_row + 4, 0, "Red", red_format)
        each_workbook[0].write(aqi_legend_row + 4, 1, "Unhealthy", red_format)
        each_workbook[0].write(aqi_legend_row + 4, 2, "151 to 200", red_format)
        text_string = (
            "Some members of the general public may experience health effects; members of sensitive groups "
            "may experience more serious health effects.")
        each_workbook[0].merge_range(aqi_legend_row + 4, 3, aqi_legend_row + 4, 4, text_string, red_format)

        each_workbook[0].set_row(aqi_legend_row + 5, 80)
        each_workbook[0].write(aqi_legend_row + 5, 0, "Purple", purple_format)
        each_workbook[0].write(aqi_legend_row + 5, 1, "Very Unhealthy", purple_format)
        each_workbook[0].write(aqi_legend_row + 5, 2, "201 to 300", purple_format)
        text_string = (
            "Health alert: The risk of health effects is increased for everyone.")
        each_workbook[0].merge_range(aqi_legend_row + 5, 3, aqi_legend_row + 5, 4, text_string, purple_format)

        each_workbook[0].set_row(aqi_legend_row + 6, 80)
        each_workbook[0].write(aqi_legend_row + 6, 0, "Maroon", maroon_format)
        each_workbook[0].write(aqi_legend_row + 6, 1, "Hazardous", maroon_format)
        each_workbook[0].write(aqi_legend_row + 6, 2, "301 and higher", maroon_format)
        text_string = "Health warning of emergency conditions: everyone is more likely to be affected."
        each_workbook[0].merge_range(aqi_legend_row + 6, 3, aqi_legend_row + 6, 4, text_string, maroon_format)

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
                    # "EAInformalGroupA@earlyalert.com", 
                "ykhan@earlyalert.com"],
                "subject": f"See attached for the Daily Air Quality Report.",
                "text": "Excel File for Daily Air Quality Report",
                # "template": template_name,
                # "h:X-Mailgun-Variables": dynamic,
                # "o:tag": tags_lists
                })
        print("********** Email Sent **********", filename, send_email.text)
        return send_email

def get_batch_aqi(lat, lon):
    # endpoint_url = f"https://api.aerisapi.com/batch/{28.150584},{34.6209471}?requests=/airquality,/airquality/forecasts/"
    endpoint_url = f"https://api.aerisapi.com/batch/{lat},{lon}?requests=/airquality,/airquality/forecasts/"
    max_retries = 5
    params={
        "client_id": AERIS_ID,
        "client_secret": AERIS_SECRET
    }

    retries = 0
    while retries < max_retries:
        try:
            response = requests.post(url=endpoint_url, params=params)
            if response.status_code == 200:
                return response.json()  # Successful response, return the result
        except ConnectionError:
            pass  # Connection error occurred, retry the request
        
        retries += 1
    return None

def prepare_report_files(formats, workbook, clientId, name, aqithreshold, aqistates, f):
    logging.info("prepare_report_files")

    workbook1_bool_populate = False
    workbook2_bool_populate = False


    workbook_sheet_name = f"{datetime.now():%Y-%m-%d}"
    worksheet1 = workbook.add_worksheet(workbook_sheet_name)
    
    date = datetime.now()
    date_tomorrow = datetime.now() + timedelta(days=1)
    # date2 = date.strftime('%Y-%m-%d') 
    worksheet2 = workbook.add_worksheet(f"{date_tomorrow.strftime('%Y-%m-%d') }")

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

    worksheet2.set_row(0, 57)
    worksheet2.set_column(0, 0, 40)
    worksheet2.set_column(1, 1, 25)
    worksheet2.set_column(2, 2, 20)
    worksheet2.set_column(3, 4, 15)

    worksheet1.merge_range(
        f"A1:{last_column}1",
        f"{name} Locations with PM 2.5 AQI > {aqithreshold}",
        title_format,
    )

    worksheet2.merge_range(
        f"A1:{last_column}1",
        f"{name} Locations with PM 2.5 AQI > {aqithreshold}",
        title_format,
    )
    
    worksheet1.merge_range(f"A2:{last_column}2", f"Valid: {date:%a %b %d, %Y}", date_format)
    worksheet2.merge_range(f"A2:{last_column}2", f"Valid: {date_tomorrow:%a %b %d, %Y}",date_format)

    
    worksheet1.write(2, 0, "Name", col_header_format)
    worksheet1.write(2, 1, "Code", col_header_format)
    worksheet1.write(2, 2, "City", col_header_format)
    worksheet1.write(2, 3, "State", col_header_format)
    worksheet1.write(2, 4, "Zipcode", col_header_format)
    worksheet1.write(2, 5, "AQI", col_header_format)
    
    worksheet2.write(2, 0, "Name", col_header_format)
    worksheet2.write(2, 1, "Code", col_header_format)
    worksheet2.write(2, 2, "City", col_header_format)
    worksheet2.write(2, 3, "State", col_header_format)
    worksheet2.write(2, 4, "Zipcode", col_header_format)
    worksheet2.write(2, 5, "AQI", col_header_format)

    workbook1_row = 2
    workbook2_row = 2
    print(f"---> {len(f['features'])}")

    for i in range(len(f['features'])):
    # for i in range(10):
    
        lat = f['features'][i]['lat']
        lon = f['features'][i]['lon']
        if isinstance(lat, float) and isinstance(lon, float):
            aqi_response = get_batch_aqi(lat, lon)
            # logging.info(f"aqi_response --> {aqi_response}")
            if aqi_response["success"]:
                if aqi_response["response"]["responses"][0]["success"]:
                    response_today_aqi = aqi_response["response"]["responses"][0]["response"][0]["periods"][0]["pollutants"]
                    today_aqi = next((p for p in response_today_aqi if p["type"] == "pm2.5"), None)["aqi"]
                if aqi_response["response"]["responses"][1]["success"]:
                    responst_forecast_aqi = aqi_response["response"]["responses"][1]["response"][0]["periods"][0]["pollutants"]
                    forecast_aqi = next((p for p in responst_forecast_aqi if p["type"] == "pm2.5"), None)["aqi"]
                print(f"{i} -- Today -- {today_aqi} & Forecast -- {forecast_aqi}")
                if today_aqi and today_aqi > aqithreshold:
                # if today_aqi > 10:
                    print(f"Today: {i} & AQI: {today_aqi}")
                    workbook1_row += 1
                    workbook1_bool_populate = True
                    json_output = f["features"][i]
                    loc_name = json_output.get('name')
                    loc_code = json_output.get('code')
                    loc_city = json_output.get('city')
                    loc_state = json_output.get('state')
                    # if clientId == 219:
                    #     insert_loc_db(loc_name, clientId, loc_code, loc_city, loc_state, aqi)
                    if today_aqi < 50:
                        row_format = green_format
                    elif 51 <= today_aqi < 101:
                        row_format = yellow_format
                    elif 101 <= today_aqi < 151:
                        row_format = orange_format
                    elif 151 <= today_aqi < 200:
                        row_format = red_format
                    elif 200 <= today_aqi < 301:
                        row_format = purple_format
                    else:
                        row_format = maroon_format
                    worksheet1.write(workbook1_row, 0, json_output.get('name'), row_format)
                    worksheet1.write(workbook1_row, 1, json_output.get('code'), row_format)
                    worksheet1.write(workbook1_row, 2, json_output.get('city'), row_format)
                    worksheet1.write(workbook1_row, 3, json_output.get('state'), row_format)
                    worksheet1.write(workbook1_row, 4, json_output.get('Zipcode'), row_format)
                    worksheet1.write(workbook1_row, 5, today_aqi, row_format)
                else:
                    pass
                
                if forecast_aqi and forecast_aqi > aqithreshold:
                # if forecast_aqi > 10:
                    print(f"Forecast: {i} & AQI: {today_aqi}")
                    workbook2_row += 1
                    workbook2_bool_populate = True
                    json_output = f["features"][i]
                    loc_name = json_output.get('name')
                    loc_code = json_output.get('code')
                    loc_city = json_output.get('city')
                    loc_state = json_output.get('state')
                    # if clientId == 219:
                        # insert_loc_db(loc_name, clientId, loc_code, loc_city, loc_state, aqi)
                    if forecast_aqi < 50:
                        row_format = green_format
                    elif 51 <= forecast_aqi < 101:
                        row_format = yellow_format
                    elif 101 <= forecast_aqi < 151:
                        row_format = orange_format
                    elif 151 <= forecast_aqi < 200:
                        row_format = red_format
                    elif 200 <= forecast_aqi < 301:
                        row_format = purple_format
                    else:
                        row_format = maroon_format
                    worksheet2.write(workbook2_row, 0, json_output.get('name'), row_format)
                    worksheet2.write(workbook2_row, 1, json_output.get('code'), row_format)
                    worksheet2.write(workbook2_row, 2, json_output.get('city'), row_format)
                    worksheet2.write(workbook2_row, 3, json_output.get('state'), row_format)
                    worksheet2.write(workbook2_row, 4, json_output.get('Zipcode'), row_format)
                    worksheet2.write(workbook2_row, 5, forecast_aqi, row_format)
                else:
                    pass            
        else:
            print("not float")
    if workbook1_bool_populate == False:
        worksheet1.merge_range("A4:F4", "No locations exceeded the threshold.", no_loc_format)
    if workbook2_bool_populate == False:
        worksheet2.merge_range("A4:F4", "No locations forecast to exceed the threshold.", no_loc_format)
    
    add_static_footer(formats, worksheet1, worksheet2, workbook1_row, workbook2_row)
    
    return worksheet1, workbook, workbook_sheet_name, workbook1_row


def get_each_client(client_data):
    logging.info("get_each_client")
    for each_id in range(len(client_data["client_id"])):  
        cid = client_data["client_id"][each_id]
        name = client_data["name"][each_id]
        aqi_threshold = client_data["aqi_threshold"][each_id]
        aqi_states = client_data["aqi_states"][each_id]
        email_list = client_data["email_list"][each_id]
        logging.info(f"{cid}, {name}, {aqi_threshold}, {aqi_states}, {email_list}")
        create_report(cid, name, aqi_threshold, aqi_states, email_list)

def create_report(clientId, name, threshold, aqistates, email_list):
    logging.info("create_report")
    formats, workbook, workbook_filename = formats_workbook(name)
    
    """These Lines are use to create new json file for each client locations"""
    # features = get_features(clientId, aqistates)
    # features_dict = [feature.attributes for feature in features.features]
    # feature_set_dict = {
    #     "features": features_dict,
    # }
    # json_data = json.dumps(feature_set_dict)
    # with open(f"{clientId}_{name}.json", "w") as fp:
    #     fp.write(json_data)

    if clientId == 369:
        feature_variable = "369_Home Depot Canada"
    elif clientId == 212:
        feature_variable = "212_PetSmart"
    elif clientId == 219:
        feature_variable = "219_Comcast"
    elif clientId == 238:
        feature_variable = "238_Office Depot"
    elif clientId == 35:
        feature_variable = "35_Home Depot"
    elif clientId == 188:
        feature_variable = "188_Charter Communications - Main - AQI 101"
    
    with open(f"{feature_variable}.json", "r") as file:
        json_features = json.load(file)
    

    worksheet1, workbook, workbook_sheet_name_today, max_row_today = prepare_report_files(formats, workbook, clientId, name, threshold, aqistates, json_features)
    workbook.close()
    print("CLOSED", name)
    # send_aqi_email(clientId, workbook.filename, email_list, workbook_filename)
    logging.info(f"Email Sent to {email_list}")



def dev_main():
    client_dict = {
        "client_id": [369, 212, 219, 238, 35, 188],
        "name": ["Home Depot Canada", "PetSmart", "Comcast", "Office Depot", "Home Depot", "Charter Communications - Main - AQI 101"],
        "aqi_threshold": [150, 150, 150, 150, 150, 101],
        "aqi_states": [('BC'), ('CA','ID','NV','OR','WA'), ('CA'), ('CA'), ('CA'), ('CA','CO','HI','ID','MT','NV','OR','WA','WY')],
        "email_list": [
            ['369_PrimaryDistribution@earlyalert.com', '369_West_Region@earlyalert.com'], 
            ['PetSmartAQIDistribution@earlyalert.com'], 
            ['219_california_region@earlyalert.com'], 
            ['OfficeDepot_AQI_Distribution@earlyalert.com'], 
            ['HomeDepot_AQI_Distribution@earlyalert.com'], 
            ['188_aqidistribution@earlyalert.com']
        ]
    }
    get_each_client(client_dict)



def main(mytimer: func.TimerRequest) -> None:
    utc_timestamp = datetime.utcnow().replace(
        tzinfo=timezone.utc).isoformat()
    logging.info(f"Function Triggered {utc_timestamp}")
    client_dict = {
        "client_id": [369, 212, 219, 238, 35, 188],
        "name": ["Home Depot Canada", "PetSmart", "Comcast", "Office Depot", "Home Depot", "Charter Communications - Main - AQI 101"],
        "aqi_threshold": [150, 150, 150, 150, 150, 101],
        "aqi_states": [('BC'), ('CA','ID','NV','OR','WA'), ('CA'), ('CA'), ('CA'), ('CA','CO','HI','ID','MT','NV','OR','WA','WY')],
        "email_list": [
            ['andrew_hamelin@homedepot.com'], 
            ['PetSmartAQIDistribution@earlyalert.com'], 
            ['219_california_region@earlyalert.com'], 
            ['OfficeDepot_AQI_Distribution@earlyalert.com'], 
            ['HomeDepot_AQI_Distribution@earlyalert.com'], 
            ['188_aqidistribution@earlyalert.com']
        ]
    }
    # get_each_client(client_dict)


    process_one_client_dict = {
        "client_id": [client_dict["client_id"][0]],
        "name": [client_dict["name"][0]],
        "aqi_threshold": [client_dict["aqi_threshold"][0]],
        "aqi_states": [client_dict["aqi_states"][0]],
        "email_list": [client_dict["email_list"][0]]
    }

    process_two_client_dict = {
        "client_id": [client_dict["client_id"][1]],
        "name": [client_dict["name"][1]],
        "aqi_threshold": [client_dict["aqi_threshold"][1]],
        "aqi_states": [client_dict["aqi_states"][1]],
        "email_list": [client_dict["email_list"][1]]
    }

    process_three_client_dict = {
        "client_id": [client_dict["client_id"][2]],
        "name": [client_dict["name"][2]],
        "aqi_threshold": [client_dict["aqi_threshold"][2]],
        "aqi_states": [client_dict["aqi_states"][2]],
        "email_list": [client_dict["email_list"][2]]
    }

    process_four_client_dict = {
        "client_id": [client_dict["client_id"][3]],
        "name": [client_dict["name"][3]],
        "aqi_threshold": [client_dict["aqi_threshold"][3]],
        "aqi_states": [client_dict["aqi_states"][3]],
        "email_list": [client_dict["email_list"][3]]
    }

    process_five_client_dict = {
        "client_id": [client_dict["client_id"][4]],
        "name": [client_dict["name"][4]],
        "aqi_threshold": [client_dict["aqi_threshold"][4]],
        "aqi_states": [client_dict["aqi_states"][4]],
        "email_list": [client_dict["email_list"][4]]
    }

    process_six_client_dict = {
        "client_id": [client_dict["client_id"][5]],
        "name": [client_dict["name"][5]],
        "aqi_threshold": [client_dict["aqi_threshold"][5]],
        "aqi_states": [client_dict["aqi_states"][5]],
        "email_list": [client_dict["email_list"][5]]
    }


    # get_each_client(client_dict)

    # process_one = multiprocessing.Process(target=get_each_client, args=(process_one_client_dict,))
    # process_two = multiprocessing.Process(target=get_each_client, args=(process_two_client_dict,))
    # process_three = multiprocessing.Process(target=get_each_client, args=(process_three_client_dict,))
    # process_four = multiprocessing.Process(target=get_each_client, args=(process_four_client_dict,))
    # process_five = multiprocessing.Process(target=get_each_client, args=(process_five_client_dict,))
    # process_six = multiprocessing.Process(target=get_each_client, args=(process_six_client_dict,))

    # process_one.start()
    # process_two.start()
    # process_three.start()
    # process_four.start()
    # process_five.start()
    # process_six.start()

    # process_one.join()
    # process_two.join()
    # process_three.join()
    # process_four.join()
    # process_five.join()
    # process_six.join()

    logging.info('Python timer trigger function ran at %s', utc_timestamp)


# def maiin():
#     utc_timestamp = datetime.utcnow().replace(
#         tzinfo=timezone.utc).isoformat()
#     print("GET TRIGERED")
#     client_dict = {
#         "client_id": [369, 212, 219, 238, 35, 188],
#         "name": ["Home Depot Canada", "PetSmart", "Comcast", "Office Depot", "Home Depot", "Charter Communications - Main - AQI 101"],
#         "aqi_threshold": [150, 150, 150, 150, 150, 101],
#         "aqi_states": [('BC'), ('CA','ID','NV','OR','WA'), ('CA'), ('CA'), ('CA'), ('CA','CO','HI','ID','MT','NV','OR','WA','WY')],
#         "email_list": [
#             ['369_PrimaryDistribution@earlyalert.com', '369_West_Region@earlyalert.com'], 
#             ['PetSmartAQIDistribution@earlyalert.com'], 
#             ['219_california_region@earlyalert.com'], 
#             ['OfficeDepot_AQI_Distribution@earlyalert.com'], 
#             ['HomeDepot_AQI_Distribution@earlyalert.com'], 
#             ['188_aqidistribution@earlyalert.com']
#         ]
#     }
#     get_each_client(client_dict)

#     logging.info('Python timer trigger function ran at %s', utc_timestamp)

# maiin()

def temp():
    client_dict = {
        "client_id": [369, 212, 219, 238, 35, 188],
        "name": ["Home Depot Canada", "PetSmart", "Comcast", "Office Depot", "Home Depot", "Charter Communications - Main - AQI 101"],
        "aqi_threshold": [150, 150, 150, 150, 150, 101],
        "aqi_states": [('BC'), ('CA','ID','NV','OR','WA'), ('CA'), ('CA'), ('CA'), ('CA','CO','HI','ID','MT','NV','OR','WA','WY')],
        "email_list": [
            ['369_PrimaryDistribution@earlyalert.com', '369_West_Region@earlyalert.com'], 
            ['PetSmartAQIDistribution@earlyalert.com'], 
            ['219_california_region@earlyalert.com'], 
            ['OfficeDepot_AQI_Distribution@earlyalert.com'], 
            ['HomeDepot_AQI_Distribution@earlyalert.com'], 
            ['188_aqidistribution@earlyalert.com']
        ]
    }

    process_one_client_dict = {
        "client_id": [client_dict["client_id"][0]],
        "name": [client_dict["name"][0]],
        "aqi_threshold": [client_dict["aqi_threshold"][0]],
        "aqi_states": [client_dict["aqi_states"][0]],
        "email_list": [client_dict["email_list"][0]]
    }

    process_two_client_dict = {
        "client_id": [client_dict["client_id"][1]],
        "name": [client_dict["name"][1]],
        "aqi_threshold": [client_dict["aqi_threshold"][1]],
        "aqi_states": [client_dict["aqi_states"][1]],
        "email_list": [client_dict["email_list"][1]]
    }

    process_three_client_dict = {
        "client_id": [client_dict["client_id"][2]],
        "name": [client_dict["name"][2]],
        "aqi_threshold": [client_dict["aqi_threshold"][2]],
        "aqi_states": [client_dict["aqi_states"][2]],
        "email_list": [client_dict["email_list"][2]]
    }

    process_four_client_dict = {
        "client_id": [client_dict["client_id"][3]],
        "name": [client_dict["name"][3]],
        "aqi_threshold": [client_dict["aqi_threshold"][3]],
        "aqi_states": [client_dict["aqi_states"][3]],
        "email_list": [client_dict["email_list"][3]]
    }

    process_five_client_dict = {
        "client_id": [client_dict["client_id"][4]],
        "name": [client_dict["name"][4]],
        "aqi_threshold": [client_dict["aqi_threshold"][4]],
        "aqi_states": [client_dict["aqi_states"][4]],
        "email_list": [client_dict["email_list"][4]]
    }

#     # get_each_client(client_dict)
    process_six_client_dict = {
        "client_id": [client_dict["client_id"][5]],
        "name": [client_dict["name"][5]],
        "aqi_threshold": [client_dict["aqi_threshold"][5]],
        "aqi_states": [client_dict["aqi_states"][5]],
        "email_list": [client_dict["email_list"][5]]
    }


    # get_each_client(client_dict)

    # process_one = multiprocessing.Process(target=get_each_client, args=(process_one_client_dict,))
    # process_two = multiprocessing.Process(target=get_each_client, args=(process_two_client_dict,))
    process_three = multiprocessing.Process(target=get_each_client, args=(process_three_client_dict,))
    # process_four = multiprocessing.Process(target=get_each_client, args=(process_four_client_dict,))
    # process_five = multiprocessing.Process(target=get_each_client, args=(process_five_client_dict,))
    process_six = multiprocessing.Process(target=get_each_client, args=(process_six_client_dict,))

    # process_one.start()
    # process_two.start()
    process_three.start()
    # process_four.start()
    # process_five.start()
    process_six.start()

    # process_one.join()
    # process_two.join()
    process_three.join()
    # process_four.join()
    # process_five.join()
    process_six.join()


temp()