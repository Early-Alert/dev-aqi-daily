import requests, logging, xlsxwriter, os
from datetime import datetime, timedelta, timezone
import multiprocessing
import azure.functions as func

GIS_USERNAME = "Developer1"
GIS_PASSWORD = "qweRTY77**"
AERIS_ID = 'IJwChrr7utrp6BFWVnJ1A'
AERIS_SECRET = 'WmzZHc1PnJiAsr996tFNwfsPt7IqetRHHuB8Ldty'

def generate_services_arcgis_online_token():
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

def generate_services_rest_portal_token():
    url = "https://maps.earlyalert.com/portal/sharing/rest/generateToken"
    params = {
            "f": "json",
            "username": 'Developer1',
            "password": 'qweRTY77**',
            "client": "referer",
            "referer": "https://maps.earlyalert.com/portal/home/"
        }
    r = requests.post(url, data=params)
    token = r.json()["token"]
    return token

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

def add_static_header(formats, workbook, name, aqi_threshold, fireweather, redflag):

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
    
    ea_logo = os.path.join("eapytools", "static", "eapytools", "images", "ealogo white.png")

    """ TODAY AQI """
    date = datetime.now()
    workbook_sheet_name = f"{datetime.now():%Y-%m-%d}"
    worksheet1 = workbook.add_worksheet(workbook_sheet_name)
    worksheet1.set_row(0, 57)
    worksheet1.set_column(0, 0, 40)
    worksheet1.set_column(1, 1, 25)
    worksheet1.set_column(2, 2, 20)
    worksheet1.set_column(3, 4, 15)
    worksheet1.merge_range(
        f"A1:{last_column}1",
        f"{name} Locations with PM 2.5 AQI > {aqi_threshold}",
        title_format,
    )
    worksheet1.merge_range(f"A2:{last_column}2", f"Valid: {date:%a %b %d, %Y}", date_format)
    worksheet1.write(2, 0, "Name", col_header_format)
    worksheet1.write(2, 1, "Code", col_header_format)
    worksheet1.write(2, 2, "City", col_header_format)
    worksheet1.write(2, 3, "State", col_header_format)
    worksheet1.write(2, 4, "Zipcode", col_header_format)
    worksheet1.write(2, 5, "AQI", col_header_format)
    
    """ FORECAST AQI """
    date_tomorrow = datetime.now() + timedelta(days=1)
    # date2 = date.strftime('%Y-%m-%d') 
    worksheet2 = workbook.add_worksheet(f"{date_tomorrow.strftime('%Y-%m-%d') }")
    worksheet2.set_row(0, 57)
    worksheet2.set_column(0, 0, 40)
    worksheet2.set_column(1, 1, 25)
    worksheet2.set_column(2, 2, 20)
    worksheet2.set_column(3, 4, 15)
    worksheet2.merge_range(
        f"A1:{last_column}1",
        f"{name} Locations with PM 2.5 AQI > {aqi_threshold}",
        title_format,
    )
    worksheet2.merge_range(f"A2:{last_column}2", f"Valid: {date_tomorrow:%a %b %d, %Y}",date_format)
    worksheet2.write(2, 0, "Name", col_header_format)
    worksheet2.write(2, 1, "Code", col_header_format)
    worksheet2.write(2, 2, "City", col_header_format)
    worksheet2.write(2, 3, "State", col_header_format)
    worksheet2.write(2, 4, "Zipcode", col_header_format)
    worksheet2.write(2, 5, "AQI", col_header_format)

    if fireweather == "T":
        """ FIRE WEATHER """
        fire_weather_worksheet = workbook.add_worksheet("Fire Weather Watch")
        fire_weather_worksheet.set_row(0, 57)
        fire_weather_worksheet.set_column(0, 0, 40)
        fire_weather_worksheet.set_column(1, 1, 25)
        fire_weather_worksheet.set_column(2, 2, 20)
        fire_weather_worksheet.set_column(3, 3, 15)
        fire_weather_worksheet.insert_image(
                    "D1",
                    ea_logo,
                    {"x_scale": 0.8, "y_scale": 0.8, "x_offset": -65, "y_offset": 11},
                )
        fire_weather_worksheet.merge_range(
                    "A1:C1",
                    f"{name} Locations in Fire Weather Watches",
                    title_format,
                )
        fire_weather_worksheet.write(0, 3, "", title_format)
        fire_weather_worksheet.merge_range(
                    "A2:D2", f"Issued at {date:%I:%M %p %Z: %a %b %d, %Y}", date_format
                )
        fire_weather_worksheet.write(2, 0, "Name", col_header_format)
        fire_weather_worksheet.write(2, 1, "Code", col_header_format)
        fire_weather_worksheet.write(2, 2, "City", col_header_format)
        fire_weather_worksheet.write(2, 3, "State", col_header_format)
    
    if redflag == "T":
        """ RED FLAG """
        red_flag_worksheet = workbook.add_worksheet("Red Flag Warnings")
        red_flag_worksheet.set_row(0, 57)
        red_flag_worksheet.set_column(0, 0, 40)
        red_flag_worksheet.set_column(1, 1, 25)
        red_flag_worksheet.set_column(2, 2, 20)
        red_flag_worksheet.set_column(3, 3, 15)
        red_flag_worksheet.insert_image(
                    "D1",
                    ea_logo,
                    {"x_scale": 0.8, "y_scale": 0.8, "x_offset": -65, "y_offset": 11},
                )
        red_flag_worksheet.merge_range(
                    "A1:C1", f"{name} Locations in Red Flag Warnings", title_format
                )
        red_flag_worksheet.write(0, 3, "", title_format)
        red_flag_worksheet.merge_range(
                    "A2:D2", f"Valid as of {date:%I:%M %p %Z: %a %b %d, %Y}", date_format
                )
        red_flag_worksheet.write(2, 0, "Name", col_header_format)
        red_flag_worksheet.write(2, 1, "Code", col_header_format)
        red_flag_worksheet.write(2, 2, "City", col_header_format)
        red_flag_worksheet.write(2, 3, "State", col_header_format)

    if fireweather == "T" and redflag == "T":
        return worksheet1, worksheet2, fire_weather_worksheet, red_flag_worksheet
    else:
        return worksheet1, worksheet2



def send_aqi_email(clientId, filename, email_list, workbook_filename):
        """Sends daily email with spreadsheet to clients."""
        sub = email_list
        # email_list = ['yasir.khan@cooperativecomputing.com']
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


def send_logs(clientId, filename, email_list, workbook_filename):
        """Sends daily email with spreadsheet to clients."""
        sub = email_list
        email_list = ['yasir.khan@cooperativecomputing.com']
        # email_list = ['daldret@earlyalert.com']
        
        send_email = requests.post("https://api.mailgun.net/v3/earlyalert.com/messages",
        auth=("api", "key-abc3ac7030c2113b91c27b6733ebe510"),
            files = [("attachment", (f"{filename}", open(f"{filename}", "rb").read()))],
            data={"from": "airquality@earlyalert.com",
                "to": email_list,
                "bcc" : ["ykhan@earlyalert.com"],
                "subject": f"Morning AQI Logs.",
                "text": "Following is the log file for morning AQI.",
                # "template": template_name,
                # "h:X-Mailgun-Variables": dynamic,
                # "o:tag": tags_lists
                })
        print("********** Email Sent **********", filename, send_email.text)
        return send_email


def get_clients_from_portal(client_id):
    aq_cid_list = []
    url = 'https://services8.arcgis.com/lrWk3ELQFeb23nh1/arcgis/rest/services/service_d8f8bb17c9274da9940c1daaeedb833a/FeatureServer/0/query?'
    token = generate_services_arcgis_online_token()
    params = {
        'where': f"aqithreshold is not null AND (aqiproductchoices = 'daily' OR aqiproductchoices = 'both') AND cid in {client_id}",
        'geometryType': 'esriGeometryEnvelope',
        'spatialRel': 'esriSpatialRelIntersects',
        'relationParam': '',
        'outFields': 'cid, company, aqithreshold, aqistates, redflag, fireweather, aqi_email',
        'returnGeometry': 'false',
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

def parse_portal_clients():
    client_list = get_clients_from_portal((369, 212, 219, 238))
    # client_list = get_clients_from_portal((369, 212, 219, 238, 35, 188))

    for each_client in client_list:
        client_id = each_client.get("cid")
        name = each_client.get("company") 
        aqi_threshold = each_client.get("aqithreshold")

        aqistates = each_client.get("aqistates")
        aqi_states = ','.join("'{}'".format(word) for word in aqistates.split(',')) 
        
        aqi_email = each_client.get('aqi_email')
        email_list = aqi_email.split(',')

        fire_weather = each_client.get("fireweather")
        red_flag = each_client.get("redflag") 
        logging.info(f"{client_id}, {name}, {aqi_threshold}, {aqi_states}, {fire_weather}, {red_flag}, {email_list}")

        create_report(client_id, name, aqi_threshold, aqi_states, red_flag, fire_weather, email_list)

def get_client_locations(client_id, aqi_states):
    location_list = []
    locations_url = "https://maps.earlyalert.com/server/rest/services/Hosted/MASTER_CLIENT_LOCATIONS/FeatureServer/0/query?"
    token = generate_services_rest_portal_token()
    where_query = f"clientId = {client_id} AND state in ({aqi_states})"
    query = {
        "where": where_query,
        'geometryType': 'esriGeometryEnvelope',
        'spatialRel': 'esriSpatialRelIntersects',
        'relationParam': '',
        'outFields': 'clientid, name, code, city, state, lat, lon',
        'returnGeometry': 'false',
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
        "token": token,
    }
    r = requests.get(locations_url, params=query)
    response = r.json()
    # response = output_data.get('features')
    return response

def get_batch_aqi(lat, lon, fire_weather, red_flag):
    if fire_weather == "T" and red_flag == "T":
        endpoint_url = f"https://api.aerisapi.com/batch/{lat},{lon}?requests=/airquality,/airquality/forecasts/,/alerts?query=type:FW.A;type:FW.W"
    else:
        endpoint_url = f"https://api.aerisapi.com/batch/{lat},{lon}?requests=/airquality,/airquality/forecasts/"
    
    max_retries = 50
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
            print("**** Error Occued ****")
            pass  # Connection error occurred, retry the request
        
        retries += 1
    return None

def create_report(client_id, name, aqi_threshold, aqi_states, red_flag, fire_weather, aqi_email):
    formats, workbook, workbook_filename = formats_workbook(name)

    locations = get_client_locations(client_id, aqi_states)
    print(f"Locations {name}", len(locations["features"]))

    worksheet1, workbook, log_string = prepare_report_files(formats, workbook, client_id, name, aqi_threshold, aqi_states, fire_weather, red_flag, locations)
    workbook.close()

    # file = os.path.join("/tmp", "logs.txt")
    # with open(file, "w") as file:
    #     file.write(log_string)
    
    logging.info(f"Workbook Close {name}")
    send_aqi_email(client_id, workbook.filename, aqi_email, workbook_filename)
    logging.info(f"Email Sent {name}")
    # send_logs(client_id, "logs.txt", aqi_email, workbook_filename)


def prepare_report_files(formats, workbook, client_id, name, aqi_threshold, aqi_states, fireweather, redflag, locations):
    logging.info("prepare_report_files")

    workbook1_bool_populate = False
    workbook2_bool_populate = False
    fire_weather_bool_populate = False
    red_flag_bool_populate = False
    if fireweather == "T" and redflag == "T":
        worksheet1, worksheet2, fire_weather_worksheet, red_flag_worksheet = add_static_header(formats, workbook, name, aqi_threshold, fireweather, redflag)
    else:
        worksheet1, worksheet2 = add_static_header(formats, workbook, name, aqi_threshold, fireweather, redflag)
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

    workbook1_row = 2
    workbook2_row = 2
    fire_weather_worksheet_row = 2
    red_flag_worksheet_row = 2


    print(f"---> {len(locations['features'])}")
    log_string = ""
    for i in range(len(locations['features'])):
        
        lat = locations["features"][i]["attributes"]["lat"]
        lon = locations["features"][i]["attributes"]["lon"]
        if isinstance(lat, float) and isinstance(lon, float):
            aqi_response = get_batch_aqi(lat, lon, fireweather, redflag)
            # aqi_response = get_batch_aqi(44.513332, -88.015831, "T", "T")
            
            if aqi_response["success"]:
                if aqi_response["response"]["responses"][0]["success"]:
                    response_today_aqi = aqi_response["response"]["responses"][0]["response"][0]["periods"][0]["pollutants"]
                    today_aqi = next((p for p in response_today_aqi if p["type"] == "pm2.5"), None)["aqi"]
                if aqi_response["response"]["responses"][1]["success"]:
                    responst_forecast_aqi = aqi_response["response"]["responses"][1]["response"][0]["periods"][0]["pollutants"]
                    forecast_aqi = next((p for p in responst_forecast_aqi if p["type"] == "pm2.5"), None)["aqi"]
                
                print(f"{i} -- Today -- {today_aqi} & Forecast -- {forecast_aqi}")
                # log_string += f"Record-Number -- {i} -- Today -- {today_aqi} & Forecast -- {forecast_aqi} && Client -- {name}\n"
                
                """ TODAY AQI """
                if today_aqi and today_aqi > aqi_threshold:
                # if today_aqi and today_aqi > 10:
                    print(f"Today: {i} & AQI: {today_aqi}")
                    workbook1_row += 1
                    workbook1_bool_populate = True
                    json_output = locations["features"][i]
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
                    worksheet1.write(workbook1_row, 0, locations["features"][i]["attributes"].get('name'), row_format)
                    worksheet1.write(workbook1_row, 1, locations["features"][i]["attributes"].get('code'), row_format)
                    worksheet1.write(workbook1_row, 2, locations["features"][i]["attributes"].get('city'), row_format)
                    worksheet1.write(workbook1_row, 3, locations["features"][i]["attributes"].get('state'), row_format)
                    worksheet1.write(workbook1_row, 4, locations["features"][i]["attributes"].get('Zipcode'), row_format)
                    worksheet1.write(workbook1_row, 5, today_aqi, row_format)
                else:
                    pass
                """ FORECAST AQI """
                if forecast_aqi and forecast_aqi > aqi_threshold:
                # if forecast_aqi and forecast_aqi > 10:
                    print(f"Forecast: {i} & AQI: {today_aqi}")
                    workbook2_row += 1
                    workbook2_bool_populate = True
                    json_output = locations["features"][i]
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
                    worksheet2.write(workbook2_row, 0, locations["features"][i]["attributes"].get('name'), row_format)
                    worksheet2.write(workbook2_row, 1, locations["features"][i]["attributes"].get('code'), row_format)
                    worksheet2.write(workbook2_row, 2, locations["features"][i]["attributes"].get('city'), row_format)
                    worksheet2.write(workbook2_row, 3, locations["features"][i]["attributes"].get('state'), row_format)
                    worksheet2.write(workbook2_row, 4, locations["features"][i]["attributes"].get('Zipcode'), row_format)
                    worksheet2.write(workbook2_row, 5, forecast_aqi, row_format)
                else:
                    pass
                """ FIRE WEATHER """
                if fireweather == "T":
                    print(f'FIRE WEATHER {aqi_response["response"]["responses"][2]["success"]} ---- {len(aqi_response["response"]["responses"][2]["response"])}')
                    if aqi_response["response"]["responses"][2]["success"] and len(aqi_response["response"]["responses"][2]["response"]) > 0 and aqi_response["response"]["responses"][2]["response"][0]["type"] == "FW.A":
                        fire_weather_bool_populate = True
                        fire_weather_worksheet_row += 1
                        fire_weather_worksheet.write(fire_weather_worksheet_row, 0, locations["features"][i]["attributes"].get('name'), red_format)
                        fire_weather_worksheet.write(fire_weather_worksheet_row, 1, locations["features"][i]["attributes"].get('code'), red_format)
                        fire_weather_worksheet.write(fire_weather_worksheet_row, 2, locations["features"][i]["attributes"].get('city'), red_format)
                        fire_weather_worksheet.write(fire_weather_worksheet_row, 3, locations["features"][i]["attributes"].get('state'), red_format)
                else:
                    pass
                """ RED FLAG """
                if redflag == "T":
                    print(f'RED FLAG {aqi_response["response"]["responses"][2]["success"]} ---- {len(aqi_response["response"]["responses"][2]["response"])}')
                    if aqi_response["response"]["responses"][2]["success"] and len(aqi_response["response"]["responses"][2]["response"]) > 0 and aqi_response["response"]["responses"][2]["response"][1]["type"] == "FW.W":
                        red_flag_bool_populate = True
                        red_flag_worksheet.write(red_flag_worksheet_row, 0, locations["features"][i]["attributes"].get('name'), red_format)
                        red_flag_worksheet.write(red_flag_worksheet_row, 1, locations["features"][i]["attributes"].get('code'), red_format)
                        red_flag_worksheet.write(red_flag_worksheet_row, 2, locations["features"][i]["attributes"].get('city'), red_format)
                        red_flag_worksheet.write(red_flag_worksheet_row, 3, locations["features"][i]["attributes"].get('state'), red_format)

        else:
            print("not float")
    if workbook1_bool_populate == False:
        worksheet1.merge_range("A4:F4", "No locations exceeded the threshold.", no_loc_format)
    if workbook2_bool_populate == False:
        worksheet2.merge_range("A4:F4", "No locations forecast to exceed the threshold.", no_loc_format)
    if fire_weather_bool_populate == False and fireweather == "T":
        fire_weather_worksheet.merge_range("A4:D4", "No locations in Fire Weather Watch today.", no_loc_format)
    if red_flag_bool_populate == False and redflag == "T":
        red_flag_worksheet.merge_range("A4:D4", "No locations in Red Flag Warnings today.", no_loc_format)

    add_static_footer(formats, worksheet1, worksheet2, workbook1_row, workbook2_row)
    
    return worksheet1, workbook, log_string

# parse_portal_clients()

# get_client_locations(369, ('BC'))


def main(mytimer: func.TimerRequest) -> None:
    utc_timestamp = datetime.utcnow().replace(
        tzinfo=timezone.utc).isoformat()
    # get_each_client([369, 212, 219, 238, 35, 188])
    parse_portal_clients()

    logging.info('Python timer trigger function ran at %s', utc_timestamp)



# def amain():
#     utc_timestamp = datetime.utcnow().replace(
#         tzinfo=timezone.utc).isoformat()
#     parse_portal_clients()

#     logging.info('Python timer trigger function ran at %s', utc_timestamp)

# amain()
