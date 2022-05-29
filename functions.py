import tkinter as tk
from tkinter import filedialog
import base64
import pandas as pd
import io
import boto3
import json
import uuid
import re
import os, fnmatch
import win32com.client as win32
from datetime import datetime
import pytz 
import datetime
import requests
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from shutil import copyfile
from datetime import datetime
from zipfile import ZipFile
import json
import subprocess
import shutil
import time
import os, glob
import boto3
import requests
import base64
import uuid
import datetime

def file_load(new_file_path, customer_list):
    df = pd.read_excel(new_file_path)
    resp = excel_to_json(df, customer_list)
    body = resp['data']
    json_parsed = json.loads(body)
    return json_parsed

def excel_to_json(df, df2):
    df = df[["Facturas ", "C", "A", "Guia", "Cliente","Fecha Guia", "Fecha Signia", "Fecha Entrega CyC", "Hora", "Cargo", "Adjunto", "Ubicacion Carpeta/Sobre", "SO", "CODIGO"]]
    df = df.rename(columns={"SO":"OC", "Facturas ":"FACTURA", "Fecha Signia":"Fecha signia","Fecha Entrega CyC":"Fecha Entrega", "Ubicacion Carpeta/Sobre":"CLAS", "CODIGO":"Codigo"})    
    df2 = df2[['Codigo', 'CONTACTO', 'attachment_condition']]
    df_join = pd.merge(df, df2, on="Codigo", how="left")
    len_total_rows = len(df_join)
    doc_filters = ["pdfxmlocgr", "pdfxml"]
    len_with_email = len(df_join[df_join.attachment_condition.isin(doc_filters)]["CONTACTO"])
    df_join['total_count'] = len_total_rows
    df_join['target_customer'] = len_with_email
    df_join['Fecha Guia'] = df_join['Fecha Guia'].dt.strftime('%Y-%m-%d')
    df_join['Fecha signia'] = df_join['Fecha signia'].dt.strftime('%Y-%m-%d')
    df_join['Fecha Entrega'] = df_join['Fecha Entrega'].dt.strftime('%Y-%m-%d')
    df_join.fillna('', inplace=True)
    json_parsed = json.loads(df_join.to_json(orient='records'), parse_int=str)
    uuid_str = str(uuid.uuid4())    
    for i in json_parsed:
        i['uuid'] = uuid_str  
    json_data =  json.dumps(json_parsed, ensure_ascii=False).replace(u'\xa0', u' ')
    return {'response': "Success", 
            'data' : json_data,
            'uuid' : uuid_str    
           }

def enable_download_headless(browser,downloads_path):
    browser.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': downloads_path}}
    browser.execute("send_command", params)

    
def wait_for_downloads(downloads_path):
    print("Waiting for downloads", end="\n")
    while any([filename.endswith(".crdownload") for filename in 
              os.listdir(downloads_path)]):
        time.sleep(1)
        print(".", end="")

def run_scraping(downloads_path, driver_path, user, password, json_parsed):
    export_zip_path = downloads_path + "\\exportar.zip"
    with ZipFile(export_zip_path, 'w') as file:
        pass    
    dict_contacts=[]
    for i in json_parsed:
        if i['CONTACTO'] != '':
            dict_contacts.append(i)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--start-maximized')
    options.add_argument('--start-fullscreen')
    options.add_argument("disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_experimental_option("prefs", {
        "download.default_directory": downloads_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False
    })
    result = None
    while result is None:
        try:
            driver = webdriver.Chrome(driver_path, options=options)
            driver.get("https://www.businessmail.net/index.htm?v=4")
            p = driver.current_window_handle
            driver.find_element_by_id("usuario").send_keys(user)
            driver.find_element_by_id("password").send_keys(password)
            m = driver.find_elements_by_xpath("//button[@class='btn btn-primary btn-lg btn-block']")[0]
            driver.execute_script("arguments[0].scrollIntoView();", m)
            ActionChains(driver).move_to_element(m).click(m).perform()
            time.sleep(60)
            chwnd = driver.window_handles
#             print(chwnd)
            for w in chwnd:
                if(w!=p):
                    driver.switch_to.window(w)
                    search_window = w 
                    break
#             print("Child window title: " + driver.title)

            enable_download_headless(driver, downloads_path)

            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '(//tr[contains(@title, "Volumen")]//span[contains(@class, "standartTreeRow")])'))).click()
            a = ActionChains(driver)

            for lst in dict_contacts:
                driver.switch_to.window(search_window)
                time.sleep(2)
                m = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "(//div[@title='Búsqueda de documentos'])")))
                a.move_to_element(m).click(m).perform()
                iframe = driver.find_element_by_xpath("//*[contains(@src, '/panelControl/FormBuscar.jsp?_sessionHash')]")
                driver.switch_to.frame(iframe)        
                driver.find_element_by_xpath(("//tr[@id='trValor1']//*[contains(@id, 'valor1Buscar')]")).clear()
                driver.find_element_by_xpath(("//tr[@id='trValor1']//*[contains(@id, 'valor1Buscar')]")).send_keys(lst['FACTURA'])
                driver.find_element_by_xpath(("//button[@id='btnAceptar']")).click()    
                driver.switch_to.default_content()
                driver.switch_to.window(search_window)
                row = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='objbox']//*[contains(@class, 'ev_dhx_web')]")))
                actionChains = ActionChains(driver)
                actionChains.context_click(row).perform()
                row1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//tr[contains(@id, '_30')]//td[2]//div[@class='sub_item_text']")))
                actionChains1 = ActionChains(driver)
                actionChains1.move_to_element(row1).click(row1).perform()
                row3 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//tr[contains(@id, '_36')]//td[2]//div[@class='sub_item_text']")))
                actionChains = ActionChains(driver)
                actionChains.move_to_element(row3).click(row3).perform()
                wait_for_downloads(downloads_path)   
                with ZipFile(export_zip_path, 'r') as zip_ref:
                    zip_ref.extractall(downloads_path)
                time.sleep(2)
                dict_contacts.pop()
            driver.close();
            driver.quit();
            print("Successfully collected the files")
            result = True
        except:
            driver.close();
            driver.quit();
            pass


def send_email(email_recipients, liquidaciones_signia_path, pedidos_path, downloads_path, json_parsed):
    def find(pattern, path):
        result = []
        for root, dirs, files in os.walk(path):
            for name in files:
                if fnmatch.fnmatch(name, pattern):
                    result.append(os.path.join(root, name))
        return result
    guia_fileLocations = []
    oc_fileLocations = []
    factura_fileLocations = []
    consolidated_file = []

    for itm in json_parsed:
        observation_guia = itm['Guia']
        observation_oc = itm['OC']
        observation_factura = ""

        if len(itm['FACTURA'])==13:
            observation_factura = itm['FACTURA'][7:]
        else:
            observation_factura = itm['FACTURA']


        guia_fileLocations = find(observation_guia + "*.pdf", liquidaciones_signia_path)
        oc_fileLocations =find(observation_oc + "*.pdf", pedidos_path)
        factura_fileLocations = find("*" + observation_factura + "*[.PDF,.xml]", downloads_path)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = 'Cobranzas.MPSA@merckgroup.com'
        mail.To = itm['CONTACTO']  
        mail.Subject = "ENVIO DE DOCUMENTO/" + itm['Cliente']
        mail.Body = 'Message body'
        message_body_1 = """<h2><style>
                    <!--
                     /* Font Definitions */
                     @font-face
                        {font-family:"Cambria Math";
                        panose-1:2 4 5 3 5 4 6 3 2 4;}
                    @font-face
                        {font-family:Calibri;
                        panose-1:2 15 5 2 2 2 4 3 2 4;}
                     /* Style Definitions */
                     p.MsoNormal, li.MsoNormal, div.MsoNormal
                        {margin-top:0cm;
                        margin-right:0cm;
                        margin-bottom:8.0pt;
                        margin-left:0cm;
                        line-height:107%;
                        font-size:11.0pt;
                        font-family:"Calibri",sans-serif;}
                    a:link, span.MsoHyperlink
                        {color:#0563C1;
                        text-decoration:underline;}
                    p
                        {margin-right:0cm;
                        margin-left:0cm;
                        font-size:11.0pt;
                        font-family:"Calibri",sans-serif;}
                    .MsoChpDefault
                        {font-family:"Calibri",sans-serif;}
                    .MsoPapDefault
                        {margin-bottom:8.0pt;
                        line-height:107%;}
                     /* Page Definitions */
                     @page WordSection1
                        {size:595.3pt 841.9pt;
                        margin:72.0pt 72.0pt 72.0pt 72.0pt;}
                    div.WordSection1
                        {page:WordSection1;}
                    -->
                    </style>

                    </head>

                    <body lang=EN-IN link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>

                    <div class=WordSection1>

                    <p style='margin-bottom:12.0pt'><span style='color:#E25041'>"""

        message_body_2 = """</span><br>
                    <br>
                    Estimados,<br>
                    <br>
                    Se adjunta documentos de las facturas:<br>
                    <br>
                    <strong><span style='font-family:"Calibri",sans-serif'>"""

        message_body_3 = itm['FACTURA']

        message_body_4 = """</span></strong><br>
                    <br>
                    <strong><span style='font-family:"Calibri",sans-serif;color:#E25041'>Call
                    center: +51 1 618 7500 Life Science opcion 2 y luego Cobranzas opción 4</span></strong><b><span
                    style='color:#E25041'><br>
                    </span></b><br>
                    Best regards / Saludos Cordiales,<br>
                    Cash Collection Life Science<br>
                    <br>
                    <strong><span style='font-size:12.0pt;font-family:"Calibri",sans-serif'>Merck</span></strong><br>
                    <strong><span style='font-family:"Calibri",sans-serif;color:#EB6B56'>Merck
                    Peruana S.A.</span></strong><strong><span style='font-family:"Calibri",sans-serif;
                    color:#54ACD2'> | Av. Manuel Olguín 325, Of. 1702. Santiago de Surco. Lima |
                    Perú</span></strong></p>

                    <p class=MsoNormal>&nbsp;</p>

                    </div>"""

        mail.HTMLBody = message_body_1 + message_body_2 + message_body_3 + message_body_4 


        if itm['attachment_condition'] == "pdfxmlocgr":
            if guia_fileLocations and oc_fileLocations and factura_fileLocations:    
                itm['ENVIO AL CLIENTE'] = "FC " + datetime.datetime.today().strftime('%d/%m/%Y')
                itm['OBSERVACION'] = "Email Sent"
                itm['Responsable'] = "Bot" 

                attachement = guia_fileLocations + oc_fileLocations + factura_fileLocations
                for attach in attachement:
                    mail.Attachments.Add(attach)
                mail.Send()      
            else:
                itm['ENVIO AL CLIENTE'] = ""
                itm['OBSERVACION'] = "Cliente falta Guia/OC/Factura en Portal"
                itm['Responsable'] = "Bot" 


        elif itm['attachment_condition'] == "pdfxml":
            if factura_fileLocations:    
                itm['ENVIO AL CLIENTE'] = "FC " + datetime.datetime.today().strftime('%d/%m/%Y')
                itm['OBSERVACION'] = "Email Sent"
                itm['Responsable'] = "Bot" 

                attachement = factura_fileLocations
                for attach in attachement:
                    mail.Attachments.Add(attach)
                mail.Send()           
            else:
                itm['ENVIO AL CLIENTE'] = ""
                itm['OBSERVACION'] = "Cliente falta Factura en Portal"
                itm['Responsable'] = "Bot" 


        else:
            itm['ENVIO AL CLIENTE'] = ""
            itm['OBSERVACION'] = "Not a 4 doc. customer"
            itm['Responsable'] = "Collector"        

    #     print(itm)
        consolidated_file.append(itm)
        
    df_consolidated = df = pd.json_normalize(consolidated_file)
    df_consolidated = df_consolidated.drop(['attachment_condition', 'total_count', 'target_customer', 'uuid', 'CONTACTO'], axis = 1)

    current_time = datetime.datetime.now(pytz.timezone('Asia/Kolkata')) 
    date_string = current_time.strftime("%d-%m-%Y")
    date_time_string = current_time.strftime("%d-%m-%Y %H-%M-%S")
    filename = "Consolidados " + date_time_string
    filepath = "C:\\Users\\mravi\\Documents\\LATAM\\temp\\{}.xlsx".format(filename)
    writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
    df_consolidated.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()    

    mail = outlook.CreateItem(0)
    mail.To = email_recipients
    mail.SentOnBehalfOfName = 'Cobranzas.MPSA@merckgroup.com'
    mail.Subject = "Consolidados " + date_string
    mail.Body = 'Message body'
    mail.HTMLBody = "Hello, Bot has successfully run the automation process. Please refer to the attachment"
    attachement = factura_fileLocations
    mail.Attachments.Add(filepath)
    mail.Send()  

    email_count = df_consolidated[(df_consolidated['OBSERVACION'])=="Email Sent"].shape[0]
    summary = { "Date" : date_string, 
                "total_count":consolidated_file[0]['total_count'],
                "target_customer": consolidated_file[0]['target_customer'],
                "email_sent": email_count,
                "missed_invoices" : int(consolidated_file[0]['target_customer']) - int(email_count),
                "uuid" : consolidated_file[0]['uuid'],
                "Status": "Success"}

    headers = {'Content-Type': 'application/json'} 
    r = requests.post('https://prod-127.westeurope.logic.azure.com:443/workflows/823d64a4551540c8b0c5e07f8350fd7d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QgtzgNppW86N3Hy5nMsyEK1U4kaVG27Uqjz2h6dH2AI', json=summary,headers=headers)
    print(f"Status Code: {r.status_code}")
    print(f"Emails sent Successfully")