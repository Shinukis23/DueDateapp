# Programa para calcular el due-date de las ordenes
# Modificado Agosto 23/2023 
import tkinter as tk
from tkinter import messagebox
from tkinter import *
import csv
from tkinter import ttk
from tkinter import simpledialog
from PIL import Image, ImageTk
import datetime #as dt
from datetime import datetime as dt
import time
from tkinter.constants import *
from tkcalendar import Calendar
#from Fun_Eficiencia import *
import pytz
import pandas as pd
import numpy as np
import xlwt
import openpyxl
import xlsxwriter
import gspread
import pygsheets
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import  timedelta, datetime
from Fun_PromesaCiente import *
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import gspread
import sys
import warnings
import threading
import gspread_dataframe as gd
import gdown
import io
import os
from urllib.parse import urlparse
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import requests
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)

file_id0 = "1F0L_aHVNNhGuV-KNnuT6nCr_X1Af3l3E" #cortes2023.xlxs
file_id1 = "15vHlzGFgi9MjxyclqmNArvheijJhLSK5" #tiempos.xls


scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]
credentials = ServiceAccountCredentials.from_json_keyfile_name("monitor-eficiencia-3a13458926a2.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials)# authenticate the JSON key with gspread
ss = file.open("EficienciaReporte")
#########333

#from oauth2client.service_account import ServiceAccountCredentials

# Credenciales
#credentials = ServiceAccountCredentials.from_json_keyfile_name("monitor-eficiencia-3a13458926a2.json", scopes)  # Reemplaza con tu ruta
#client = gspread.authorize(credentials)

####
###

service = build('drive', 'v3', credentials=credentials)

# Define the URL to download the file from
file_url = service.files().get(fileId=file_id0, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)

# Define the filename to save the downloaded file as
filename = f"cortes2023.xlsx"

# Download the file
try:
    request = service.files().get_media(fileId=file_id0)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")

# Define the URL to download the file from
file_url = service.files().get(fileId=file_id1, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)

# Define the filename to save the downloaded file as
filename = f"tiempos.xls"

# Download the file
try:
    request = service.files().get_media(fileId=file_id1)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")    

# Define the URL to download the file from


url0 ="https://drive.google.com/file/d/11FW_HPRLaR-h2bk9eMs5VDTQvJ8Sg3B4/view?usp=share_link" #festivos 2023
url = "https://drive.google.com/file/d/1CgwK7OCtiTXggOZpw3E39PgMlPEQo8WP/view?usp=share_link" #claves.csv
url1 = "https://drive.google.com/file/d/1zTGwpZEryABsYqAf0xGISu1rYZ9MNTsF/view?usp=sharing" # partes exportacion


#1Q8apVPagzMYeOfm4PI_-q0bCFqkPgYjNc2xSD_T12MY
#url1 = "https://drive.google.com/file/d/1d2dtM81uB4Ag8vS8lMruAIdEGJNxN44q/view?usp=sharing"
#https://drive.google.com/file/d/1d2dtM81uB4Ag8vS8lMruAIdEGJNxN44q/view?usp=sharing
output_path = 'claves.csv'
gdown.download(url, output_path, quiet=False,fuzzy=True)
output_path = 'festivos2023.csv'
gdown.download(url0, output_path, quiet=False,fuzzy=True)


with open(r'claves.csv', mode='r') as infile:
    reader = csv.reader(infile)
    with open('claves_new.csv', mode='w') as outfile:
        writer = csv.writer(outfile)
        users = {rows[0]:rows[1] for rows in reader}
#print(users)


USA = 0
MEXICO = 1
MEXICO2 = 2
USA2 = 3
dfFest = pd.read_csv(r'festivos2023.csv')
df = pd.read_excel(r'Tiempos.xls')
dc = pd.read_excel(r'Cortes2023.xlsx')

##################################s = file.open("EficienciaReporte")
df1 = pd.DataFrame({
    "User": '',
    "Order #": '',
    "Part Store #": '',
    "Route":'',
    "Created":'',
    "Due Date part":'',
    "Due Date Order":'',
    "DueDate change":'',
    "Reason":'',
}, index=["Dummy"])
#print(dfFest)
dfpartes = []
festivosusa= dfFest["fechaUSA"].tolist()
festivosmex = dfFest["fechaMEX"].tolist()
festivosusa1= dfFest["fechaUSAd"].tolist()
festivosmex1 = dfFest["fechaMEXd"].tolist()
output_path = 'PartesExportacion.csv'
df= df.set_index('Store')
#dc = dc.to_dict()
#print(dc)
dc = dict(dc.set_index('DIA').groupby(level = 0).\
    apply(lambda x : x.to_dict(orient= 'list')))
def ExportFile():
    global credentials
#    global url1
    global dfpartes
#    gdown.download(url1, output_path, quiet=False,fuzzy=True)
#    dfpartes = pd.read_csv(r'PartesExportacion.csv')
    #print(dfpartes)
     # Abrir la hoja de cálculo
    file = gspread.authorize(credentials) 
    ss1 = file.open("Listado de partes Importacion-Exportacion")

    # Obtener el primer archivo de la hoja de cálculo
    worksheet = ss1.get_worksheet(0)

    # Obtener todos los valores de la hoja de cálculo
    values = worksheet.get_all_values()

    # Crear un DataFrame de pandas
    # dfpartes = pd.DataFrame(values)

    # Guardar DataFrame en formato Excel
    
    # Crear un DataFrame de pandas
    dfpartes = pd.DataFrame(values[1:], columns=values[0])  # Ignorar la primera fila (encabezados) al crear DataFrame
    
    # Eliminar la primera columna
    dfpartes = dfpartes.drop(dfpartes.columns[0], axis=1)
    dfpartes['No Parte'] = dfpartes['No Parte'].astype(str)
    print(dfpartes.head()) 
    

    dfpartes.to_excel("archivo.xlsx", index=False)
    # Cambia el nombre si lo deseas
    print('Archivo descargado y guardado como archivo.xlsx')
    #print(dfpartes.loc[5, 'No Parte'])
current_time = datetime.now()
ExportFile()
#print("fuera de la funcion:")
#print(dfpartes.loc[5, 'No Parte'])

def Actualizacion():
    #global DOF
    #global resultWeekend
    #global diff
    #global result1
    #global result
    #global Tipo_Cambio
    while True:
        ExportFile()
        #tk.Label(main_window,font=("Courier 9 bold"),fg="red", text = DOF).place(x = 14,y = 40)
        time.sleep(5*60*60)

#Create a separate thread for printing the message
thread = threading.Thread(target=Actualizacion)
thread.daemon = True  # Set the thread as a daemon so it won't prevent the script from exiting
thread.start()

def main_program(super_user,hoja):

    def add_action():
        
        selected_radio = var.get()
        selected_combobox = combobox.get()
        if selected_radio in [1,2,3,6,7,8] and selected_combobox == "EBAY TJ":
            messagebox.showwarning("Aerta", "No se puede usar ruta EBAY TJ si la parte esta en USA ")
           # text_widget.configure(state='normal')
           # text_widget.insert(tk.END, "No se puede usar ruta EBAY TJ si la parte esta en USA \n") 
           # text_widget.see(tk.END)
           # text_widget.configure(state='disabled')
        elif selected_radio == 2 and selected_combobox == "WILL CALL 1":
            messagebox.showwarning("Alerta", "No se envia a WILL CALL1 desde Nirvana ")
            #text_widget.configure(state='normal')
            #text_widget.insert(tk.END, "No se envia a WILL CALL1 desde Nirvana \n") 
            #text_widget.see(tk.END)
            #text_widget.configure(state='disabled')
        elif selected_radio == 1 and selected_combobox == "WILL CALL 1":
            messagebox.showwarning("Alerta", "No esta activada produccion TAP1-WILL CALL1")
            #text_widget.configure(state='normal')
            #text_widget.insert(tk.END, "No esta activada produccion TAP1-WILL CALL1 \n") 
            #text_widget.see(tk.END)
            #text_widget.configure(state='disabled')             
        else:        
            pacifictime = datetime.now(pytz.timezone('US/pacific'))
            current_time = pacifictime.strftime("%Y-%m-%d %H:%M:%S")
            dia = pacifictime.weekday()    
            FECHA = pacifictime.date()
            TIEMPO = pacifictime.time()
            if dia != 6: 
                #data.append([selected_radio, selected_combobox, current_time])
                text_widget.configure(state='normal')
                text_widget.insert(tk.END, 'Orden :' + str(orden.get()) + '   ') 
                text_widget.insert(tk.END,   dt.strptime(current_time, "%Y-%m-%d %H:%M:%S").strftime("%a,%d %b, %Y") + '\n')
                text_widget.see(tk.END)
                text_widget.insert(tk.END, 'Store :' + str(selected_radio) + '   ')
                text_widget.see(tk.END)
                text_widget.insert(tk.END, 'Drop :' + str(selected_combobox)+ '\n' )
                text_widget.see(tk.END)
                
                text_widget.configure(state='disabled')

                ##print("Radio Button Selection:", selected_radio)
                ##print("Combobox Value:", selected_combobox)
                ##print("Combobox Value:", current_time)
                ##print("Current Day of the week:", pacifictime.weekday())
                ##print("Current Date :", pacifictime.date())
                ##print("Current time :", pacifictime.time())
                #dia = pacifictime.weekday()    
                #FECHA = pacifictime.date()
                #TIEMPO = pacifictime.time()
                #print(TIEMPO)
                #print(FECHA)
                #print(dia)
                #if dia == 6:  
                    
                    #print("Domingo no es dia laborable")
                #    messagebox.showinfo("showinfo", "Domingo no es dia loaborable")
                #    clearOrden()

                fechaProm = fechaPromesa(selected_radio,selected_combobox,dia,FECHA,TIEMPO)
                ### -> fechas.append(fechaProm) 
                ####print("busco",fechaProm.strftime('%m-%d-%Y'))
                #if fechaProm.strftime('%m-%d-%Y') in  festivosmex  :#or festivosusa :
                if selected_radio in [4,10,15] and selected_combobox in ["15 EAST", "15 NORTH 2", "15 SOUTH", "15 WEST","5 NORTH","5 NORTH 2","5 NORTH 3","5 EAST","5 WEST","AREA SD","SD AUX"
                         ,"SHOP SD","SHIPPING","WILL CALL 1","WILL CALL 6","WILL CALL 7","SHIP WC1", "SHIP WC6"]:
                    fechaProm = verificaFestivo(fechaProm,MEXICO)
                elif selected_radio in [1,2,3,6,7,8] and selected_combobox in ["ENSENADA","TIJUANA","PAQUETERIA TJ"]:         
                    fechaProm = verificaFestivo(fechaProm,USA)
                elif selected_radio in [4,10,15] and selected_combobox in ["ENSENADA","TIJUANA","PAQUETERIA TJ"]:         
                    fechaProm = verificaFestivo(fechaProm,MEXICO2)    
                elif selected_radio in [1,2,3,6,7,8] and selected_combobox in ["15 EAST", "15 NORTH 2", "15 SOUTH", "15 WEST","5 NORTH","5 NORTH 2","5 NORTH 3","5 EAST","5 WEST","AREA SD","SD AUX"
                         ,"SHOP SD","SHIPPING","WILL CALL 1","WILL CALL 6","WILL CALL 7","SHIP WC1", "SHIP WC6"]:         
                    fechaProm = verificaFestivo(fechaProm,USA2)        
               # if festivosmex.count(fechaProm.strftime('%m-%d-%Y')):
               #     print("si esta",fechaProm.strftime('%m-%d-%Y'))

                #   temporal = fechaProm
                #    fechaProm +=  timedelta(hours = 24) #pd.to_datetime(FECHA) +
                #   messagebox.showwarning("Dia Festivo en USA el  " + temporal.strftime("%a,%d %b, %Y"), "Fecha compromiso cambio por dia festivo a "  + fechaProm.strftime("%a,%d %b, %Y"))
                #elif fechaProm.strftime('%m-%d-%Y') in festivosmex :
                #   print(fechaProm.strftime('%m-%d-%Y'))
                #   temporal = fechaProm
                #   fechaProm +=  timedelta(hours = 24) #pd.to_datetime(FECHA) +
                #   messagebox.showwarning("DIA FESTIVO " + temporal.strftime("%a,%d %b, %Y"), "Fecha compromiso cambio por dia festivo a "  + fechaProm.strftime("%a,%d %b, %Y"))
                fechas.append(fechaProm) 
                #data.append(fechaProm)
                #data.append([selected_radio, selected_combobox, current_time, str(fechaProm)])
                ####### data.append([selected_radio])
                ####### data.append([selected_combobox])
                ####### data.append([current_time])
                ####### data.append([str(fechaProm)])
                data.extend([selected])
                data.extend([str(orden.get())])
                data.extend([current_time])
                data.extend([selected_radio])
                data.extend([selected_combobox])    
                data.extend([str(fechaProm)])
                df1['Part Store #'] = selected_radio
                df1['Route'] = selected_combobox
                df1['Created'] = current_time
                df1['Due Date part'] = str(fechaProm)
                text_widget.configure(state='normal')
                text_widget.insert(END, "Fecha llegada a ruta o locacion seleccionada:" + '\n'+ dt.strptime(str(fechaProm), '%Y-%m-%d' ).strftime("%a,%d %b, %Y") + '\n')
                text_widget.see(END)

                text_widget.configure(state='disabled')
                entry1.config(state= "disabled")
                subir2()
            else:
                messagebox.showinfo("showinfo", "Domingo no es dia loaborable")
                clearOrden()
                
    """def verificaFestivo(fechaProm,pais):

            if pais :
                if festivosmex.count(fechaProm.strftime('%m-%d-%Y')):
                 print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                #print("Fecha promesa de entrega es dia festivo en MEXICO !Consulta con Import/Export)")
                 messagebox.showerror("DIA FESTIVO", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es dia festivo en MEXICO !Consulta con Import/Export)")
            else: 
                if festivosusa.count(fechaProm.strftime('%m-%d-%Y')):
                 print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                #print("Fecha promesa de entrega es dia festivo en USA !Consulta con Import/Export)")
                 messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es dia festivo en USA !Consulta con Import/Export)")
            #   temporal = fechaProm

            #fechaProm +=  timedelta(hours = 24) #pd.to_datetime(FECHA) +
            return fechaProm  """
    def verificaFestivo(fechaProm,pais):
        
        if pais == 0 or pais ==1:             
                    if festivosusa.count(fechaProm.strftime('%m-%d-%Y')) or festivosusa1.count(fechaProm.strftime('%m-%d-%Y')):
                        #    if pais == 0 :
                        if festivosusa1.count(fechaProm.strftime('%m-%d-%Y')):
                        #if festivosmex.count(fechaProm.strftime('%m-%d-%Y')):
                                #print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                        #print("Fecha promesa de entrega es dia festivo en MEXICO !Consulta con Import/Export)")
                                messagebox.showerror("DIA FESTIVO", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es 1 dia despues del festivo en USA !Consulta con Import/Export!")
                        else :
                                #print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                        #print("Fecha promesa de entrega es dia festivo en MEXICO !Consulta con Import/Export)")
                                messagebox.showerror("DIA FESTIVO", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es dia festivo en USA !Consulta con Import/Export!")    
                                
                    if festivosmex.count(fechaProm.strftime('%m-%d-%Y')):
                    #else: 
                        if festivosmex1.count(fechaProm.strftime('%m-%d-%Y')):
                                #print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                        #print("Fecha promesa de entrega es dia festivo en USA !Consulta con Import/Export)")
                                messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es 1 dia despues del festivo en MEXICO !Consulta con Import/Export!")
                            #   temporal = fechaProm
                        else:
                                #print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                        #print("Fecha promesa de entrega es dia festivo en USA !Consulta con Import/Export)")
                                messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es dia festivo en MEXICO !Consulta con Import/Export)")
        '''if pais ==3:                         
                    if festivosusa.count(fechaProm.strftime('%m-%d-%Y')) or festivosusa1.count(fechaProm.strftime('%m-%d-%Y')):
                        #    if pais == 0 :
                        if festivosusa1.count(fechaProm.strftime('%m-%d-%Y')):
                        #if festivosmex.count(fechaProm.strftime('%m-%d-%Y')):
                                #print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                        #print("Fecha promesa de entrega es dia festivo en MEXICO !Consulta con Import/Export)")
                                messagebox.showerror("DIA FESTIVO", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es 1 dia despues del festivo en USA !Consulta con tu encargado!")
                        else :
                                #print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                        #print("Fecha promesa de entrega es dia festivo en MEXICO !Consulta con Import/Export)")
                                messagebox.showerror("DIA FESTIVO", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es dia festivo en USA !Consulta con tu encargado!")    '''
        if pais == 2 :                         
                    if festivosmex.count(fechaProm.strftime('%m-%d-%Y')):
                    #else: 
                        if festivosmex1.count(fechaProm.strftime('%m-%d-%Y')):
                                #print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                        #print("Fecha promesa de entrega es dia festivo en USA !Consulta con Import/Export)")
                                messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es 1 dia despues del festivo en MEXICO !Consulta con tu encargado!")
                            #   temporal = fechaProm
                        else:
                                #print(fechaProm.strftime('%m-%d-%Y')," es dia festivo")
                        #print("Fecha promesa de entrega es dia festivo en USA !Consulta con Import/Export)")
                                messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime('%m-%d-%Y') + " es dia festivo en MEXICO !Consulta con tu encargado!")        
        #fechaProm +=  timedelta(hours = 24) #pd.to_datetime(FECHA) +
        return fechaProm          

    def submit_action():
        nonlocal dates
        ##print("Submit button pressed.")
        if fechas : # si no hay partes en la orden no calcula la Fecha promesa
            ##print(fechas)
            dates = largest_date(fechas)
            #text_widget.configure(state='normal')
            #text_widget.insert(tk.END, "Fecha Compromiso para el cliente es :" + '\n'+ dt.strptime(dates, '%Y-%m-%d' ).strftime("%a,%d %b, %Y") + '\n') 
            #text_widget.configure(state='disabled')
            messagebox.showinfo("showinfo", "Fecha Compromiso para el cliente es :"+ dt.strptime(dates, '%Y-%m-%d').strftime("%a,%d %b, %Y"))
            #compromiso = Label(main_window,font=("Courier 6 bold"),fg="blue", text = "Fecha Compromiso para el cliente es :")
            #compromiso.place(x = 150,y = 60) 
            #compromiso = Label(main_window,font=("Courier 13 bold"),fg="red", text =  dt.strptime(dates, '%Y-%m-%d').strftime("%a,%d %b, %Y") )######
            #compromiso.place(x = 230,y = 200) 
            fechas.clear()
            add_button['state'] = DISABLED
            submit_button['state'] = DISABLED
            cambio_button['state'] = NORMAL
            enviar_button['state'] = DISABLED #NORMAL
            entry1.config(state= "disabled")
            clear_button['state'] = NORMAL
            data.extend(["-","-","-","-","-","-"])   
            data.extend([dates])
            subir()

    def clearOrden():
        #data.clear()
        #orden.set('')
        #fechas.clear()
        text_widget.configure(state='normal')
        text_widget.delete(1.0,END)
        text_widget.configure(state='disabled')
        entry1.config(state= "normal")
        add_button['state'] = DISABLED
        submit_button['state'] = DISABLED
        cambio_button['state'] = DISABLED
        enviar_button['state'] = DISABLED
        #part_button['state'] = DISABLED
        #compromiso = Label(main_window,font=("Courier 10 bold"), text = "                                                            ")
        #compromiso.place(x = 230,y = 180) 
        #compromiso = Label(main_window,font=("Courier 13 bold"),fg="red", text = "                                           " )
        #compromiso.place(x = 230,y = 200)

        #if data :
        data.clear()
        orden.set('')
        fechas.clear()
        data.extend(["-","-","-","-","-","-","-"])   
            #data.extend([dates])
        data.extend(["Clear"])
        subir2()
        #else  :
        #    print(data)
        #    data.clear()
        #    orden.set('')
        #    fechas.clear()

    def close_window():
        main_window.destroy()
        login()

    def validation(i,text, new_text): 
        #if len(new_text) == 0  or len(new_text) < 10 and text.isdecimal():
            return len(new_text) == 0  or len(new_text) < 10 and text.isdecimal()
        #else: 
        #    text_widget.configure(state='normal')
        #    text_widget.delete(1.0,END)
        #    text_widget.configure(state='disabled')  

    def form_complete(event):
        if len(orden.get()) <= 1 :
            add_button['state'] = DISABLED
            submit_button['state'] = DISABLED
            cambio_button['state'] = DISABLED
            #part_button['state'] = DISABLED
            text_widget.configure(state='normal')
            text_widget.delete(1.0,END)
            text_widget.configure(state='disabled')

        else:
            add_button['state'] = NORMAL
            submit_button['state'] = NORMAL
            part_button['state'] = NORMAL
    ##################    
    def seleccionar():
        #print("Super User")
        f = g


    def check_export():
        ExportFile()
        ###output_path = 'PartesExportacion.csv'
        ###gdown.download(url1, output_path, quiet=False,fuzzy=True)
        ###dfpartes = pd.read_csv(r'PartesExportacion.csv')
        #global output_path

        #dpartes= dfpartes.stack().tolist()
        #dpartes.sort()
        #print(dpartes[(dpartes['No Parte']== 204)]['Descripcion'])out
        export_window = tk.Tk()
        export_window.title("Listado de partes import/export")
        #export_window.geometry("1250x300")#1200x310]export_window.geometry("350x300")
        export_window.geometry("350x300")
        export_window.iconbitmap("logoicon.ico")
        export_window.resizable(False, False)
        export_window.attributes('-topmost',True)

        def search_df():
            #global dfpartes
            # Get the user input from the Entry widget
            #search_num = int(codigo_field.get())
            search_num = codigo_field.get()
            #print(dfpartes[0])
            #print(dfpartes.loc[5, 'No Parte'])
            # Filter the DataFrame to find all rows with matching age
            matches = dfpartes[(dfpartes['No Parte'] == str(search_num)) & (dfpartes['Destino'] == "Mexico-Estados Unidos")]
            matches2 = dfpartes[(dfpartes['No Parte'] == str(search_num)) & (dfpartes['Destino'] == "Estados Unidos-Mexico")]
            #print("matches",matches)
            #print("matches2",matches2)
            #print ("maximo",max(matches, key=len))
            if (len(matches)==0 and len(matches2)==0):
                export_window.geometry("350x300")
                match_label.configure(text="No cruza",font=("Courier 18 bold"), fg="red")
                match_label2.configure(text="",font=("Courier 13 bold"), fg="red")
                match_label3.configure(text="",font=("Courier 13 bold"), fg="red")
                match_label4.configure(text="",font=("Courier 13 bold"), fg="red")
            else:
                # Convert the search result to a string
                export_window.geometry("1250x300")
                search_result_str = "\n".join([f"{row['Destino']} - " for _, row in matches.iterrows()])
                search_result_str3 = "\n".join([f"{row['Destino']} - " for _, row in matches2.iterrows()])
                search_result_str2 = "\n".join([f" - {row['Descripcion']}" for _, row in matches.iterrows()])
                search_result_str4 = "\n".join([f" - {row['Descripcion']}" for _, row in matches2.iterrows()])
                # Set the text of the Label widget
                #print(len(search_result_str))
                #print(len(search_result_str3))
                #print(len(search_result_str2))
                #print(len(search_result_str4))
                #export_window.geometry("")
                #export_window.geometry(f"{300+(len(search_result_str4)*4)}x300")
                #print(f"{300+(len(search_result_str4)*4)}")
                #print(f"{(len(search_result_str4)*4)}")
                #print(len(search_result_str))
                
                
                match_label.configure(text=search_result_str,font=("Courier 8 bold"),fg="blue")
                match_label2.configure(text=search_result_str2,font=("Courier 8 bold"), fg="black")
                match_label3.configure(text=search_result_str3,font=("Courier 8 bold"),fg="green")
                match_label4.configure(text=search_result_str4,font=("Courier 8 bold"), fg="black")
                #export_window.geometry(f"{match_label4.winfo_width()+match_label.winfo_width()}x300")
                #print(match_label4)
                #print(match_label3)
                #measure = Label(frame, font = ("Purisa", 10), text = "The width of this in pixels is.....", bg = "yellow")
                #measure.grid(row = 0, column = 0) # put the label in
                if len(search_result_str3) == 0 and len(search_result_str) != 0 :
                    match_label2.update_idletasks() # this is VERY important, it makes python calculate the width
                    export_window.geometry(f"{match_label2.winfo_width()+194}x300")
                    #print('string3',len(search_result_str3))
                    #print('string',len(search_result_str))
                elif len(search_result_str3) != 0 and len(search_result_str) == 0 :
                    match_label4.update_idletasks() 
                    export_window.geometry(f"{match_label4.winfo_width()+194}x300") 
                    #print('string', len(search_result_str))  
                    #print('string3',len(search_result_str3))
                else : 
                    match_label4.update_idletasks() 
                    export_window.geometry(f"{match_label4.winfo_width()+194}x300") 
                    #print('else', len(search_result_str))  
                    #print('else',len(search_result_str3))    
                #width = match_label.winfo_width() # get the width
                #print('Tamano', match_label.winfo_width())
                
                #match_label4.configure(text=search_result_str4,font=("Courier 8 bold"), fg="black")

        instructions = tk.Label(export_window, text="Introduzca el codigo que esta buscando:")
        instructions.place(x=20,y=10)
        codigo_field = tk.Entry(export_window,width=5,state="normal")
        codigo_field.place(x=245,y=10)
        # Bind the Enter key to the search function

        codigo_field.bind("<Return>", lambda event: search_df())
        codigo_field.focus()
        # Create a Label widget to display the search results
        match_label = tk.Label(export_window, text="",justify=LEFT)
        match_label.place(x=30,y=60)
        match_label2 = tk.Label(export_window, text="",justify=LEFT)
        match_label2.place(x=180,y=60)
        match_label3 = tk.Label(export_window, text="",justify=LEFT)
        match_label3.place(x=30,y=150)
        match_label4 = tk.Label(export_window, text="",justify=LEFT)
        match_label4.place(x=180,y=150)
        export_window.mainloop()


    def submit():
        nonlocal cambio
        ##print("Upload to Cloud button pressed.")
        #date = Label(main_window, text = "                                                                   ")
        #date.place(x=200,y=450)
        #date.destroy()
        #data.append([str(orden.get()),dates,selected])
        ###########3 data.append([str(orden.get())])
        #########333 data.append([dates])
        #########3 data.append([selected])
        ####data.extend([str(orden.get())])
        #data lc
        if cambio == 1:
         data.extend(["-","-","-","-","-","-"])   
         data.extend([dates])
         data.extend("X")
         cambio = 0
        else :
         data.extend(["-","-","-","-","-","-"])   
         data.extend([dates])
        #data.extend("x")
        ######data.extend([selected])
        df1['Order #'] = str(orden.get())
        df1['Due Date Order'] = dates
        df1['User'] = selected
        ##print(data)
        df1.append(data)
        ##print(df1)
        cambio_button['state'] = DISABLED
        enviar_button['state'] = DISABLED
        text_widget.configure(state='normal')
        text_widget.delete(1.0,END)
        text_widget.configure(state='disabled')
        entry1.config(state= "normal")
        orden.set('')
        #compromiso = Label(main_window,font=("Courier 10 bold"), text = "                                                            ")
        #compromiso.place(x = 230,y = 180) 
        #compromiso = Label(main_window,font=("Courier 13 bold"),fg="red", text = "                                           "  )
        #compromiso.place(x = 230,y = 200)
        #df1 = df1.to_json()
        #df1.head().to_dict()
        subir()

    def promesa():
        #Asigno el valor de ruta
        #for i in range(len(ds2)) :
        St = selected_radio
        Rt = selected_combobox
        fecha1 = df.at[St,Rt] # Busco en en tabla de tiempos la combinacion de ruta y drop para saber que corte es
  
    def largest_date(dates):
        return max(dates).strftime('%Y-%m-%d') 


    def fechaPromesa(selected_radio,selected_combobox,dia,FECHA,TIEMPO): 
         fechaCompromiso = ''
         #print(TIEMPO)
         #print(FECHA)
         #print(dia)
         def tabla(tiempo,c,b):
          nonlocal fechaCompromiso
          #print(tiempo)
          #print(c)
          #print(b)  
          if TIEMPO < tiempo.time() :
              a=dc.get(dia)
              delt = a.get(c)        
              fechaCompromiso = pd.to_datetime(FECHA) + timedelta(hours = delt[0])
              #print("Fecha compromis:" + str(fechaCompromiso.date()))
          else :
              a=dc.get(dia)
              delt = a.get(b)      
              fechaCompromiso = pd.to_datetime(FECHA) + timedelta(hours = delt[0])
              #print("Fecha compromis:" + str(fechaCompromiso.date()))  
         def tabla1(tiempo,c,b):
          nonlocal fechaCompromiso 
          #print(tiempo)
          #print(c)
          #print(b)  
          if TIEMPO < tiempo.time() :
              fechaCompromiso = pd.to_datetime(FECHA) + timedelta(hours = 0) 
              #print("Fecha compromis:" + str(fechaCompromiso.date()))    
          else :      
              fechaCompromiso = pd.to_datetime(FECHA) + timedelta(hours = 0)
              #print("Fecha compromis:" + str(fechaCompromiso.date()))
             
         St = selected_radio
         Rt = selected_combobox
         fecha1 = df.at[St,Rt] 
         #print("fecha1: ", fecha1)  
         tiempo1 = datetime(2022,1,1,12,31,00) # asigno tiempos iniciales para comparar 12:31
         tiempo2 = datetime(2022,1,1,13,1,00) # asigno tiempos iniciales para comparar 13:01
         tiempo3 = datetime(2022,1,1,14,1,00) # asigno tiempos iniciales para comparar 14:01
         tiempo4 = datetime(2022,1,1,16,1,00) # asigno tiempos iniciales para comparar 16:01
         tiempo5 = datetime(2022,1,1,17,1,00) # asigno tiempos iniciales para comparar 17:01  Todas las tiendas cierre 
         tiempo6 = datetime(2022,1,1,15,1,00) # asigno tiempos iniciales para comparar 15:01 Economy Sabado
         
         if fecha1 != 99:
            if dia in range(0,6) :   
             if fecha1 == 1 and dia==5: 
                tabla(tiempo2,'1.2','1.3')
             elif fecha1 == 1:
                tabla(tiempo4,1,'1.1')     
             if fecha1 == 2 and dia==5: 
                tabla(tiempo2,'2.2','2.3')
             elif fecha1 == 2:  
                tabla(tiempo4,2,'2.1')
             if fecha1 == 3 and dia==5:
                tabla(tiempo2,'3.2','3.3')
             elif fecha1 == 3:
                tabla(tiempo4,3,'3.1')      
             if fecha1 == 4: 
                tabla(tiempo3,4,'4.1')     
             if fecha1 == 5:    
                tabla(tiempo1,5,'5.1')  
             if fecha1 == 6:    
                tabla(tiempo4,6,'6.1')         
             if fecha1 == 7:    
                tabla(tiempo3,7,'7.1')  
             if fecha1 == 8:    
                tabla(tiempo1,8,'8.1')
             if fecha1 == 9:    
                tabla(tiempo4,9,'9.1')
             if fecha1 == 10:  
                tabla(tiempo4,10,'10.1')
             if fecha1 == 11: 
                tabla(tiempo4,11,'11.1') 
             if fecha1 == 12:    
                tabla(tiempo1,12,'12.1')
             if fecha1 == 13:
                tabla(tiempo4,13,'13.1')
             if fecha1 == 14:    
                tabla(tiempo3,14,'14.1')
             if fecha1 == 15:    
                tabla(tiempo1,15,'15.1')
             if fecha1 == 16:    
                tabla(tiempo4,16,'16.1') ###############
             if fecha1 == 17 and dia==5: 
                tabla(tiempo3,'17.2','17.3')  
             elif fecha1 == 17:
                tabla(tiempo5,17,'17.1') 
             if fecha1 == 18 and dia==5: 
                tabla(tiempo6,'18.2','18.3')           
             elif fecha1 == 18:    
                tabla(tiempo5,18,'18.1')
             if fecha1 == 19 and dia==5: 
                tabla(tiempo2,'19.2','19.3')   
             elif fecha1 == 19:    
                tabla(tiempo5,19,'19.1') 
             if fecha1 == 20 and dia==5: 
                tabla(tiempo3,'20.2','20.3')    
             elif fecha1 == 20:    
                tabla(tiempo5,20,'20.1')          
            elif dia== 6 :
             print("Domingo no es dia laborable")
             return()   
         elif fecha1 == 99: # Todo lo que sea 99
             tabla1(tiempo4,1,'1.1')
         #print("Fecha compromisoddd:" + str(fechaCompromiso)) 
         #print(fechaCompromiso)
         return fechaCompromiso.date() #Regreso la fecha compromisa de la parte


    ###################    
    main_window = tk.Tk()
    main_window.title("Fecha compromiso TAPATIO")
    main_window.geometry("310x250")
    main_window.iconbitmap("logoicon.ico")
    main_window.resizable(False, False)
    #Display image
    image = Image.open("logo-new.png")
    image = image.resize((70,30), Image.ANTIALIAS)
    image = ImageTk.PhotoImage(image)
    label_image = tk.Label(image=image).place(x=0,y=1)

    label = tk.Label(main_window, text="", anchor="w",font=("times", 7))
    label.pack(side="top")
    #print("eso:",selected)
    if super_user == 1:
       label.config(text=selected + " (SUPER USER)")
    else    :
       label.config(text=selected)
    main_window.bind("<KeyRelease>", form_complete)
    # Display clock
    clock = tk.Label(main_window, justify=tk.RIGHT,font=("times", 9, "bold"),fg="blue")
    #clock.pack()
    clock.place(x=130,y=16)
    
    

    def changePromiseDay():
        nonlocal cambio
        calendar_window = tk.Tk()
        calendar_window.title("Login")
        calendar_window.geometry("300x300")
        calendar_window.iconbitmap("logoicon.ico")
        calendar_window.resizable(False, False)
        calendar_window.attributes('-topmost',True)
        ##print("Promise Day button pressed.")

        def grad_date():
             nonlocal dates
             nonlocal cambio
             ##print(dates)
             ##print(cal.get_date())# > dates :
             c = cal.get_date()
             d = dt.strptime(c, '%m/%d/%y')
             e = dt.strptime(dates, '%Y-%m-%d')
             #print(d)
             #print(e)
             if e < d :
              #date.config(text = "Nueva Fecha compromiso es: " + cal.get_date())
              #fech = cal.get_date()
              #print(dates)
              #text_widget.configure(state='normal')
              #text_widget.insert(tk.END, 'La nueva fecha compromiso es : ' + '\n'+ dt.strptime(cal.get_date(), '%m/%d/%y').strftime("%a,%d %b, %Y"))
              #text_widget.configure(state='disabled')
              #compromiso = Label(main_window,font=("Courier 10 bold"), text = "                                                            ")
              #compromiso.place(x = 230,y = 180) 
              #compromiso = Label(main_window,font=("Courier 13 bold"),fg="red", text = "                                           " )
              #compromiso.place(x = 230,y = 200)
              messagebox.showinfo("showinfo", "La nueva Fecha Compromiso para el cliente cambio de :"+ '\n'+ dt.strptime(dates, '%Y-%m-%d').strftime("%a,%d %b, %Y") + 
                "  a  " + dt.strptime(cal.get_date(), '%m/%d/%y').strftime("%a,%d %b, %Y"))
              #compromiso = Label(main_window,font=("Courier 10 bold"), text = "La nueva Fecha Compromiso para el cliente cambio de :")
              #compromiso.place(x = 230,y = 180)  
              #f1=("Times", 22, 'overstrike')
              #tk.Label(my_w,fg='green',text='strikethrough text word',font=f1)
              #compromiso = Label(main_window,font=("Courier 13 bold overstrike"),fg="red", text =  dt.strptime(dates, '%Y-%m-%d').strftime("%a,%d %b, %Y"))
              #compromiso.place(x = 230,y = 200)
              #compromiso = Label(main_window,font=("Courier 13 bold"),fg="green", text =  dt.strptime(cal.get_date(), '%m/%d/%y').strftime("%a,%d %b, %Y"))
              #compromiso.place(x = 410,y = 200)
              #data.append([str(orden.get()),d,selected])
              #print(data)
              #dates = cal.get_date()
              dates = dt.strptime(cal.get_date(), '%m/%d/%y').strftime("%Y-%m-%d")
              df1['DueDate change'] = '*'
              #if cambio == 1:
              data.extend(["-","-","-","-","-","-"])   
              data.extend([dates])
              #print(dates)
              data.extend("X")
              cambio = 0
              calendar_window.destroy()
              subir()

        cal = Calendar(calendar_window, selectmode = 'day',
                       year = current_time.year, month = current_time.month,
                       day = current_time.day)

        cal.pack(pady = 20)
        #cal.place(x = 100,y = 100)
        Button(calendar_window, text = "Cambiar Fecha Compromiso",command = grad_date).pack(pady = 5)    
        date = Label(main_window, text = "")
        date.place(x=200,y=450)
        cambio  = 1

        #data.extend("x")

    def subir():
      
        #ss = file.open("EficienciaReporte")
        #hoja = ss.worksheet(selected)     
        hoja.append_row(data)  
        data.clear() # habilitar cuando este lita la subida
        
    def subir2():

        hoja = ss.worksheet(selected)
        ##print(data)
        hoja.append_row(data)
        data.clear() # habilitar cuando este lita la subida

    def tick():
        time_string = time.strftime("%H:%M:%S")
        clock.config(text=time_string)
        clock.after(200, tick)

    tick()
    #global orden
    #tk.Label(main_window,text='Seleccione la tienda donde se encuetra la parte: ',font=("Courier 10 bold"),fg="blue").place(x=0,y=110)
    var = tk.IntVar()
    var.set(1)
    radio1 = tk.Radiobutton(main_window, text="TAP 1", variable=var, value=1).place(x=0,y=70)  #65
    radio2 = tk.Radiobutton(main_window, text="TAP 2", variable=var, value=2).place(x=0,y=90)
    radio3 = tk.Radiobutton(main_window, text="TAP 4 & 14", variable=var, value=4).place(x=60,y=70)
    radio4 = tk.Radiobutton(main_window, text="TAP 6", variable=var, value=6).place(x=60,y=90)
    radio5 = tk.Radiobutton(main_window, text="TAP 7 & 8", variable=var, value=7).place(x=150,y=70)
    radio6 = tk.Radiobutton(main_window, text="TAP 10", variable=var, value=10).place(x=150,y=90)
    radio7 = tk.Radiobutton(main_window, text="TAP 15", variable=var, value=15).place(x=230,y=70)
    

    part_button = tk.Button(main_window, text="Import/Export", command=check_export,state=tk.NORMAL)#.place(x=10,y=450)
    part_button.place(x=215,y=90)
    # Submit button

    #tk.Label(main_window,text='Hora corte (manual)',font=("Courier 10 bold"),fg="green").place(x=0,y=110)
    if super_user :
     var2 = tk.IntVar()
     #var2.set()
     check = tk.Checkbutton(main_window, text="Hora corte (manual)",font=("Courier 10 bold"),fg="black", variable=var2, onvalue=1, offvalue=0,command=seleccionar)
     check.place(x=0,y=450)
     #Checkbutton(frame, text="Con azúcar", variable=leche, 
     #       onvalue=1, offvalue=0).pack(anchor=W)


    global entry1

    name = Label(main_window,font=("Courier 10 bold"), text = "Job #:").place(x = 4,y = 40)   ########### CLEAR BUTTON
    #entry1 = Entry(main_window).place(x = 300, y = 140)
    orden = tk.StringVar(main_window)
    entry1 = tk.Entry(main_window, textvariable=orden,width=20, validate="key",validatecommand=(main_window.register(validation), "%d","%S","%P"))   
    entry1.place(x = 50, y = 40)
    # Combobox
    combobox = ttk.Combobox(main_window,state="readonly",values=[
                     "15 EAST", "15 NORTH 2", "15 SOUTH", "15 WEST","5 NORTH","5 NORTH 2","5 NORTH 3","5 EAST","5 WEST","AREA SD","SD AUX"
                     ,"SHOP SD","ENSENADA","TIJUANA","EBAY TJ","SHIPPING","WILL CALL 1","WILL CALL 6","WILL CALL 7","PAQUETERIA TJ"
                     ,"SHIP WC1", "SHIP WC6"])
    combobox.pack(padx=3,side = LEFT)
    combobox.set("15 EAST")
    
    #tk.Label(main_window, text='Seleccione la ruta o locacion de entrega: ',font=("Courier 10 bold"),fg="blue").place(x=0,y=230)
    tk.Label(main_window, text='Ver.11 Ago/2023',font=("Courier 6 bold"),fg="black").place(x=210,y=2)        
    # Add button
    add_button = tk.Button(main_window, text="Add", command=add_action,state=tk.DISABLED)#.place(x=10,y=450)
    add_button.pack(side = LEFT)
    main_window.bind('<Return>', lambda event=None: add_button.invoke())
    # Submit button
    submit_button = tk.Button(main_window, text="Calcular", command=submit_action,state=tk.DISABLED)#.place(x=60,y=450)
    submit_button.pack(side = LEFT)

    cambio_button = tk.Button(main_window, text="Cambiar", command=changePromiseDay,state=tk.DISABLED)
    #cambio_button.place(x=425,y=450)
    cambio_button.pack(side = LEFT)

    text_widget = tk.Text(main_window, state='disabled',height=5, width=37)
    text_widget.place(x=4,y=154)
    
    tk.Label(main_window, text=selected_user).place(x=430,y=450)
    clear_button = tk.Button(main_window, text="Clear", command=clearOrden,state=tk.NORMAL)    ########### CLEAR BUTTON
    clear_button.place(x=177,y=37)
    #cambio_button = tk.Button(main_window, text="Cambiar", command=changePromiseDay,state=tk.DISABLED)
    #cambio_button.place(x=425,y=450)
    enviar_button = tk.Button(main_window, text="Enviar", command=submit,state=tk.DISABLED)
    enviar_button.place(x=485,y=450)
    tk.Button(main_window, text="Salir", command=close_window).place(x=240,y=37)

    #compromiso = Label(main_window,font=("Courier 10 bold"), text = "gdfgdfgdfgdfgdfgdf")
    #compromiso.place(x = 230,y = 180) 

    
    
    data = []
    fechas = []
    dates = ''
   # compromiso=''
    cambio = 0
    main_window.attributes('-topmost',True)
    main_window.mainloop()

def login():
    def enter():
        global selected
        super_user = 0 

        selected = combobox.get()
        password = simpledialog.askstring("Password", "Enter the password for '" + selected + "':", show='*')
        if password == users[selected]:
            if selected in ["MANUEL RAZO","EMMANUEL LOPEZ","KARLA CRUZ","MIGUEL CERVANTES","JUAN ORTIZ"]:
               super_user = 1
               ##print(selected," is a Super User")
            ##print("Stamp:", selected)
            ##ss = file.open("EficienciaReporte")
            hoja = ss.worksheet(selected) 
            login_window.destroy()
            main_program(super_user,hoja)
        else:
            messagebox.showerror("Error", "Incorrect password")

    login_window = tk.Tk()
    login_window.title("Login")
    login_window.geometry("300x150")
    login_window.iconbitmap("logoicon.ico")
    login_window.resizable(False, False)

    imag = Image.open("logo-new.png")
    imag1 = imag.resize((130,80), Image.ANTIALIAS)
    imag1 = ImageTk.PhotoImage(imag1)
    label_image2 = tk.Label(image=imag1).place(x=0,y=40)

    image = Image.open("motor.png")
    image1 = image.resize((150,110), Image.ANTIALIAS)
    image1 = ImageTk.PhotoImage(image1)
    label_image1 = tk.Label(image=image1).place(x=160,y=20)
   
    combobox = ttk.Combobox(login_window, values=list(users.keys()))
    combobox.current(0)
    combobox.pack()

    enter_button = tk.Button(login_window, text="Enter", command=enter)
    enter_button.pack()
    login_window.bind('<Return>', lambda event=None: enter_button.invoke())
    login_window.attributes('-topmost',True)

    login_window.mainloop()
    
selected_user = ""
login()
