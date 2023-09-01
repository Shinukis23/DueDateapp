import tkinter as tk
from tkinter import messagebox
from tkinter import *
import csv
from selenium import webdriver
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
#import xml.etree.ElementTree as ET
import pandas as pd
import numpy as np
import xlwt
import gspread
import locale
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
import ssl
from bs4 import BeautifulSoup
from urllib.request import urlopen
import certifi
import re
import urllib.request
import requests
sess = requests.Session()
adapter = requests.adapters.HTTPAdapter(max_retries = 20)
sess.mount('http://', adapter)
 
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)

scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

proxies = {'https': 'http:tapaserver.dyndns.org'}
credentials = ServiceAccountCredentials.from_json_keyfile_name("monitor-eficiencia-3a13458926a2.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials)# authenticate the JSON key with gspread
ss = file.open("EficienciaReporte") #1
# Define the Drive API client
service = build('drive', 'v3', credentials=credentials)

url3 = "https://www.banxico.org.mx/SieInternet/consultarDirectorioInternetAction.do?sector=6&accion=consultarCuadro&idCuadro=CF102&locale=es"
html = sess.get(url3).content
df_list = pd.read_html(html)
df = df_list[0]
#print(df)
resultWeekend=df.iloc[8,2]
#resultday=df.iloc[7,2]
#print(resultWeekend)
i = 1
super_user = 0 
result = ""
result1 = ""
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
DOF = url = URL = page = soup = texto = "" 
diff = 0.0
data = []
fechas = []
dates = ''

cambio = 0
entry1 = 0
Tipo_Cambio = 0.0

def scrap_web():
    global DOF
    global Tipo_Cambio
    global result
    global result1
    global diff
    global i
    url =  service.files().get_media(fileId="18F4Ix9C_2q7tjimDycg6v7XnCAb258FM").execute()
    URL = url.decode('utf-8')
    #url = "https://dof.gob.mx/index.php#gsc.tab=0"
    unverified_context = ssl._create_unverified_context()
    page = urlopen(URL, context=unverified_context)
    #page = urlopen(URL)
    url =  service.files().get_media(fileId="1Y-5sSIKrF1HmV58gPJpTWZE7yB4Qa9Ee").execute()
    diff = float(url.decode('utf-8'))
    ##print(diff)
    page = page.read().decode("utf-8")
    soup = BeautifulSoup(page, "html.parser")
    texto =soup.get_text()
    if datetime.today().weekday() != 6 and datetime.today().weekday() != 5:
    #if datetime.today().weekday() == 6 or datetime.today().weekday() == 5:
        resultado = re.search('DOLAR (.*\d)UDIS', texto)
        #resultado1 = re.search('Tipo de Cambio y Tasas al (.*)', texto)
        #result1 = resultado1.group(1)
        #print(resultado)
        if resultado is None:
            result = resultWeekend
            date = datetime.today() - timedelta(days = 3)
            result1 = date.strftime("%Y/%m/%d")#, %H:%M:%S")
            #print("test",result1)
        else:
            result=resultado.group(1)
            resultado1 = re.search('Tipo de Cambio y Tasas al (.*)', texto)
            result1 = resultado1.group(1)
            #print("este es",result) 
            
    elif datetime.today().weekday() == 6 or datetime.today().weekday() == 5:
    #elif datetime.today().weekday() != 6 and datetime.today().weekday() != 5:
        result = resultWeekend
        resultado1 = re.search('Tipo de Cambio y Tasas al (.*)', texto)
        result1 = resultado1.group(1)
        date=dt.strptime(result1, "%d/%m/%Y")
        if datetime.today().weekday() == 6: 
            date = date - timedelta(days = 2)
            #print(date)  
        elif datetime.today().weekday() == 5: 
            date = date - timedelta(days = 1)
            
        result1 = date.strftime("%Y/%m/%d")#, %H:%M:%S")
        print("Es dia inhabil:",result) 
    Tipo_Cambio = float(result)
    DOF = "Tipo de Cambio y Tasas al "+ result1
    i = i + 1
    print(datetime.today().strftime("%H:%M:%S"))


scrap_web()
url = "https://drive.google.com/file/d/1CgwK7OCtiTXggOZpw3E39PgMlPEQo8WP/view?usp=share_link" #claves.csv
output_path = 'claves.csv'
gdown.download(url, output_path, quiet=False,fuzzy=True)

with open(r'claves.csv', mode='r') as infile:
    reader = csv.reader(infile)
    with open('claves_new.csv', mode='w') as outfile:
        writer = csv.writer(outfile)
        users = {rows[0]:rows[1] for rows in reader}
#print(users)


USA = 0
MEXICO = 1


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
control = 1
current_time = datetime.now()


def submit():
    global diff
    global result1
    global result
    global DOF
    global Tipo_Cambio
    global resultWeekend
    global control
    global i
    Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="black",text="Actualizando...").place(x=195,y=16)
    T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="green",text="                  ").place(x=270,y=16)
    scrap_web()
    enviar_button['state'] = NORMAL
    Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="black",text="Tipo_Cambio: ").place(x=195,y=16)
    name = Label(main_window,font=("Courier 9 bold"), text = DOF).place(x = 14,y = 40)
    Tipo_Cambio = float(result)
    if control == 1:
        T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="green",text=Tipo_Cambio).place(x=270,y=16)  
        control = 0
    else :    
        T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="red",text=Tipo_Cambio).place(x=270,y=16)
        control = 1#time.sleep(1)
    #T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="green",text=Tipo_Cambio).place(x=270,y=16)
    #T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="green",text=Tipo_Cambio).place(x=270,y=16)
    Mex = Label(main_window,font=("Courier 17 bold"),fg="red", text = "                         ").place(x = 34,y = 130)
    USD = Label(main_window,font=("Courier 17 bold"),fg="red", text = "                         ").place(x = 34,y = 180)
    entry1.config(state= "normal")
    print ("Update",i)
    print (Tipo_Cambio)
    print (DOF)
    orden.set('')

def submit_action():
    global dates

    fechas.clear()
    submit_button['state'] = DISABLED
    enviar_button['state'] = NORMAL
    entry1.config(state= "disabled")
    clear_button['state'] = NORMAL
    data.extend(["-","-","-","-","-","-"])   
    data.extend([dates])
    subir()

def clearOrden():
        
    entry1.config(state= "normal")
    submit_button['state'] = DISABLED
    enviar_button['state'] = NORMAL
    Mex = Label(main_window,font=("Courier 17 bold"),fg="red", text = "                         ").place(x = 34,y = 130)
    USD = Label(main_window,font=("Courier 17 bold"),fg="red", text = "                         ").place(x = 34,y = 180)
    data.clear()
    orden.set('')
    fechas.clear()
    data.extend(["-","-","-","-","-","-","-"])   
    data.extend(["Clear"])

def close_window():
    main_window.destroy()
    login()

def validation(i,text, new_text): 
    return len(new_text) == 0  or len(new_text) < 10 and text.isdecimal()
  
def form_complete(event):
    if len(orden.get()) <= 0 :
        submit_button['state'] = DISABLED

    else:
        submit_button['state'] = NORMAL
        Mex = Label(main_window,font=("Courier 17 bold"),fg="red", text = "                         ").place(x = 34,y = 130)
        USD = Label(main_window,font=("Courier 17 bold"),fg="red", text = "                         ").place(x = 34,y = 180)
    ##################    
def seleccionar():
    #print("Super User")
    f = g


def subir(): 
    global diff
    global result
    global control
    global Tipo_Cambio
    global i
    global DOF
    #print('subir',result)
    pacifictime = datetime.now(pytz.timezone('US/pacific'))
    current_time = pacifictime.strftime("%Y-%m-%d %H:%M:%S")
    Tipo_Cambio = float(result)
    if control == 1:
        T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="green",text=Tipo_Cambio).place(x=270,y=16)  
        control = 0
    else :    
        T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="red",text=Tipo_Cambio).place(x=270,y=16)
        control = 1#time.sleep(1)
    #T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="green",text=Tipo_Cambio).place(x=270,y=16)
    name = Label(main_window,font=("Courier 9 bold"), text = DOF).place(x = 14,y = 40)
    PrecioPower=orden.get()
    PrecioMex = float(PrecioPower) * Tipo_Cambio 
    PrecioUSD = float(PrecioMex / (Tipo_Cambio-diff))
    Mex = Label(main_window,font=("Courier 18 bold"),fg="red", text =float("{:.4f}".format(PrecioMex))).place(x = 34,y = 130)
    USD = Label(main_window,font=("Courier 18 bold"),fg="red", text =float("{:.4f}".format(PrecioUSD))).place(x = 34,y = 180)
    orden.set('')
    data.clear()
    data.extend([selected,"-",current_time,"-","-","-","-","-","-",Tipo_Cambio,PrecioPower,locale.currency(PrecioUSD, grouping=True),locale.currency(PrecioMex, grouping=True)]) # 1
    hoja.append_row(data)  #1
    
    data.clear() 
    entry1.config(state= "normal")# habilitar cuando este lita la subida
        
def subir2():

    #hoja = ss.worksheet(selected)
    ##print(data)
    #hoja.append_row(data)
    data.clear() # habilitar cuando este lita la subida

def tick():
    time_string = time.strftime("%H:%M:%S")
    clock.config(text=time_string)
    clock.after(200, tick)



def Actualizacion():
    global DOF
    global resultWeekend
    global diff
    global result1
    global result
    global Tipo_Cambio
    #global i
    while True:
        submit()
        #scrap_web()
        time.sleep(1*60*60)

def login():
    def enter():
        global selected
        global super_user #= 0 

        selected = combobox.get()
        password = simpledialog.askstring("Password", "Enter the password for '" + selected + "':", show='*')
        if password == users[selected]:
            if selected in ["MANUEL RAZO","EMMANUEL LOPEZ","KARLA CRUZ","MIGUEL CERVANTES","JUAN ORTIZ"]:
               super_user = 1
             
            login_window.destroy()
            #main_program(super_user)
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
hoja = ss.worksheet(selected)  
main_window = tk.Tk()
main_window.title("Tipo de Cambio")
main_window.geometry("330x250")
main_window.iconbitmap("logoicon.ico")
main_window.resizable(False, False)
#Display image
image = Image.open("logo-new.png")
image = image.resize((80,35), Image.ANTIALIAS)
image = ImageTk.PhotoImage(image)
label_image = tk.Label(image=image).place(x=10,y=1)
label = tk.Label(main_window, text="", anchor="w",font=("times", 7))
label.pack(side="top")
#print("eso:",selected)
if super_user == 1:
    label.config(text=selected + " (SUPER USER)")
else    :
    label.config(text=selected)
main_window.bind("<KeyRelease>", form_complete)
# Display clock
clock = tk.Label(main_window, justify=tk.RIGHT,font=("times", 10, "bold"),fg="blue")
#clock.pack()
clock.place(x=130,y=16)
tick()
if super_user :
    var2 = tk.IntVar()
    check = tk.Checkbutton(main_window, text="Hora corte (manual)",font=("Courier 10 bold"),fg="black", variable=var2, onvalue=1, offvalue=0,command=seleccionar)
    check.place(x=0,y=450)


name = Label(main_window,font=("Courier 9 bold"), text = DOF).place(x = 14,y = 40)
name = Label(main_window,font=("Courier 12 bold"), text = "Precio en PowerLink (USD):").place(x = 14,y = 60)   ########### CLEAR BUTTON
#entry1 = Entry(main_window).place(x = 300, y = 140)
orden = tk.StringVar(main_window)
entry1 = tk.Entry(main_window, textvariable=orden,width=20,borderwidth=2, validate="key",validatecommand=(main_window.register(validation), "%d","%S","%P"))   
#entry1.insert(0, "This is Temporary Text...")
entry1.place(x = 16, y = 80)
entry1.focus_set()
#entry1.bind("<FocusIn>", temp_text)
tk.Label(main_window, text='Ver.7.1 June/2023',font=("Courier 7 bold"),fg="black").place(x=225,y=1)        
# Add button
PrecioClMex = Label(main_window,font=("Courier 14 bold"),fg="blue", text = "Precio al Cliente en Pesos").place(x = 14,y = 110)
PrecioClUSD = Label(main_window,font=("Courier 14 bold"),fg="blue", text = "Precio al Cliente en Dolares").place(x = 14,y = 160)
Simb = Label(main_window,font=("Courier 17 bold"),fg="red", text ="$").place(x = 12,y = 130)
Simb = Label(main_window,font=("Courier 17 bold"),fg="red", text ="$").place(x = 12,y = 180)    
submit_button = tk.Button(main_window, text="Calcular", command=submit_action,state=tk.DISABLED)#.place(x=60,y=450)
submit_button.place(x = 250, y = 220)
main_window.bind('<Return>', lambda event=None: submit_button.invoke())
    
tk.Label(main_window, text=selected_user).place(x=430,y=450)
clear_button = tk.Button(main_window, text="Clear", command=clearOrden,state=tk.NORMAL)    ########### CLEAR BUTTON
clear_button.place(x=10,y=220)
enviar_button = tk.Button(main_window, text="Update", command=submit,state=tk.NORMAL)
enviar_button.place(x=200,y=220)
tk.Button(main_window, text="Salir", command=close_window).place(x=50,y=220)
Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="black",text="Tipo_Cambio: ").place(x=195,y=16)

T_Cambio = tk.Label(main_window,font=("times", 10, "bold"),fg="green",text=Tipo_Cambio).place(x=270,y=16)
#clock.pack()
#T_Cambio


thread = threading.Thread(target=Actualizacion)
thread.daemon = True  # Set the thread as a daemon so it won't prevent the script from exiting
thread.start()
main_window.attributes('-topmost',True)
main_window.mainloop()
