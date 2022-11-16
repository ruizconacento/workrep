#import tkinter as tk
from asyncio.windows_events import NULL
from cmath import nan

from curses.ascii import isdigit
from sqlite3 import Timestamp
from tkinter import *
from tkinter import filedialog
from numpy import NaN, chararray #dialogo de archivos
import pandas as pd
import psycopg2
import json
import requests
from datetime import datetime
import math

urlR = ""
HEIGHT = 600
WIDTH = 1200
excel_path = ""
excel_path_equipos = ""
excel_path_tecnicos = ""


def funcion_abrir():
        #print("this is the entry:", entry)
        root.filename = filedialog.askopenfilename(title = "Seleccionar excel", filetypes=[("Excel files", "*.xlsx *.xls *.xlrd")])
        global excel_path
        excel_path = root.filename
        #labelExcelUrl = excel_path

        labelServicioUrl.config(text=excel_path)

        #my_label = tk.Label(root, text=root.filename)
        #my_label.pack()



def funcion_abrir_equipos():
        #print("this is the entry:", entry)
        root.filename = filedialog.askopenfilename(title = "Seleccionar excel", filetypes=[("Excel files", "*.xlsx *.xls *.xlrd")])
        global excel_path_equipos
        excel_path_equipos = root.filename
        #labelExcelUrl = excel_path_equipos

        labelEquiposUrl.config(text=excel_path_equipos)

        #my_label = tk.Label(root, text=root.filename)
        #my_label.pack()


def funcion_abrir_tecnicos():
        #print("this is the entry:", entry)
        root.filename = filedialog.askopenfilename(title = "Seleccionar excel", filetypes=[("Excel files", "*.xlsx *.xls *.xlrd")])
        global excel_path_tecnicos
        excel_path_tecnicos = root.filename
        #labelExcelUrl = excel_path_equipos

        labelTecnicosUrl.config(text=excel_path_tecnicos)

        #my_label = tk.Label(root, text=root.filename)
        #my_label.pack()


def funcion_abrir_direcciones():
        #print("this is the entry:", entry)
        root.filename = filedialog.askopenfilename(title = "Seleccionar excel", filetypes=[("Excel files", "*.xlsx *.xls *.xlrd")])
        global excel_path_direcciones
        excel_path_direcciones = root.filename
        #labelExcelUrl = excel_path_equipos

        labelDireccionesUrl.config(text=excel_path_direcciones)

        #my_label = tk.Label(root, text=root.filename)
        #my_label.pack()




def hacer_request(jsonAEnviar):
        global urlR
        jsonR = jsonAEnviar
        response = requests.post(urlR, json=jsonR)
        MainLabel.config(text=response)


def enviar_db_servicios():

############# POR SI QUIERO TENER DATOS DEL DB EN TXT excterno
        #myfile = open("url.txt", "rt") # open lorem.txt for reading text
        #contents = myfile.read()         # read the entire file into a string
        #myfile.close()                   # close the file
        #print(contents)                  # print contents


        print (excel_path)

        if excel_path != "":
                print("intentando enviar:")
                df = pd.read_excel (excel_path) #for an earlier version of Excel, you may need to use the file extension of 'xls'
                #df = df.drop(df.index[df.eq('').all(axis=1)]) #drop all null rows

                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df['Plan mant.preventivo'] = df['Plan mant.preventivo'].fillna('', inplace=True)


                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df = df.iloc[8:] #no hay que bajar columnas


                connection = psycopg2.connect(
                        host='database-sf.czf4buzwfvge.us-east-2.rds.amazonaws.com',
                        user='SFpostgres',
                        password='sf_admin_fuerza',
                        database='postgres'
                )
                cursor = connection.cursor()


                for index, row in df.iterrows():
                        print(row[0],row[1],row[2],row[3])
                        fecha1 = None

                        

                        if isinstance(row[0], str):
                                #if row[0].count(".") > 0:  #cuando tiene .
                                if row[0].count("-") > 0:  #cuando tiene -
                                        #fecha1 = row[0].replace(".","/") #cuando tiene .
                                        fecha1 = row[0].replace("-","/") #cuando tiene -
                                else:
                                        fecha1 = row[0]
                        else:
                                fecha1 = row[0]
                        



                        tipo = type(fecha1)
                        if isinstance(fecha1, pd.Timestamp):
                                #fecha1 =  str(fecha1)
                                #dt_obj = datetime.fromtimestamp(fecha1)

                                fecha1 = fecha1.date()
                                #datetime.date(2013, 12, 25)

                                #fecha1 = dt_obj
                                #fecha1 = dt_obj.date()
                                fecha1 = str(fecha1)
                                fecha1 = fecha1.replace("-","/")
                                
                                date_time_obj = datetime.strptime(fecha1.strip(), '%Y/%m/%d')
                                fecha1 = date_time_obj.strftime("%Y/%m/%d")


                        ###convertir de D/M/YYYY a YYYY/D/M para postgres
                        #date_time_obj = datetime.strptime(fecha1.strip(), '%d/%m/%Y')
                        ###print("The type of the date is now",  type(date_time_obj))
                        #fecha2 = date_time_obj.strftime("%Y/%d/%m")



                        fecha2 = fecha1 #para cuando no se ocupa hacer cambio de mover fechas
                        print(fecha2)
                        #print(fecha2)

                        #print(type(row[1])) 
                        #print(len(row[1]))                        
                        #print(len(row[1].strip()))

                        mantp = None  #ignorado por mientras
                        if math.isnan(row[2]):
                                mantp = None
                        else:
                                mantp = row[2]

                        equipoiddd = None
                        equipoid = None
                        if math.isnan(row[9]):
                                equipoid = None
                        else:
                                equipoid = int(row[9]) #a veces me convierte a float esto lo hace siempre int

                        clienteid = None
                        if math.isnan(row[6]):
                                clienteid = None
                        else:
                                clienteid = int(row[6])

                        serie = None
                        if (row[8]) is None:
                                serie = ""
                        else:
                                serie = str(row[8])

                        denominacion = None
                        if (row[7]) is None:
                                denominacion = ""
                        else:
                                denominacion = str(row[7])

                        textob = None
                        if (row[5]) is None:
                                textob = ""
                        else:
                                textob = str(row[5])

                        print(equipoiddd)
                        tipoEquipoOb = type(equipoiddd)
                        print(tipoEquipoOb)

                        #postgres_insert_query = """ INSERT INTO tabla_programacion (p_orden, p_fecha_pivote, p_mant, p_tec, p_status, p_textob, p_cliente, p_denomicacion, p_numserie, p_equipo, p_ci, p_sis_status, p_sis_tecnico) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

                        postgres_insert_query = """ SELECT * FROM __insert_or_update_programacion_carga(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

                        #record_to_insert = (row[1],row[0],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],0,0)
                        #record_to_insert = (row[1],fecha2,mantp, row[3], row[4],row[5],clienteid,row[7],row[8],None,row[10],0,None)

                        #record_to_insert = (row[1],fecha2, None, row[3], row[4],row[5],clienteid,row[7],str(row[8]),equipoid,row[10],0,None)
                        
                        #record_to_insert = (row[1],fecha2, None, row[3], row[4],row[5],  clienteid, None, serie, None, row[10] ,0,None)  #test
                        record_to_insert = (row[1],fecha2, None, row[3], row[4], textob,  clienteid, denominacion, serie, equipoid, row[10] ,0,None)  #un poco mejorado
                        
                        
                        cursor.execute(postgres_insert_query, record_to_insert)
                        ##cursor.execute("Select InsertTest (temp1 :=" + str(tempDecimal) + ", mac := '" +  macString + "')")
                        ##row = cursor.fetchone()
                        
                connection.commit()

                        
                
                MainLabel.config(text="carga completa")
                ########hacer_request(json_data)#########
                print ("dd")
        else:
                MainLabel.config(text="error al leer archivo")
        

def enviar_db_equipos():

        print (excel_path_equipos)
        if excel_path_equipos != "":
                print("intentando enviar:")
                df = pd.read_excel (excel_path_equipos) #for an earlier version of Excel, you may need to use the file extension of 'xls'
                #df = df.drop(df.index[df.eq('').all(axis=1)]) #drop all null rows

                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df['Plan mant.preventivo'] = df['Plan mant.preventivo'].fillna('', inplace=True)


                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df = df.iloc[8:] #no hay que bajar columnas


                connection = psycopg2.connect(
                        host='database-sf.czf4buzwfvge.us-east-2.rds.amazonaws.com',
                        user='SFpostgres',
                        password='sf_admin_fuerza',
                        database='postgres'
                )
                cursor = connection.cursor()


                for index, row in df.iterrows():
                        print(row[0],row[1],row[2],row[3])

                        curso = False
                        dc3 = False
                        if row[10] == "SI":
                                curso = True
                        
                        if row[11] == "SI":
                                dc3 = True


                        a = row[4]
                        #if[row[4] ]

                        e_latitud = None
                        if math.isnan(row[4]):
                                e_latitud = None
                        else:
                                e_latitud = str(row[4])

                        e_longitud = None
                        if math.isnan(row[5]):
                                e_longitud = None
                        else:
                                e_longitud = str(row[4])


                        tiempo = 0
                        if math.isnan(row[8]):
                                tiempo = 0
                        else:
                                tiempo = row[8]

                        #e_longitud = str(row[5])

                        clienteid = None
                        if math.isnan(row[2]):
                                clienteid = None
                        else:
                                clienteid = int(row[2])


                        #REVISAR EQUIPO NUMERO 318 DEL ARCHIVO EQUIPO1
                        equipoid = 0
                        if isinstance(row[0], str):
                                equipoid = None
                        else:
                                #postgres_insert_query = """ INSERT INTO tabla_equipos (e_equipo_id, e_tecnico, e_cliente, e_denominacion, e_latitud, e_longitud, e_frecuencia, e_kw, e_tiempo, e_diaservicio, e_curso, e_dc3, e_extra, e_sys_point) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                                #postgres_insert_query = """ SELECT * from __insert_equipo (e_equipo_id, e_tecnico, e_cliente, e_denominacion, e_latitud, e_longitud, e_frecuencia, e_kw, e_tiempo, e_diaservicio, e_curso, e_dc3, e_extra) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                                #postgres_insert_query = """ SELECT * from __insert_equipo (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                                postgres_insert_query = """ SELECT * from __insert_or_update_equipo (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

                                #record_to_insert = (row[1],row[0],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],0,0)
                                #record_to_insert = (row[0],row[1],row[2], row[3], row[4],row[5],row[6],row[7],row[8],row[9],curso,dc3,row[12],None)
                                e_equipo_id = int(row[0])
                                e_tecnico = str(row[1])


                                #record_to_insert = (row[0],str(row[1]),row[2], str(row[3]), str(row[4]),str(row[5]),str(row[6]),row[7],row[8],str(row[9]),curso,dc3,str(row[12]))
                                #record_to_insert = (e_equipo_id,e_tecnico,row[2], str(row[3]), str(row[4]),str(row[5]),str(row[6]),row[7],row[8],str(row[9]),curso,dc3,str(row[12]))
                                # record_to_insert = (e_equipo_id, 
                                #                         e_tecnico, 
                                #                         row[2], 
                                #                         str(row[3]), 
                                #                         e_latitud, 
                                #                         e_longitud, 
                                #                         str(row[6]), 
                                #                         str(row[7]), 
                                #                         row[8],
                                #                         str(row[9]),
                                #                         curso,
                                #                         dc3,
                                #                         str(row[12]))

                                record_to_insert = (e_equipo_id, 
                                                        e_tecnico, 
                                                        clienteid,
                                                        str(row[3]), 
                                                        e_latitud, 
                                                        e_longitud, 
                                                        str(row[6]), 
                                                        str(row[7]), 
                                                        tiempo,
                                                        str(row[9]),
                                                        curso,
                                                        dc3,
                                                        str(row[12]))
                                
                                cursor.execute(postgres_insert_query, record_to_insert)
                                ##cursor.execute("Select InsertTest (temp1 :=" + str(tempDecimal) + ", mac := '" +  macString + "')")
                                ##row = cursor.fetchone()



                        
                connection.commit()

                        
                
                MainLabel.config(text="carga completa")
                ########hacer_request(json_data)#########
                print ("dd")
        else:
                MainLabel.config(text="error al leer archivo")
         

def enviar_db_tecnicos():

        print (excel_path_tecnicos)
        if excel_path_tecnicos != "":
                print("intentando enviar:")
                df = pd.read_excel (excel_path_tecnicos) #for an earlier version of Excel, you may need to use the file extension of 'xls'
                #df = df.drop(df.index[df.eq('').all(axis=1)]) #drop all null rows

                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df['Plan mant.preventivo'] = df['Plan mant.preventivo'].fillna('', inplace=True)


                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df = df.iloc[8:] #no hay que bajar columnas


                connection = psycopg2.connect(
                        host='database-sf.czf4buzwfvge.us-east-2.rds.amazonaws.com',
                        user='SFpostgres',
                        password='sf_admin_fuerza',
                        database='postgres'
                )
                cursor = connection.cursor()


                

              
                for index, row in df.iterrows():
                        print(row[0],row[1],row[2],row[3])

                        #if index != 0:
                        postgres_insert_query = """ INSERT INTO tabla_tecnicos (tt_tecnico_identificador, tt_tecnico_nombre, tt_centro_nombre, tt_centro_id, tt_almacen, tt_tipo_tecnico, tt_rol_tecnico, tt_consecutivo) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"""
                        #record_to_insert = (row[1],row[0],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],0,0)
                        record_to_insert = (row[0], row[1], row[2],row[3],row[4],row[5],row[6],consecutivo)
                        cursor.execute(postgres_insert_query, record_to_insert)
                        ##cursor.execute("Select InsertTest (temp1 :=" + str(tempDecimal) + ", mac := '" +  macString + "')")
                        ##row = cursor.fetchone()
                        
                connection.commit()

                        
                
                MainLabel.config(text="carga completa")
                ########hacer_request(json_data)#########
                print ("dd")
        else:
                MainLabel.config(text="error al leer archivo")
    
def enviar_db_direcciones():

        print (excel_path_direcciones)
        if excel_path_direcciones != "":
                print("intentando enviar:")
                df = pd.read_excel (excel_path_direcciones) #for an earlier version of Excel, you may need to use the file extension of 'xls'
                #df = df.drop(df.index[df.eq('').all(axis=1)]) #drop all null rows

                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df['Plan mant.preventivo'] = df['Plan mant.preventivo'].fillna('', inplace=True)


                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df = df.iloc[8:] #no hay que bajar columnas


                connection = psycopg2.connect(
                        host='database-sf.czf4buzwfvge.us-east-2.rds.amazonaws.com',
                        user='SFpostgres',
                        password='sf_admin_fuerza',
                        database='postgres'
                )
                cursor = connection.cursor()


                for index, row in df.iterrows():

                        if row[0].isdigit():
                                print(row[0],row[1],row[2],row[3])

                                #REVISAR EQUIPO NUMERO 318 DEL ARCHIVO EQUIPO1
                        
                                postgres_insert_query = """ INSERT INTO tabla_direcciones_clientes (dh_cliente_id, dh_pais, dh_nombre_uno, dh_nombre_dos, dh_poblacion, dh_codigo_postal, dh_region, dh_conq_busq, dh_calle, dh_telefono_uno, dh_telefax, dh_direccion) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                                #record_to_insert = (row[1],row[0],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],0,0)
                                record_to_insert = (row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[12]) #la 11 es una columna vacía
                                cursor.execute(postgres_insert_query, record_to_insert)

                                ##cursor.execute("Select InsertTest (temp1 :=" + str(tempDecimal) + ", mac := '" +  macString + "')")
                                ##row = cursor.fetchone()

                        
                connection.commit()

                        
                
                MainLabel.config(text="carga completa")
                ########hacer_request(json_data)#########
                print ("dd")
        else:
                MainLabel.config(text="error al leer archivo")
         

def funcion_enlazar_equipo_direccion():



        connection = psycopg2.connect(
                host='database-sf.czf4buzwfvge.us-east-2.rds.amazonaws.com',
                user='SFpostgres',
                password='sf_admin_fuerza',
                database='postgres'
        )
        cursor = connection.cursor()


        postgres_search_query = """ SELECT * FROM tabla_equipos"""
        #record_to_insert = (row[1],row[0],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],0,0)
        #record_to_insert = (row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[12]) #la 11 es una columna vacía
        cursor.execute(postgres_search_query)

        ##cursor.execute("Select InsertTest (temp1 :=" + str(tempDecimal) + ", mac := '" +  macString + "')")
        ##row = cursor.fetchone()
        
        #rows = cursor.fetchall()
                
        r = [dict((cursor.description[i][0], value) \
                for i, value in enumerate(row)) for row in cursor.fetchall()]
        
        for eq in r:
                denom = eq["e_denominacion"] 
                #print(denom)
                equipoid = eq["e_equipo_id"]

                



                # result = word.find('for')
                # print("Substring 'for ' found at index:", result)
                
                # # How to use find()
                # if (word.find('pawan') != -1):
                # print("Contains given substring ")
                # else:
                # print("Doesn't contains given substring")



                #https://bobbyhadz.com/blog/javascript-extract-number-from-string
                #link para buscar solo el numero

                numcrudo = [int(s) for s in denom.split() if s.isdigit()]

                numbueno = ''
                buscar_equipo_direccion_query1 = ""

                ifbusq = 0

                if(len(numcrudo) > 0):
                        numbueno = str(numcrudo[0])
                        buscar_equipo_direccion_query1 = """SELECT * from __buscar_equipo_en_direcciones_con_numero(%s,%s,%s)"""
                        input_insert = (denom, 0.6,numbueno)
                        ifbusq = 1
                else:
                        t = 1
                        buscar_equipo_direccion_query1 = """SELECT * from __buscar_equipo_en_direcciones(%s,%s)"""
                        input_insert = (denom, 0.6)                        

                #buscar_equipo_direccion_query = """SELECT * from __buscar_equipo_en_direcciones(%s,%s)"""
                #buscar_equipo_direccion_query = """SELECT * from __buscar_equipo_en_direcciones_con_numero(%s,%s,%s)"""





                if ifbusq > 0:
                        cursor.execute(buscar_equipo_direccion_query1, input_insert)

                        r2 = [dict((cursor.description[i][0], value) \
                                for i, value in enumerate(row)) for row in cursor.fetchall()]

                        print("res:") 
                        print(r2)                        


                        if(len(r2) == 0):

                                buscar_equipo_direccion_query3 = """SELECT * from __buscar_equipo_en_direcciones(%s,%s)"""
                                input_insert3 = (denom, 0.6)   
                                cursor.execute(buscar_equipo_direccion_query3, input_insert3)
                                r2 = [dict((cursor.description[i][0], value) \
                                        for i, value in enumerate(row)) for row in cursor.fetchall()]

                                print("res:") 
                                print(r2)      


                else:
                        cursor.execute(buscar_equipo_direccion_query1, input_insert)

                        r2 = [dict((cursor.description[i][0], value) \
                                for i, value in enumerate(row)) for row in cursor.fetchall()]

                        print("res:") 
                        print(r2)



                if(len(r2) > 0):
                        if (r2.count == 1):  #mismo en ambos. solo toma el primero de la lista
                                hijo_id = r2[0]["dh_cliente_id"]
                                #print(hijo_id)
                                editar_equipo_query = """UPDATE tabla_equipos SET e_cliente_hijo_id = %s WHERE e_equipo_id = %s"""
                                input_edit = (hijo_id, equipoid)
                                cursor.execute(editar_equipo_query, input_edit)
                        else:
                                hijo_id = r2[0]["dh_cliente_id"]
                                #print(hijo_id)
                                editar_equipo_query = """UPDATE tabla_equipos SET e_cliente_hijo_id = %s WHERE e_equipo_id = %s"""
                                input_edit = (hijo_id, equipoid)
                                cursor.execute(editar_equipo_query, input_edit)

        connection.commit()
        cursor.connection.close()
                
        
        MainLabel.config(text="carga completa")
        ########hacer_request(json_data)#########
        print ("dd")


def funcion_equipos_solo_lat_long():

        print (excel_path_equipos)
        if excel_path_equipos != "":
                print("intentando enviar:")
                df = pd.read_excel (excel_path_equipos) #for an earlier version of Excel, you may need to use the file extension of 'xls'
                #df = df.drop(df.index[df.eq('').all(axis=1)]) #drop all null rows

                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df['Plan mant.preventivo'] = df['Plan mant.preventivo'].fillna('', inplace=True)


                #df['Pl.MantPrv'] = df['Pl.MantPrv'].fillna('', inplace=True)
                #df = df.iloc[8:] #no hay que bajar columnas


                connection = psycopg2.connect(
                        host='database-sf.czf4buzwfvge.us-east-2.rds.amazonaws.com',
                        user='SFpostgres',
                        password='sf_admin_fuerza',
                        database='postgres'
                )
                cursor = connection.cursor()


                for index, row in df.iterrows():

                        #lat = row.latitude

                        print(row[1])
                        print(row[2])

                        e_latitud = None
                        if math.isnan(row[1]):
                                e_latitud = None
                        else:
                                e_latitud = str(row[1])

                        e_longitud = None
                        if math.isnan(row[2]):
                                e_longitud = None
                        else:
                                e_longitud = str(row[2])


                        #REVISAR EQUIPO NUMERO 318 DEL ARCHIVO EQUIPO1
                        equipoid = 0
                        if isinstance(row[0], str):
                                equipoid = None
                        else:
                                if not math.isnan(row[0]):
                                        #postgres_insert_query = """ INSERT INTO tabla_equipos (e_equipo_id, e_tecnico, e_cliente, e_denominacion, e_latitud, e_longitud, e_frecuencia, e_kw, e_tiempo, e_diaservicio, e_curso, e_dc3, e_extra, e_sys_point) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                                        #postgres_insert_query = """ SELECT * from __insert_equipo (e_equipo_id, e_tecnico, e_cliente, e_denominacion, e_latitud, e_longitud, e_frecuencia, e_kw, e_tiempo, e_diaservicio, e_curso, e_dc3, e_extra) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                                        #postgres_insert_query = """ SELECT * from __insert_equipo (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                                        #postgres_insert_query = """ SELECT * from __insert_or_update_equipo (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""


                                        editar_equipo_query = """UPDATE tabla_equipos SET e_latitud = %s, e_longitud = %s WHERE e_equipo_id = %s"""

                                        #record_to_insert = (row[1],row[0],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],0,0)
                                        #record_to_insert = (row[0],row[1],row[2], row[3], row[4],row[5],row[6],row[7],row[8],row[9],curso,dc3,row[12],None)
                                        #e_latitud = str(row[1])
                                        #e_longitud = str(row[2])

                                        e_equipo_id = int(row[0])

                                        #record_to_insert = (row[0],str(row[1]),row[2], str(row[3]), str(row[4]),str(row[5]),str(row[6]),row[7],row[8],str(row[9]),curso,dc3,str(row[12]))
                                        #record_to_insert = (e_equipo_id,e_tecnico,row[2], str(row[3]), str(row[4]),str(row[5]),str(row[6]),row[7],row[8],str(row[9]),curso,dc3,str(row[12]))
                                        record_to_insert = (e_latitud, 
                                                                e_longitud,
                                                                e_equipo_id)

                                        cursor.execute(editar_equipo_query, record_to_insert)
                                        ##cursor.execute("Select InsertTest (temp1 :=" + str(tempDecimal) + ", mac := '" +  macString + "')")
                                        ##row = cursor.fetchone()



                        
                connection.commit()

                        
                
                MainLabel.config(text="carga completa")
                ########hacer_request(json_data)#########
                print ("dd")
        else:
                MainLabel.config(text="error al leer archivo")
     

def funcion_test_iot():

        dataForm = {
                "User": "NotificationUser",
                "Password": "hN*MSWYFEx4znLVbzg^KhyjDXmam6i",
                "SensorId": 1,
                "Type": 1
        }

        r = requests.post("https://galaxyapidemo.azurewebsites.net//api/PushNotification/AddNotification", data=dataForm)
        print(r.status_code, r.reason)


###para correr programa en terminal
#python sf_excel_postgres.py

root = Tk() #siempre va esto primero para dibujar la ventana main
root.title("Excel a SAP")

canvas = Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack()


background_image = PhotoImage(file='./excel_wallpaper2.png')
background_label = Label(root, image=background_image)
background_label.place(relwidth=1,relheight=1)




#myLabel = tk.Label(root, text="hola mundo") #lavel

#myLabel.pack() #ponerla en la pantalla
button = Button(root, text="Cargar servicios", font=("Arial",15), command=lambda: funcion_abrir())
button.place(relx=0.03, rely=0.08, relwidth=0.15, relheight=.05)

button = Button(root, text="Enviar servicios", font=("Arial",15), command=lambda: enviar_db_servicios())
button.place(relx=0.03, rely=0.13, relwidth=0.15, relheight=.05)


labelServicioUrl = Label(root, text="")
labelServicioUrl.place(relx=0.2, rely=0.08, relheight=.05)



###########
button = Button(root, text="Cargar equipos", font=("Arial",15), command=lambda: funcion_abrir_equipos())
button.place(relx=0.03, rely=0.23, relwidth=0.15, relheight=.05)

button = Button(root, text="Enviar equipos", font=("Arial",15), command=lambda: enviar_db_equipos())
button.place(relx=0.03, rely=0.28, relwidth=0.15, relheight=.05)


labelEquiposUrl = Label(root, text="")
labelEquiposUrl.place(relx=0.2, rely=0.23, relheight=.05)


########## tecnicos
button = Button(root, text="Cargar tecnicos", font=("Arial",15), command=lambda: funcion_abrir_tecnicos())
button.place(relx=0.03, rely=0.38, relwidth=0.15, relheight=.05)

button = Button(root, text="Enviar tecnicos", font=("Arial",15), command=lambda: enviar_db_tecnicos())
button.place(relx=0.03, rely=0.43, relwidth=0.15, relheight=.05)


labelTecnicosUrl = Label(root, text="")
labelTecnicosUrl.place(relx=0.2, rely=0.38, relheight=.05)


########## direcciones
button = Button(root, text="Cargar direcciones", font=("Arial",15), command=lambda: funcion_abrir_direcciones())
button.place(relx=0.03, rely=0.53, relwidth=0.15, relheight=.05)

button = Button(root, text="Enviar direcciones", font=("Arial",15), command=lambda: enviar_db_direcciones())
button.place(relx=0.03, rely=0.58, relwidth=0.15, relheight=.05)


labelDireccionesUrl = Label(root, text="")
labelDireccionesUrl.place(relx=0.2, rely=0.53, relheight=.05)


########## funcion especial
button = Button(root, text="Enlace * equipos direcciones", font=("Arial",15), command=lambda: funcion_enlazar_equipo_direccion())
button.place(relx=0.7, rely=0.23, relwidth=0.25, relheight=.05)


button = Button(root, text="equipos actualizar solo lat y long", font=("Arial",15), command=lambda: funcion_equipos_solo_lat_long())
button.place(relx=0.7, rely=0.38, relwidth=0.25, relheight=.05)


####funcion prueba especial
button = Button(root, text="test", font=("Arial",15), command=lambda: funcion_test_iot())
button.place(relx=0.7, rely=0.53, relwidth=0.25, relheight=.05)




lower_frame = Frame(root, bg='#23693d', bd=10)
lower_frame.place(relx=0.5, rely=0.65, relwidth=0.95, relheight=0.4, anchor='n')

MainLabel = Label(lower_frame, text="Respuesta del servidor:")
MainLabel.place(relwidth = 1, relheight = 1)

root.mainloop()


