from tkinter import * 
from tkinter import scrolledtext as st
from base64 import b64decode
import time
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
import json
from pathlib import Path
from datetime import datetime, timedelta
from tkinter import ttk
from tkinter import *
import sqlite3

msg_alert = []
op = int

#Funciones
def Digiturno10AM():
    # Ejecutar WebDriver
    s = Service("C:\webdrivers\edgedriver.exe")
    driver = webdriver.Edge(service=s)
    driver.maximize_window()
    from datetime import datetime
    fecha = datetime.today().strftime('%d-%m-%y')
    hora_actual = datetime.now().strftime('%H-%M')


    # Creacion de listas para almacenar las variables recogidas
    nombre_punto = []
    datos_caja_rapida = []
    datos_caja_preferencial = []
    datos_caja_general = []
    atencion = []
    pro_espera = []
    pro_atencion = []
    array = []
    contador = 1   

    # Lista de paginas a recoger informacion
    db_name = str(Path.home() / "Documents" / "BD_Puntos" / "Puntos.db")
    conn = sqlite3.connect(db_name)
    # Leer la tabla y guardarla en un DataFrame de pandas
    df = pd.read_sql_query("SELECT * FROM Puntos", conn)
    # Transformar una columna en una lista
    nombre_puntos = df['Nombre'].tolist()
    links_puntos = df['Link'].tolist()
    # Cerrar la conexión a la base de datos
    conn.close()

    for i in range(len(nombre_puntos)):
        try:
            url = str(links_puntos[i])
            driver.get(url)
            time.sleep(1)
            try:
                driver.switch_to.frame(driver.find_element(By.XPATH,"/html/frameset/frameset/frame[3]"))
                login_user = driver.find_element(By.ID,"txtLogin").send_keys("edgar")
                pass_user= driver.find_element(By.ID,"txtClave").send_keys("123")
                
                try:
                    oficina= driver.find_element(By.ID, "txtOficina").click()
                    time.sleep(1)
                    loggin = driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[3]/td[2]").click()
                    driver.switch_to.default_content()
                except:
                    caja = driver.find_element(By.ID, "txtCaja").click()
                    time.sleep(1)
                    loggin = driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[3]/td[2]").click()
                    driver.switch_to.default_content()

                try:
                    nombre_punto.append(nombre_puntos[i])
                    driver.switch_to.frame(driver.find_element(By.XPATH,"/html/frameset/frameset/frame[1]"))
                    ocultar = driver.find_element(By.XPATH,"/html/body/table/tbody/tr[1]/td/font").click()
                    estadistica = driver.find_element(By.XPATH,"/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/u").click()
                    estadistico_servicio = driver.find_element(By.XPATH,"/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/div/table/tbody/tr[2]/td/a").click()
                    driver.switch_to.default_content()
                    driver.switch_to.frame(driver.find_element(By.XPATH, "/html/frameset/frameset/frame[3]"))
                    hora_inicio = driver.find_element(By.ID, "txtHoraInicial")
                    hora_inicio.click()
                    hora_inicio.clear()
                    hora_inicio.send_keys("06:00")
                    hora_final = driver.find_element(By.ID, "txtHoraFinal")
                    hora_final.click()
                    hora_final.clear()
                    hora_final.send_keys("09:30")
                    buscar_digiturno = driver.find_element(By.XPATH, "/html/body/form/table[3]/tbody/tr")
                    buscar_digiturno.click()
                    time.sleep(2)

                    #Crear carpeta de evidencia
                    try:
                        año = datetime.today().strftime("%Y")
                        mes = datetime.today().strftime("%m")
                        dia = datetime.today().strftime("%d")
                        try:
                            evidencia = str(Path.home() / "Documents" / "Reporte Digiturno")
                            os.mkdir(evidencia)
                        except:
                            pass
                        try:
                            evidencia = str(Path.home() / "Documents" / "Reporte Digiturno" / "Evidencia")
                            os.mkdir(evidencia)
                        except:
                            pass
                        try:
                            evidencia = str(Path.home() / "Documents" / "Reporte Digiturno" / "Evidencia" / año)
                            os.mkdir(evidencia)
                        except:
                            pass
                        try:
                            c_screen = str(Path.home() / "Documents" / "Reporte Digiturno" / "Evidencia" / año / mes)
                            os.mkdir(c_screen)
                        except:
                            pass
                        c_screen = str(Path.home() / "Documents" / "Reporte Digiturno" / "Evidencia" / año / mes / dia)
                        os.mkdir(c_screen)
                    except:
                        pass

                    #Almacenar Screen    
                    nombre_foto = f'{str(contador)} _ {str(fecha)}_{str(hora_actual)}'+'.png'
                    ruta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Evidencia" / año / mes / dia / nombre_foto)
                    driver.save_screenshot(ruta)
                    array.append(contador)
                    contador += 1
                    driver.switch_to.default_content()
                    driver.switch_to.frame(driver.find_element(By.XPATH, "/html/frameset/frameset/frame[3]"))

                    try:
                        nombre_pagina = nombre_puntos[i]
                        pagina = driver.find_element(By.XPATH,"/html/body")
                        tabla = driver.find_element(By.XPATH, "/html/body/form/table[2]")
                        
                        # Caja general
                        try:
                            info_tabla=tabla.text
                            info_tabla
                            pos_g = info_tabla.find('GENERAL')
                            general = info_tabla[pos_g:]
                            general = general.split()
                            d_cg=general[5]
                            d_cg=d_cg[1:]
                            datos_caja_general.append(d_cg)
                        except:
                            d_cg="N/A"
                            datos_caja_general.append(d_cg)

                        # Caja rapida
                        try:
                            info_tabla=tabla.text
                            info_tabla
                            pos_r = info_tabla.find('RAPIDA')
                            rapida = info_tabla[pos_r:]
                            rapida = rapida.split()
                            d_cr=rapida[5]
                            d_cr=d_cr[1:]
                            datos_caja_rapida.append(d_cr)
                        except:
                            d_cr="N/A"
                            datos_caja_rapida.append(d_cr)

                        # Caja preferencial
                        try:
                            try:
                                info_tabla=tabla.text
                                pos_p = info_tabla.find('PREFERENCIAL')
                                preferencial = info_tabla[pos_p:]
                                preferencial = preferencial.split()
                                d_cp=preferencial[5]
                                datos_caja_preferencial.append(d_cp)
                            except:
                                info_tabla=tabla.text
                                pos_p = info_tabla.find('PREFENCIAL')
                                preferencial = info_tabla[pos_p:]
                                preferencial = preferencial.split()
                                d_cp=preferencial[5]
                                d_cp=d_cp[1:]
                                datos_caja_preferencial.append(d_cp)
                        except:
                            d_cp="N/A"
                            datos_caja_preferencial.append(d_cp)


                    except:
                        pass
                    
                    try:
                        info = driver.find_element(By.XPATH, "/html/body/form/table[4]/tbody")
                        text_data = info.text
                        pos=text_data.find('Total Servicios')
                        new_data = text_data[pos:]
                        new_data=new_data.split()
                        numero = new_data[2]
                        t_espera = new_data[3]
                        t_espera=t_espera[1:]
                        t_atencion = new_data[4]
                        t_atencion=t_atencion[1:]
                        atencion.append(numero)
                        pro_espera.append(t_espera)
                        pro_atencion.append(t_atencion)
                    except:
                        pass
                        
                except:
                    datos_caja_rapida.append("N/A")
                    datos_caja_preferencial.append("N/A")
                    datos_caja_general.append("N/A")
                    atencion.append("0")
                    pro_espera.append("N/A")
                    pro_atencion.append("N/A")
                    msg_alert.append(f'No se pudo acceder al punto {str(nombre_puntos[i])}')
                    array.append('N/A')
                    pass
                    

            except:
                msg_alert.append(f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
                pass
        except:
            msg_alert.append(f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
            pass
            

    # Cerrar driver
    driver.close()

    #Convertir datos de atencion a número
    atencion = list(map(int,atencion))

    #Crear DataFrame
    df = pd.DataFrame({'Nombre punto':nombre_punto,'Aten.':atencion,'Pro.de Espera':pro_espera,'Pro.Atención':pro_atencion,'Datos Caja General':datos_caja_general,'Datos Caja Rapida':datos_caja_rapida,'Datos Caja Preferencial':datos_caja_preferencial})
    df['Indice foto']=array

    # crear carpeta
    try:
        año = datetime.today().strftime("%Y")
        mes = datetime.today().strftime("%m")
        dia = datetime.today().strftime("%d")
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno 9_30")
            os.mkdir(carpeta)
        except:
            pass
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno 9_30"/ año)
            os.mkdir(carpeta)
        except:
            pass
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno 9_30"/ año/ mes)
            os.mkdir(carpeta)
        except:
            pass
        carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno 9_30"/ año / mes / dia)
        os.mkdir(carpeta)
    except:
        print("Error al crear carpeta")
        msg_alert.append('Error al crear carpeta')
        pass
    #Crear excel de info
    fecha = datetime.today().strftime('%d-%m-%y')
    archivo = f'Puntos Digiturnos '+str(fecha)+'.xlsx'
    ruta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno 9_30" / año / mes / dia / archivo)
    df.to_excel(ruta, sheet_name="Digiturnos",index=0)

def DigiturnoAnterior():
    # Ejecutar WebDriver
    s = Service("C:\webdrivers\msedgedriver.exe")
    driver = webdriver.Edge(service=s)
    driver.maximize_window()

    # Creacion de listas para almacenar las variables recogidas
    nombre_punto = []
    atencion = []
    pro_espera = []
    pro_atencion = []

    # Lista de paginas a recoger informacion
    db_name = str(Path.home() / "Documents" / "BD_Puntos" / "Puntos.db")
    conn = sqlite3.connect(db_name)
    # Leer la tabla y guardarla en un DataFrame de pandas
    df = pd.read_sql_query("SELECT * FROM Puntos", conn)
    # Transformar una columna en una lista
    nombre_puntos = df['Nombre'].tolist()
    links_puntos = df['Link'].tolist()
    # Cerrar la conexión a la base de datos
    conn.close()

    # Seleccionar dia anterior
    from datetime import datetime, timedelta
    now = datetime.now()-timedelta(days=1)
    fecha_ayer = now.strftime('%d/%m/%Y')

    for i in range(len(nombre_puntos)):
        try:
            url = str(links_puntos[i])
            driver.get(url)
            time.sleep(1)
            try:
                driver.switch_to.frame(driver.find_element(
                    By.XPATH, "/html/frameset/frameset/frame[3]"))
                login_user = driver.find_element(By.ID, "txtLogin")
                login_user.send_keys("edgar")
                pass_user = driver.find_element(By.ID, "txtClave")
                pass_user.send_keys("123")

                try:
                    oficina = driver.find_element(By.ID, "txtOficina")
                    oficina.click()
                    time.sleep(1)
                    loggin = driver.find_element(
                        By.XPATH, "/html/body/form/table/tbody/tr[3]/td[2]")
                    loggin.click()
                    driver.switch_to.default_content()
                except:
                    caja = driver.find_element(By.ID, "txtCaja")
                    caja.click()
                    time.sleep(1)
                    loggin = driver.find_element(
                        By.XPATH, "/html/body/form/table/tbody/tr[3]/td[2]").click()
                    driver.switch_to.default_content()

                try:
                    nombre_punto.append(nombre_puntos[i])
                    driver.switch_to.frame(driver.find_element(
                        By.XPATH, "/html/frameset/frameset/frame[1]"))
                    ocultar = driver.find_element(By.XPATH,"/html/body/table/tbody/tr[1]/td/font").click()
                    estadistica = driver.find_element(
                        By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/u")
                    estadistica.click()
                    estadistico_servicio = driver.find_element(
                        By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/div/table/tbody/tr[2]/td/a")
                    estadistico_servicio.click()
                    driver.switch_to.default_content()
                    driver.switch_to.frame(driver.find_element(
                        By.XPATH, "/html/frameset/frameset/frame[3]"))
                    hora_inicio = driver.find_element(By.ID, "txtHoraInicial")
                    hora_inicio.click()
                    hora_inicio.clear()
                    hora_inicio.send_keys("06:00")
                    hora_final = driver.find_element(By.ID, "txtHoraFinal")
                    hora_final.click()
                    hora_final.clear()
                    hora_final.send_keys("22:00")
                    fecha_inicio = driver.find_element(By.ID, "dtpFecha")
                    fecha_inicio.click()
                    fecha_inicio.clear()
                    fecha_inicio.send_keys(str(fecha_ayer))
                    fecha_final = driver.find_element(By.ID, "dtpFechaF")
                    fecha_final.click()
                    fecha_final.clear()
                    fecha_final.send_keys(str(fecha_ayer))
                    buscar_digiturno = driver.find_element(
                        By.XPATH, "/html/body/form/table[3]/tbody/tr")
                    buscar_digiturno.click()
                    time.sleep(5)
                    driver.switch_to.default_content()
                    driver.switch_to.frame(driver.find_element(
                        By.XPATH, "/html/frameset/frameset/frame[3]"))

                    try:
                        info = driver.find_element(
                            By.XPATH, "/html/body/form/table[4]/tbody")
                        text_data = info.text
                        pos = text_data.find('Total Servicios')
                        new_data = text_data[pos:]
                        new_data = new_data.split()
                        numero = new_data[2]
                        t_espera = new_data[3]
                        t_espera = t_espera[1:]
                        t_atencion = new_data[4]
                        t_atencion = t_atencion[1:]
                        atencion.append(numero)
                        pro_espera.append(t_espera)
                        pro_atencion.append(t_atencion)
                    except:
                        pass

                except:
                    atencion.append("0")
                    pro_espera.append("N/A")
                    pro_atencion.append("N/A")
                    print(
                        f'No se pudo acceder al punto {str(nombre_puntos[i])}')
                    msg_alert.append(
                        f'No se pudo acceder al punto {str(nombre_puntos[i])}')
                    pass

            except:
                print(
                    f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
                msg_alert.append(
                    f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
                pass

        except:
            print(
                f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
            msg_alert.append(
                f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
            pass

    # Cerrar driver
    driver.close()

    # Convertir datos de atencion a número
    atencion = list(map(int, atencion))

    # Crear DataFrame
    df = pd.DataFrame({'Nombre punto': nombre_punto, 'Aten.': atencion,
                      'Pro.de Espera': pro_espera, 'Pro.Atención': pro_atencion})
    
    # crear carpeta
    fecha_a = datetime.now()-timedelta(1)
    año = fecha_a.strftime("%Y")
    mes_a = fecha_a.strftime('%m')
    dia_a = fecha_a.strftime('%d')
    try:
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno")
            os.mkdir(carpeta)
        except:
            pass
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior")
            os.mkdir(carpeta)
        except:
            pass
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior" / año)
            os.mkdir(carpeta)
        except:
            pass
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior"/ año / mes_a)
            os.mkdir(carpeta)
        except:
            pass
        carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior"/ año / mes_a / dia_a)
        os.mkdir(carpeta)
    except:
        print('Error al crear carpeta')
        msg_alert.append('Error al crear carpeta')
        pass

    # Crear excel de info
    now = datetime.now()-timedelta(days=1)
    fecha_ayer = now.strftime('%d-%m-%Y')
    archivo = f'Puntos Digiturnos dia anterior '+str(fecha_ayer) +'.xlsx'
    ruta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior" / año / mes_a / dia_a / archivo)
    df.to_excel(ruta, sheet_name="Digiturnos", index=0)

def DigiturnoFinde():
    # Ejecutar WebDriver
    s = Service("C:\webdrivers\msedgedriver.exe")
    driver = webdriver.Edge(service=s)
    driver.maximize_window()

    # Creacion de listas para almacenar las variables recogidas
    nombre_punto = []
    atencion = []
    pro_espera = []
    pro_atencion = []

    # Lista de paginas a recoger informacion
    db_name = str(Path.home() / "Documents" / "BD_Puntos" / "Puntos.db")
    conn = sqlite3.connect(db_name)
    # Leer la tabla y guardarla en un DataFrame de pandas
    df = pd.read_sql_query("SELECT * FROM Puntos", conn)
    # Transformar una columna en una lista
    nombre_puntos = df['Nombre'].tolist()
    links_puntos = df['Link'].tolist()
    # Cerrar la conexión a la base de datos
    conn.close()

    # Seleccionar dia anterior
    from datetime import datetime, timedelta
    now = datetime.now()-timedelta(days=op)
    fecha_ayer = now.strftime('%d/%m/%Y')

    for i in range(len(nombre_puntos)):
        try:
            url = str(links_puntos[i])
            driver.get(url)
            time.sleep(1)
            try:
                driver.switch_to.frame(driver.find_element(
                    By.XPATH, "/html/frameset/frameset/frame[3]"))
                login_user = driver.find_element(By.ID, "txtLogin")
                login_user.send_keys("edgar")
                pass_user = driver.find_element(By.ID, "txtClave")
                pass_user.send_keys("123")

                try:
                    oficina = driver.find_element(By.ID, "txtOficina")
                    oficina.click()
                    time.sleep(1)
                    loggin = driver.find_element(
                        By.XPATH, "/html/body/form/table/tbody/tr[3]/td[2]")
                    loggin.click()
                    driver.switch_to.default_content()
                except:
                    caja = driver.find_element(By.ID, "txtCaja")
                    caja.click()
                    time.sleep(1)
                    loggin = driver.find_element(
                        By.XPATH, "/html/body/form/table/tbody/tr[3]/td[2]").click()
                    driver.switch_to.default_content()

                try:
                    nombre_punto.append(nombre_puntos[i])
                    driver.switch_to.frame(driver.find_element(
                        By.XPATH, "/html/frameset/frameset/frame[1]"))
                    ocultar = driver.find_element(By.XPATH,"/html/body/table/tbody/tr[1]/td/font").click()
                    estadistica = driver.find_element(
                        By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/u")
                    estadistica.click()
                    estadistico_servicio = driver.find_element(
                        By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/div/table/tbody/tr[2]/td/a")
                    estadistico_servicio.click()
                    driver.switch_to.default_content()
                    driver.switch_to.frame(driver.find_element(
                        By.XPATH, "/html/frameset/frameset/frame[3]"))
                    hora_inicio = driver.find_element(By.ID, "txtHoraInicial")
                    hora_inicio.click()
                    hora_inicio.clear()
                    hora_inicio.send_keys("06:00")
                    hora_final = driver.find_element(By.ID, "txtHoraFinal")
                    hora_final.click()
                    hora_final.clear()
                    hora_final.send_keys("22:00")
                    fecha_inicio = driver.find_element(By.ID, "dtpFecha")
                    fecha_inicio.click()
                    fecha_inicio.clear()
                    fecha_inicio.send_keys(str(fecha_ayer))
                    fecha_final = driver.find_element(By.ID, "dtpFechaF")
                    fecha_final.click()
                    fecha_final.clear()
                    fecha_final.send_keys(str(fecha_ayer))
                    buscar_digiturno = driver.find_element(
                        By.XPATH, "/html/body/form/table[3]/tbody/tr")
                    buscar_digiturno.click()
                    time.sleep(5)
                    driver.switch_to.default_content()
                    driver.switch_to.frame(driver.find_element(
                        By.XPATH, "/html/frameset/frameset/frame[3]"))

                    try:
                        info = driver.find_element(
                            By.XPATH, "/html/body/form/table[4]/tbody")
                        text_data = info.text
                        pos = text_data.find('Total Servicios')
                        new_data = text_data[pos:]
                        new_data = new_data.split()
                        numero = new_data[2]
                        t_espera = new_data[3]
                        t_espera = t_espera[1:]
                        t_atencion = new_data[4]
                        t_atencion = t_atencion[1:]
                        atencion.append(numero)
                        pro_espera.append(t_espera)
                        pro_atencion.append(t_atencion)
                    except:
                        pass

                except:
                    atencion.append("0")
                    pro_espera.append("N/A")
                    pro_atencion.append("N/A")
                    print(
                        f'No se pudo acceder al punto {str(nombre_puntos[i])}')
                    msg_alert.append(
                        f'No se pudo acceder al punto {str(nombre_puntos[i])}')
                    pass

            except:
                print(
                    f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
                msg_alert.append(
                    f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
                pass

        except:
            print(
                f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
            msg_alert.append(
                f'No se pudo acceder al aplicativo de {str(nombre_puntos[i])}')
            pass

    # Cerrar driver
    driver.close()

    # Convertir datos de atencion a número
    atencion = list(map(int, atencion))

    # Crear DataFrame
    df = pd.DataFrame({'Nombre punto': nombre_punto, 'Aten.': atencion,
                      'Pro.de Espera': pro_espera, 'Pro.Atención': pro_atencion})
    
    print(f'Puntos: {len(nombre_punto)}')
    print(f'N atendidos: {len(atencion)}')
    print(f'T - Espera: {len(pro_espera)}')
    print(f'T - Tencion: {len(pro_atencion)}')
    
    # crear carpeta
    fecha_a = datetime.now()-timedelta(op)
    año = fecha_a.strftime('%Y')
    mes_a = fecha_a.strftime('%m')
    dia_a = fecha_a.strftime('%d')
    try:
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno")
            os.mkdir(carpeta)
        except:
            pass
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior")
            os.mkdir(carpeta)
        except:
            pass
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior" / año)
            os.mkdir(carpeta)
        except:
            pass
        try:
            carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior"/ año / mes_a)
            os.mkdir(carpeta)
        except:
            pass
        carpeta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior"/ año / mes_a / dia_a)
        os.mkdir(carpeta)
    except:
        print('Error al crear carpeta')
        msg_alert.append('Error al crear carpeta')
        pass

    # Crear excel de info
    now = datetime.now()-timedelta(days=op)
    fecha_ayer = now.strftime('%d-%m-%Y')
    archivo = f'Puntos Digiturnos dia anterior '+str(fecha_ayer) +'.xlsx'
    ruta = str(Path.home() / "Documents" / "Reporte Digiturno" / "Digiturno Dia Anterior" / año / mes_a / dia_a / archivo)
    df.to_excel(ruta, sheet_name="Digiturnos", index=0)


# Crear interfaz grafica
root = Tk()
root.title("Digiturnos Unificado")
root.geometry("400x520")
root.resizable(0,0)
root['bg']="#19376D"

# Vistas 
def Vista10Am():   
    Menu_op()
    cargar_imagen()
    #Etiquetas y botones
    etiqueta1 = Label(root, text="Informacion Digiturnos 9:30AM - Subdireccion de medicamentos", font=("Arial",8, "bold"),foreground='white')
    etiqueta1.place(x=20,y=210)
    etiqueta1['bg']='#19376D'

    btn1 = Button(root, text="Extraer informacion de digiturnos",command=Digiturno10AM,bg="#0B2447",foreground="white",borderwidth=4)
    btn1.place(x=100,y=250)

    #Crear lista de elementos
    lista_elementos = Listbox(root,width=56, height=10)

    #Ubicamos la lista de elementos
    lista_elementos.place(x=30, y=290)

    #Añadir elementos a lista
    def mostrar_errores():
        for i in range(len(msg_alert)):
            lista_elementos.insert(END, msg_alert[i])

    def LimpiarError():
        lista_elementos.delete(0,END)
        msg_alert.clear()

    btn2 = Button(root, text="Mostrar errores", command=mostrar_errores,
              bg="#FC2947", foreground="white", borderwidth=4)
    btn2.place(x=80, y=463)

    btn3 = Button(root, text="Limpiar errores", command=LimpiarError,
              bg="#379237", foreground="white", borderwidth=4)
    btn3.place(x=220, y=463)

def VistaDiaAnterior():
    Menu_op()
    cargar_imagen()
    etiqueta2 = Label(root, text="Informacion Digiturnos Día Anterior - Subdireccion de medicamentos", font=("Arial", 8, "bold"),foreground='white')
    etiqueta2.place(x=5, y=210)
    etiqueta2['bg']='#19376D'

    btn_e = Button(root, text="Extraer informacion de digiturnos",command=DigiturnoAnterior, bg="#0B2447", foreground="white", borderwidth=4)
    btn_e.place(x=100, y=250)

    # Crear lista de elementos
    lista_elementos = Listbox(root, width=56, height=10)

    # Ubicamos la lista de elementos
    lista_elementos.place(x=30, y=290)

    # Añadir elementos a lista

    def mostrar_errores():
        for i in range(len(msg_alert)):
            lista_elementos.insert(END, msg_alert[i])

    def LimpiarError():
        lista_elementos.delete(0,END)
        msg_alert.clear()

    btn2 = Button(root, text="Mostrar errores", command=mostrar_errores,
              bg="#FC2947", foreground="white", borderwidth=4)
    btn2.place(x=80, y=463)

    btn3 = Button(root, text="Limpiar errores", command=LimpiarError,
              bg="#379237", foreground="white", borderwidth=4)
    btn3.place(x=220, y=463)

def VistaEditar():
    Menu_op()


    db_name = str(Path.home() / "Documents" / "BD_Puntos" / "Puntos.db")
    # Function to Execute Database Querys
    def run_query(query, parameters = ()):
        with sqlite3.connect(db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parameters)
            conn.commit()
        return result

    # Obtener puntos de la base de datos
    def get_puntos():
        # cleaning Table 
        records = tree.get_children()
        for element in records:
            tree.delete(element)
        # getting data
        query = 'SELECT * FROM Puntos ORDER BY ID DESC'
        db_rows = run_query(query)
        # filling data
        for row in db_rows:
            tree.insert('', 0, text = row[1], values = row[2])

    # User Input Validation
    def validation():
        return len(nombre.get()) != 0 and len(link.get()) != 0

    # Agregar punto
    def add_punto():
        if validation():
            query = 'INSERT INTO Puntos VALUES(NULL, ?, ?)'
            parameters =  (nombre.get(), link.get())
            run_query(query, parameters)
            message['text'] = 'Punto {} agregado con exito!'.format(nombre.get())
            nombre.delete(0, END)
            link.delete(0, END)
        else:
            message['text'] = 'Nombre y link requerido!'
        get_puntos()

    # Eliminar punto
    def delete_punto():
        message['text'] = ''
        try:
           tree.item(tree.selection())['text'][0]
        except IndexError as e:
            message['text'] = 'Seleccione un punto'
            return
        message['text'] = ''
        nombre = tree.item(tree.selection())['text']
        query = 'DELETE FROM Puntos WHERE Nombre = ?'
        run_query(query, (nombre, ))
        message['text'] = 'Punto {} eliminado!'.format(nombre)
        get_puntos()

    # Funcion para editar puntos
    def edit_records(new_nombre, nombre , new_link, old_link):
        query = 'UPDATE Puntos SET Nombre = ?, Link = ? WHERE Nombre = ? AND Link = ?'
        parameters = (new_nombre, new_link, nombre, old_link)
        run_query(query, parameters)
        message['text'] = 'Punto {} actualizado con exito'.format(new_nombre)
        get_puntos()

    def edit_punto():
        message['text'] = ''
        try:
            tree.item(tree.selection())['values'][0]
        except IndexError as e:
            message['text'] = 'Seleccione un punto'
            return
        nombre = tree.item(tree.selection())['text']
        old_link = tree.item(tree.selection())['values'][0]
        edit_wind = Toplevel()
        edit_wind.geometry("450x150")
        edit_wind.title = 'Edit Product'
        #Nombre Viejo
        Label(edit_wind, text = 'Nombre Viejo:').grid(row = 0, column = 1)
        Entry(edit_wind, textvariable = StringVar(edit_wind, value = nombre), state = 'readonly',width=56).grid(row = 0, column = 2)
        #Nuevo Nombre
        Label(edit_wind, text = 'Nombre Nuevo:').grid(row = 1, column = 1)
        new_nombre = Entry(edit_wind,width=56)
        new_nombre.grid(row = 1, column = 2)

        #Link Viejo
        Label(edit_wind, text = 'Link Viejo:').grid(row = 2, column = 1)
        Entry(edit_wind, textvariable = StringVar(edit_wind, value = old_link), state = 'readonly',width=56).grid(row = 2, column = 2)
        #Link Nuevo
        Label(edit_wind, text = 'Link Nuevo:').grid(row = 3, column = 1)
        new_link= Entry(edit_wind,width=56)
        new_link.grid(row = 3, column = 2)

        def destruir():
            edit_wind.destroy()

        boton_actualizar = Button(edit_wind, bg="green",text = 'Actualizar', command = lambda: [edit_records(new_nombre.get(), nombre, new_link.get(), old_link),destruir()])
        boton_actualizar.place(x=200,y=100)
        edit_wind.mainloop()


    frame = LabelFrame(root, text = 'Registrar nuevo punto')
    frame.place(x=0,y=30)

    #Nombre Input
    Label(frame, text = 'Nombre: ').grid(row = 1, column = 0)
    nombre = Entry(frame,width=56)
    nombre.focus()
    nombre.grid(row = 1, column = 1)

    #Link Input
    Label(frame, text = 'Link: ').grid(row = 2, column = 0)
    link = Entry(frame,width=56)
    link.grid(row = 2, column = 1)

    # Button Add punto
    ttk.Button(frame, text = 'Guardar Punto',command = add_punto).grid(row = 3, columnspan = 2, sticky = W + E)

    # Output Messages 
    message = Label(text = '', fg = 'white',bg="#19376D")
    message.place(x=80,y=150)

    # Table
    tree = ttk.Treeview(height = 10, columns = 2)
    tree.place(x=0,y=200)
    tree.heading('#0', text = 'Nombre', anchor = CENTER)
    tree.heading('#1', text = 'Link', anchor = CENTER)

    # Filling the Rows
    get_puntos()

    # Buttons
    boton_cargar = Button(text = 'ELIMINAR PUNTO',command=delete_punto)
    boton_cargar.place(x=50,y=450)
    boton_actualizar = Button(text = 'ACTUALIZAR PUNTO',command=edit_punto)
    boton_actualizar.place(x=250,y=450)


def VistaFinde():
    Menu_op()
    cargar_imagen()
    etiqueta2 = Label(root, text="Informacion Digiturnos Día Anterior - Subdireccion de medicamentos", font=("Arial", 8, "bold"),foreground='white')
    etiqueta2.place(x=5, y=180)
    etiqueta2['bg']='#19376D'

    btn_e = Button(root, text="Extraer informacion de digiturnos",command=DigiturnoFinde, bg="#0B2447", foreground="white", borderwidth=4)
    btn_e.place(x=100, y=250)

    # Crear lista de elementos
    lista_elementos = Listbox(root, width=56, height=10)

    # Ubicamos la lista de elementos
    lista_elementos.place(x=30, y=290)

    # Añadir elementos a lista

    def mostrar_errores():
        for i in range(len(msg_alert)):
            lista_elementos.insert(END, msg_alert[i])

    def LimpiarError():
        lista_elementos.delete(0,END)
        msg_alert.clear()

    btn2 = Button(root, text="Mostrar errores", command=mostrar_errores,
              bg="#FC2947", foreground="white", borderwidth=4)
    btn2.place(x=80, y=463)

    btn3 = Button(root, text="Limpiar errores", command=LimpiarError,
              bg="#379237", foreground="white", borderwidth=4)
    btn3.place(x=220, y=463)

    # Funcion elegir dia
    def seleccion():
        global op
        if control1.get() == 1:
            var = 3
            op = var
        elif control2.get() == 1:
            var = 2
            op = var

    control1 = IntVar()
    control2 = IntVar()

    opcion_1 = Checkbutton(root, text="Viernes", variable=control1,bg='#A5D7E8')
    opcion_1.place(x=60, y=210)
    opcion_1.deselect()

    opcion_2 = Checkbutton(root, text="Sabado", variable=control2,bg='#A5D7E8')
    opcion_2.place(x=150, y=210)
    opcion_2.deselect()

    muestra_seleccion = Button(root, text="Seleccionar Día", command=seleccion,bg="#FEDB39").place(x=250, y=210)

def VistaFindeFes():
    Menu_op()
    cargar_imagen()
    etiqueta2 = Label(root, text="Informacion Digiturnos Día Anterior - Subdireccion de medicamentos", font=("Arial", 8, "bold"),foreground='white')
    etiqueta2.place(x=5, y=180)
    etiqueta2['bg']='#19376D'

    btn_e = Button(root, text="Extraer informacion de digiturnos",command=DigiturnoFinde, bg="#0B2447", foreground="white", borderwidth=4)
    btn_e.place(x=100, y=250)

    # Crear lista de elementos
    lista_elementos = Listbox(root, width=56, height=10)

    # Ubicamos la lista de elementos
    lista_elementos.place(x=30, y=290)

    # Añadir elementos a lista

    def mostrar_errores():
        for i in range(len(msg_alert)):
            lista_elementos.insert(END, msg_alert[i])

    def LimpiarError():
        lista_elementos.delete(0,END)
        msg_alert.clear()

    btn2 = Button(root, text="Mostrar errores", command=mostrar_errores,
              bg="#FC2947", foreground="white", borderwidth=4)
    btn2.place(x=80, y=463)

    btn3 = Button(root, text="Limpiar errores", command=LimpiarError,
              bg="#379237", foreground="white", borderwidth=4)
    btn3.place(x=220, y=463)

    # Funcion elegir dia
    def seleccion():
        global op
        if control1.get() == 1:
            var = 4
            op = var
        elif control2.get() == 1:
            var = 3
            op = var

    control1 = IntVar()
    control2 = IntVar()

    opcion_1 = Checkbutton(root, text="Viernes", variable=control1,bg='#A5D7E8')
    opcion_1.place(x=60, y=210)
    opcion_1.deselect()

    opcion_2 = Checkbutton(root, text="Sabado", variable=control2,bg='#A5D7E8')
    opcion_2.place(x=150, y=210)
    opcion_2.deselect()

    muestra_seleccion = Button(root, text="Seleccionar Día", command=seleccion,bg="#FEDB39").place(x=250, y=210)

def VistaDiaAtipico():
    Menu_op()
    cargar_imagen()
    etiqueta2 = Label(root, text="Informacion Digiturnos Día Anterior - Subdireccion de medicamentos", font=("Arial", 8, "bold"),foreground='white')
    etiqueta2.place(x=5, y=180)
    etiqueta2['bg']='#19376D'

    btn_e = Button(root, text="Extraer informacion de digiturnos",command=DigiturnoFinde, bg="#0B2447", foreground="white", borderwidth=4)
    btn_e.place(x=30, y=250)

    # Crear lista de elementos
    lista_elementos = Listbox(root, width=56, height=10)

    # Ubicamos la lista de elementos
    lista_elementos.place(x=30, y=290)

    # Añadir elementos a lista

    def mostrar_errores():
        for i in range(len(msg_alert)):
            lista_elementos.insert(END, msg_alert[i])

    def LimpiarError():
        lista_elementos.delete(0,END)
        msg_alert.clear()

    btn2 = Button(root, text="Mostrar errores", command=mostrar_errores,
              bg="#FC2947", foreground="white", borderwidth=4)
    btn2.place(x=80, y=463)

    btn3 = Button(root, text="Limpiar errores", command=LimpiarError,
              bg="#379237", foreground="white", borderwidth=4)
    btn3.place(x=220, y=463)

    # Funcion elegir dia
    op = 0

    def asignar():
        global op
        op = entrada.get()
        op = int(op)

    def mostrar_aviso():
        messagebox.showinfo("Aviso", "Para saber que número asignar se deben contar los dias desde el dia anterior a la fecha actual hasta la fecha deseada.\nEjemplo: si hoy es lunes 27 y quiero asignar los dias para el viernes 24, cuento desde el domingo hasta el viernes, obteniendo 3 dias (domingo+sabado+viernes) ")


    entrada = Entry(root,width=30)
    entrada.place(x=30,y=215)

    muestra_seleccion = Button(root, text="Seleccionar Día", command=asignar,bg="#FEDB39").place(x=250, y=210)

    boton = Button(root, text="Ayuda con seleccion", command=mostrar_aviso,bg="#F806CC",foreground="white", borderwidth=4)
    boton.place(x=245, y=250)

def destruir_vista():
     # eliminar cualquier contenido antiguo
    for widget in root.winfo_children():
        widget.destroy()

def mostrar_ruta():
        ruta = str(Path.home() / "Documents")
        var = ""
        if os.path.isdir(ruta):
            var = "La carpeta existe"
        else:
            var = "La carpeta no existe"

        messagebox.showinfo("Ruta de links", ruta)
        messagebox.showinfo("Existe la carpeta?", var)

def Menu_op():
    # Crear Menu
    MenuBar = Menu(root)
    root.config(menu=MenuBar)

    # Opciones Digiturnos En Menu
    archivoMenu = Menu(MenuBar,tearoff=0)
    archivoMenu.add_command(label="Digiturnos 9:30",command=lambda:[destruir_vista(),Vista10Am()])
    archivoMenu.add_separator()
    archivoMenu.add_command(label="Digiturnos Dia Anterior",command=lambda:[destruir_vista(),VistaDiaAnterior()])
    archivoMenu.add_separator()
    archivoMenu.add_command(label="Digiturnos Dia Anterior Fin De Semana",command=lambda:[destruir_vista(),VistaFinde()])
    archivoMenu.add_separator()
    archivoMenu.add_command(label="Digiturnos Dia Anterior Fin De Semana festivo", command=lambda:[destruir_vista(),VistaFindeFes()])
    archivoMenu.add_separator()
    archivoMenu.add_command(label="Digiturnos Dia Anterior Dia Atipico", command=lambda:[destruir_vista(),VistaDiaAtipico()])

    # Opcion editar puntos
    editarMenu = Menu(MenuBar,tearoff=0)
    editarMenu.add_command(label="Editar puntos",command=lambda:[destruir_vista(),VistaEditar()])
    editarMenu.add_command(label="¿Donde ubicar carpeta de puntos?",command=mostrar_ruta)

    #Opcion Salida En Menu
    salirMenu = Menu(MenuBar,tearoff=0)
    salirMenu.add_command(label="Salir",command=root.quit)

    # Agregar menus
    MenuBar.add_cascade(label="Digiturnos",menu=archivoMenu)
    MenuBar.add_cascade(label="Editar",menu=editarMenu)
    MenuBar.add_cascade(label="Salir",menu=salirMenu)

# Menu de opciones
Menu_op()

raw64 = "iVBORw0KGgoAAAANSUhEUgAAASwAAACWCAMAAABThUXgAAADAFBMVEUnuu1AyPVFzPdRy/ZczfZlzvdu0Pd30vd+1PiE1viJ1/iN2PmR2fiV2vmY2/mb3Pmc3Pkja7IAuO8iK4klNZCh3vil4Pmo4PoAufE3vu9HwvBSx/JczPQms+kAseoAneEAn+IAoeMAouQAouUAo+YApecApugAqewArO4ApegAnN8Am94AnNsAmt0AmNwAmNsAltkAktYAkdIAjdIAi84Aic8Ah8sAiM4AgscBf8UAfsUAfMQAe8IAecEAd78Adb0Ac7wAcroXQZcnMo4AqOoAqesAq+0AnuEAneAAmdwAl9sAldoAlNgAk9gAk9cAktYAkNUAj9QAjtMAjdIAjNEAi9EAitAAh80AhswAhcsAg8oAgcgAgMcAeMAYRJoAkdUAg8kAgskArOcAqOMAcLkAb7gAcboAbrcApOEAm+IOqOdgxe8AbLXR5fTp9Pv6+v3////9/f4Aod674fT2+f3d7fgAbbcAarQAabMxodqys9UAaLLx8fg0PJNdYqgSLosAZ7HT1Oe/wN24udja2uttcbAaH4ERVqcAZ7KVl8U9RJgoM48dJoYZJIUgLotES5t1ebUjM40AZrHNzeTr7PVUW6QqK4gKcLsAZbCKjcAmMIuqrdLm5vIPYK8AntwWIYMoLYkfOJAcPZQoLooAZK8AkNgPZbImMY0JarYAh9EcnNoipN5bteNwueKBx+pGq98Qm9x/groTSJwAYq4AYa2Yzuzz9foSTaAQltdOU58Am9pkaawAYK0QjdAbTaIAl9er1u+cnspDg8AfkdERRJkxldAjJ4YLWqkvNY4Acb0UhcgAX6vg4e/I3/FGntO81+yjxuTHyOGRv+EihMWBuN4AXqtfp9YAXapTlcoAXaoPUqOlps4LN5EJZLEAXKoJXasVeb4AW6kAWqgzisUAWacNF34AWKYAWKaCl8gqe7wXP5YAV6YAVqVons8AVaQWb7YHTJ4IV6YAVqQAUqIFTp95qdQAU6MAVKMETZ8AU6IAUaECT6ADTp8AUKAAUaEFTZ9pRtuoAAABAHRSTlP7///////////////////////////////9+/v6+vv//f/////////8//////z////9/f/9///9//3////////////9/////////////////////////////////////v/////////////////////////////////////////////////////////////////////////////////////////////////////////////////9//////////////////////////////////////////////3///////////////////3////9/////////////////f////////39Lk4CVwAAL3BJREFUeAGs0wWSxDAMRFFbYR6m+x90p1tKZC/TCxQN/uqEEElECiqpgrqqoakb1aqOejU8jZsJZrekdh9bUrObYFwN0KuOWtWY2lRQQkFSRIl/JbIPB60lUQpBK+aqrJbSUjigY62+ex1rslQ4aNHj61ZeC2/gjSaLNSW1eq3VJbUaHFBrrqpmq4qx1lx/jiVxfwwHrSXony+r3nalubwVWCsYkQsm+vmwLJYHW6E9Yk0+rWxYXbqsmqm8FqzDivIPrU7hgFqSlnqzKcoDZc+clTibCw53/YlLAp9krKf2Y770odRwyqN5MCV/glaMdbjdw9Mdh3vcH6+90FIXuNHDQBiG77DiH0LLzMwcKjPT/e/Qb7yziVsr5b4GVdS4j0b9S0vpH+1/v5j45fiGGj1IDU+Xwl8l+r9ZfNH+TGQFLFhpum6YpoHFWShuJeIJlMRKJLkUlUyn0j9UBgtbKYWVFSebwsGVpOgtVBzvMxGerJt6WI7K63qBVhGrVK6UK0FVbKlatV6v1bkGdqPZardEHarb6fX6g+FoPJlMZ3OBhTSdvossYoITgpRFVBQeu2mRSuMsti1ft3qzzFutwpbcIr3AWaYX+IH0UqJsNpskr00WZVIGp6/1tW3buRzoivli0XFKpVLZdV3P8yp+Bcff2dnd3avu7e/vH6DDw6Ojev24fnJycnp62myfnZ2fX1xcXl5dda47XWjdgOt2Hk6WQZnimBSLERkuYZZFKSmVa/Uxu7eoFP4FlUKhFOZLhpKl4LSGFFFRgsoJqCgfgQpYd4LqkDpCx8fHp1urC3T/cHV1dd3tEdY4wNJisUd9jYwwBmMyFIipYCz2du+ZKU4MxVQLiSoRTUXJVBursmxFY8VUL61orOS5gtU1rICFwdpiaZoGLButOVnMEgVg6CXYAjsskkoxU1Ol3qBiq1BKpWIrGiv0ioqsVCrZ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"
img = PhotoImage(data=raw64)

# Cargar Imagen
def cargar_imagen():
    label_imagen = Label(root,image=img)
    label_imagen.pack()
    label_imagen.place(x=50,y=20)

cargar_imagen()

etiqueta_info = Label(root, text="Jorge Andrés Guerrero Navarro - Subdireccion de medicamentos  \n \n Universidad de Córdoba - Ingeniería Industrial \n \n 2023", font=("Arial", 9, "bold"),foreground='white')
etiqueta_info.place(x=10, y=250)
etiqueta_info['bg']='#19376D'

icono64 = "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"
icono = PhotoImage(data=icono64)

root.iconphoto(True, icono)



root.mainloop()

# Jorge Andres Guerrero Navarro