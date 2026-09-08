# -*- coding: latin-1 -*-
import numpy as np
import math
import pandas as pd
from datetime import datetime
import pytz
import smtplib
from email.message import EmailMessage
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import PatternFill
from openpyxl.styles import Font


def ejecutar_cruce_seguro(CONTADORES,TRAFICO,SITIOS,PQRSD,DISPOSITIVOS):

    matricula1_velocidad_bajada_mostrar = 26.40
    matricula1_velocidad_bajada = 26400

    matricula1_velocidad_subida_mostrar = 6.60
    matricula1_velocidad_subida = 6600

    matricula2_velocidad_bajada_mostrar = 28.50
    matricula2_velocidad_bajada = 28500

    matricula2_velocidad_subida_mostrar = 7.13
    matricula2_velocidad_subida = 7130

    matricula3_velocidad_bajada_mostrar = 32.00
    matricula3_velocidad_bajada = 32000

    matricula3_velocidad_subida_mostrar = 8.00
    matricula3_velocidad_subida = 8000

    matricula4_velocidad_bajada_mostrar = 41.90
    matricula4_velocidad_bajada = 41900

    matricula4_velocidad_subida_mostrar = 10.48
    matricula4_velocidad_subida = 10480

    dataframe_1=dataframe[['Numero de contrato','Departamento','Ciudad','Identificador beneficiario','Grupo','Perfil','Zona','Velocidad de subida [Kbps]','Velocidad de bajada [Kbps]','Fecha de programacion','Fecha de ejecucion','Rango','Duracion de la prueba [s]','Dispositivo','Tipo de solucion','Tipo de centro digital','Estado de la prueba','Identificador de la prueba','Centro poblado','Dane institucion educativa','Tipo']]
    # Filtrar solo la columna 'Identificador beneficiario' y 'estado' de dataframe2
    dataframe_2 = dataframe2[['Id_Beneficiario', 'Estado']]
    # Realizar el inner join
    dataframe_1 = pd.merge(dataframe_1, dataframe_2, left_on='Identificador beneficiario', right_on='Id_Beneficiario', how='inner')

    #filtros
    dataframe_1=dataframe_1[['Numero de contrato','Departamento','Ciudad','Identificador beneficiario','Grupo','Perfil','Zona','Velocidad de subida [Kbps]','Velocidad de bajada [Kbps]','Fecha de programacion','Fecha de ejecucion','Rango','Duracion de la prueba [s]','Dispositivo','Tipo de solucion','Tipo de centro digital','Estado de la prueba','Identificador de la prueba','Centro poblado','Dane institucion educativa','Tipo','Estado']]
    dataframe_1 = dataframe_1[(dataframe_1['Estado'] == 'EN OPERACIÓN')]
    dataframe_1 = dataframe_1[~(dataframe_1['Tipo'] == 'forzada')]
    dataframe_1['Identificador de la prueba'] = dataframe_1['Identificador de la prueba'].astype(str)
    dataframe_1 = dataframe_1[(dataframe_1['Rango'] >= 6) & (dataframe_1['Rango'] <= 12)]
    #dataframe_1 = dataframe_1[~(dataframe_1['Grupo'] == 'GRUPO 3A')]

    '''
    #filtro por matricula
    matricula1 = dataframe_1[dataframe_1['Matricula'].astype(str) == "Matricula <= 50"]
    matricula2 = dataframe_1[dataframe_1['Matricula'].astype(str) == "51 >= Matricula <= 150"]
    matricula3 = dataframe_1[dataframe_1['Matricula'].astype(str) == "151 >= Matricula <= 400"]
    matricula4 = dataframe_1[
        dataframe_1['Matricula'].astype(str).str.contains("Matricula > 400", na=False) |
        dataframe_1['Matricula'].astype(str).str.contains("Mejorada", na=False)
    ]
    '''

    #filtro por perfil
    matricula1 = dataframe_1[dataframe_1['Perfil'].astype(str) == "PERFIL 1"]
    matricula2 = dataframe_1[dataframe_1['Perfil'].astype(str) == "PERFIL 2"]
    matricula3 = dataframe_1[dataframe_1['Perfil'].astype(str) == "PERFIL 3+"]
    matricula4 = dataframe_1[dataframe_1['Perfil'].astype(str) == "PERFIL 4+"]

    print(dataframe.columns)

    #matricula1

    matricula1_rango6 = matricula1[(matricula1['Rango'] == 6)]

    #subida 0.05
    num_filas_6 = (len(matricula1_rango6))
    matricula1_posicion_prueba_p5_subida_6 = math.ceil(num_filas_6*0.05)

    matricula1_prueba_p5_subida_6 = 0 # Initialize
    matricula1_resultado_p5_subida_6 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_6 = math.ceil(num_filas_6*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_6 = 0 # Initialize
    matricula1_resultado_p95_subida_6 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_6 = math.ceil(num_filas_6*0.05)
    matricula1_prueba_p5_bajada_6 = 0 # Initialize
    matricula1_resultado_p5_bajada_6 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_6 = math.ceil(num_filas_6*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_6 = 0 # Initialize
    matricula1_resultado_p95_bajada_6 = 'NO APLICA' # Initialize


    if num_filas_6 > 0:
        matricula1_rango6=matricula1_rango6.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_6 = matricula1_rango6.iloc[matricula1_posicion_prueba_p5_subida_6-1]
        matricula1_prueba_p5_subida_6 = fila_subida_p5_6['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_6 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_6='CUMPLE'
        else:
           matricula1_resultado_p5_subida_6='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_6 = matricula1_rango6.iloc[matricula1_posicion_prueba_p95_subida_6-1]
        matricula1_prueba_p95_subida_6 = fila_subida_p95_6['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_6 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_6='CUMPLE'
        else:
           matricula1_resultado_p95_subida_6='NO CUMPLE'

        #bajada 0.05
        matricula1_rango6=matricula1_rango6.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_6 = matricula1_rango6.iloc[matricula1_posicion_prueba_p5_bajada_6-1]
        matricula1_prueba_p5_bajada_6 = fila_bajada_p5_6['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_6 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_6='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_6='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_6 = matricula1_rango6.iloc[matricula1_posicion_prueba_p95_bajada_6-1]
        matricula1_prueba_p95_bajada_6 = fila_bajada_p95_6['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_6 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_6='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_6='NO CUMPLE'


    matricula1_rango7 = matricula1[(matricula1['Rango'] == 7)]


    #subida 0.05
    num_filas_7 = (len(matricula1_rango7))
    matricula1_posicion_prueba_p5_subida_7 = math.ceil(num_filas_7*0.05)

    matricula1_prueba_p5_subida_7 = 0 # Initialize
    matricula1_resultado_p5_subida_7 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_7 = math.ceil(num_filas_7*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_7 = 0 # Initialize
    matricula1_resultado_p95_subida_7 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_7 = math.ceil(num_filas_7*0.05)
    matricula1_prueba_p5_bajada_7 = 0 # Initialize
    matricula1_resultado_p5_bajada_7 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_7 = math.ceil(num_filas_7*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_7 = 0 # Initialize
    matricula1_resultado_p95_bajada_7 = 'NO APLICA' # Initialize

    if num_filas_7 > 0:
        matricula1_rango7=matricula1_rango7.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_7 = matricula1_rango7.iloc[matricula1_posicion_prueba_p5_subida_7-1]
        matricula1_prueba_p5_subida_7 = fila_subida_p5_7['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_7 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_7='CUMPLE'
        else:
           matricula1_resultado_p5_subida_7='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_7 = matricula1_rango7.iloc[matricula1_posicion_prueba_p95_subida_7-1]
        matricula1_prueba_p95_subida_7 = fila_subida_p95_7['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_7 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_7='CUMPLE'
        else:
           matricula1_resultado_p95_subida_7='NO CUMPLE'

        #bajada 0.05
        matricula1_rango7=matricula1_rango7.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_7 = matricula1_rango7.iloc[matricula1_posicion_prueba_p5_bajada_7-1]
        matricula1_prueba_p5_bajada_7 = fila_bajada_p5_7['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_7 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_7='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_7='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_7 = matricula1_rango7.iloc[matricula1_posicion_prueba_p95_bajada_7-1]
        matricula1_prueba_p95_bajada_7 = fila_bajada_p95_7['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_7 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_7='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_7='NO CUMPLE'



    matricula1_rango8 = matricula1[(matricula1['Rango'] == 8)]


    #subida 0.05
    num_filas_8 = (len(matricula1_rango8))
    matricula1_posicion_prueba_p5_subida_8 = math.ceil(num_filas_8*0.05)

    matricula1_prueba_p5_subida_8 = 0 # Initialize
    matricula1_resultado_p5_subida_8 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_8 = math.ceil(num_filas_8*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_8 = 0 # Initialize
    matricula1_resultado_p95_subida_8 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_8 = math.ceil(num_filas_8*0.05)
    matricula1_prueba_p5_bajada_8 = 0 # Initialize
    matricula1_resultado_p5_bajada_8 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_8 = math.ceil(num_filas_8*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_8 = 0 # Initialize
    matricula1_resultado_p95_bajada_8 = 'NO APLICA' # Initialize

    if num_filas_8 > 0:
        matricula1_rango8=matricula1_rango8.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_8 = matricula1_rango8.iloc[matricula1_posicion_prueba_p5_subida_8-1]
        matricula1_prueba_p5_subida_8 = fila_subida_p5_8['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_8 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_8='CUMPLE'
        else:
           matricula1_resultado_p5_subida_8='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_8 = matricula1_rango8.iloc[matricula1_posicion_prueba_p95_subida_8-1]
        matricula1_prueba_p95_subida_8 = fila_subida_p95_8['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_8 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_8='CUMPLE'
        else:
           matricula1_resultado_p95_subida_8='NO CUMPLE'

        #bajada 0.05
        matricula1_rango8=matricula1_rango8.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_8 = matricula1_rango8.iloc[matricula1_posicion_prueba_p5_bajada_8-1]
        matricula1_prueba_p5_bajada_8 = fila_bajada_p5_8['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_8 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_8='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_8='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_8 = matricula1_rango8.iloc[matricula1_posicion_prueba_p95_bajada_8-1]
        matricula1_prueba_p95_bajada_8 = fila_bajada_p95_8['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_8 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_8='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_8='NO CUMPLE'



    matricula1_rango9 = matricula1[(matricula1['Rango'] == 9)]


    #subida 0.05
    num_filas_9 = (len(matricula1_rango9))
    matricula1_posicion_prueba_p5_subida_9 = math.ceil(num_filas_9*0.05)

    matricula1_prueba_p5_subida_9 = 0 # Initialize
    matricula1_resultado_p5_subida_9 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_9 = math.ceil(num_filas_9*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_9 = 0 # Initialize
    matricula1_resultado_p95_subida_9 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_9 = math.ceil(num_filas_9*0.05)
    matricula1_prueba_p5_bajada_9 = 0 # Initialize
    matricula1_resultado_p5_bajada_9 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_9 = math.ceil(num_filas_9*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_9 = 0 # Initialize
    matricula1_resultado_p95_bajada_9 = 'NO APLICA' # Initialize

    if num_filas_9 > 0:
        matricula1_rango9=matricula1_rango9.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_9 = matricula1_rango9.iloc[matricula1_posicion_prueba_p5_subida_9-1]
        matricula1_prueba_p5_subida_9 = fila_subida_p5_9['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_9 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_9='CUMPLE'
        else:
           matricula1_resultado_p5_subida_9='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_9 = matricula1_rango9.iloc[matricula1_posicion_prueba_p95_subida_9-1]
        matricula1_prueba_p95_subida_9 = fila_subida_p95_9['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_9 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_9='CUMPLE'
        else:
           matricula1_resultado_p95_subida_9='NO CUMPLE'

        #bajada 0.05
        matricula1_rango9=matricula1_rango9.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_9 = matricula1_rango9.iloc[matricula1_posicion_prueba_p5_bajada_9-1]
        matricula1_prueba_p5_bajada_9 = fila_bajada_p5_9['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_9 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_9='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_9='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_9 = matricula1_rango9.iloc[matricula1_posicion_prueba_p95_bajada_9-1]
        matricula1_prueba_p95_bajada_9 = fila_bajada_p95_9['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_9 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_9='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_9='NO CUMPLE'



    matricula1_rango10 = matricula1[(matricula1['Rango'] == 10)]


    #subida 0.05
    num_filas_10 = (len(matricula1_rango10))
    matricula1_posicion_prueba_p5_subida_10 = math.ceil(num_filas_10*0.05)

    matricula1_prueba_p5_subida_10 = 0 # Initialize
    matricula1_resultado_p5_subida_10 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_10 = math.ceil(num_filas_10*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_10 = 0 # Initialize
    matricula1_resultado_p95_subida_10 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_10 = math.ceil(num_filas_10*0.05)
    matricula1_prueba_p5_bajada_10 = 0 # Initialize
    matricula1_resultado_p5_bajada_10 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_10 = math.ceil(num_filas_10*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_10 = 0 # Initialize
    matricula1_resultado_p95_bajada_10 = 'NO APLICA' # Initialize

    if num_filas_10 > 0:
        matricula1_rango10=matricula1_rango10.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_10 = matricula1_rango10.iloc[matricula1_posicion_prueba_p5_subida_10-1]
        matricula1_prueba_p5_subida_10 = fila_subida_p5_10['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_10 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_10='CUMPLE'
        else:
           matricula1_resultado_p5_subida_10='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_10 = matricula1_rango10.iloc[matricula1_posicion_prueba_p95_subida_10-1]
        matricula1_prueba_p95_subida_10 = fila_subida_p95_10['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_10 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_10='CUMPLE'
        else:
           matricula1_resultado_p95_subida_10='NO CUMPLE'

        #bajada 0.05
        matricula1_rango10=matricula1_rango10.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_10 = matricula1_rango10.iloc[matricula1_posicion_prueba_p5_bajada_10-1]
        matricula1_prueba_p5_bajada_10 = fila_bajada_p5_10['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_10 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_10='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_10='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_10 = matricula1_rango10.iloc[matricula1_posicion_prueba_p95_bajada_10-1]
        matricula1_prueba_p95_bajada_10 = fila_bajada_p95_10['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_10 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_10='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_10='NO CUMPLE'



    matricula1_rango11 = matricula1[(matricula1['Rango'] == 11)]


    #subida 0.05
    num_filas_11 = (len(matricula1_rango11))
    matricula1_posicion_prueba_p5_subida_11 = math.ceil(num_filas_11*0.05)

    matricula1_prueba_p5_subida_11 = 0 # Initialize
    matricula1_resultado_p5_subida_11 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_11 = math.ceil(num_filas_11*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_11 = 0 # Initialize
    matricula1_resultado_p95_subida_11 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_11 = math.ceil(num_filas_11*0.05)
    matricula1_prueba_p5_bajada_11 = 0 # Initialize
    matricula1_resultado_p5_bajada_11 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_11 = math.ceil(num_filas_11*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_11 = 0 # Initialize
    matricula1_resultado_p95_bajada_11 = 'NO APLICA' # Initialize

    if num_filas_11 > 0:
        matricula1_rango11=matricula1_rango11.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_11 = matricula1_rango11.iloc[matricula1_posicion_prueba_p5_subida_11-1]
        matricula1_prueba_p5_subida_11 = fila_subida_p5_11['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_11 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_11='CUMPLE'
        else:
           matricula1_resultado_p5_subida_11='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_11 = matricula1_rango11.iloc[matricula1_posicion_prueba_p95_subida_11-1]
        matricula1_prueba_p95_subida_11 = fila_subida_p95_11['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_11 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_11='CUMPLE'
        else:
           matricula1_resultado_p95_subida_11='NO CUMPLE'

        #bajada 0.05
        matricula1_rango11=matricula1_rango11.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_11 = matricula1_rango11.iloc[matricula1_posicion_prueba_p5_bajada_11-1]
        matricula1_prueba_p5_bajada_11 = fila_bajada_p5_11['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_11 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_11='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_11='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_11 = matricula1_rango11.iloc[matricula1_posicion_prueba_p95_bajada_11-1]
        matricula1_prueba_p95_bajada_11 = fila_bajada_p95_11['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_11 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_11='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_11='NO CUMPLE'



    matricula1_rango12 = matricula1[(matricula1['Rango'] == 12)]

    #subida 0.05
    num_filas_12 = (len(matricula1_rango12))
    matricula1_posicion_prueba_p5_subida_12 = math.ceil(num_filas_12*0.05)

    matricula1_prueba_p5_subida_12 = 0 # Initialize
    matricula1_resultado_p5_subida_12 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_12 = math.ceil(num_filas_12*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_12 = 0 # Initialize
    matricula1_resultado_p95_subida_12 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_12 = math.ceil(num_filas_12*0.05)
    matricula1_prueba_p5_bajada_12 = 0 # Initialize
    matricula1_resultado_p5_bajada_12 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_12 = math.ceil(num_filas_12*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_12 = 0 # Initialize
    matricula1_resultado_p95_bajada_12 = 'NO APLICA' # Initialize

    if num_filas_12 > 0:
        matricula1_rango12=matricula1_rango12.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_12 = matricula1_rango12.iloc[matricula1_posicion_prueba_p5_subida_12-1]
        matricula1_prueba_p5_subida_12 = fila_subida_p5_12['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_12 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_12='CUMPLE'
        else:
           matricula1_resultado_p5_subida_12='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_12 = matricula1_rango12.iloc[matricula1_posicion_prueba_p95_subida_12-1]
        matricula1_prueba_p95_subida_12 = fila_subida_p95_12['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_12 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_12='CUMPLE'
        else:
           matricula1_resultado_p95_subida_12='NO CUMPLE'

        #bajada 0.05
        matricula1_rango12=matricula1_rango12.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_12 = matricula1_rango12.iloc[matricula1_posicion_prueba_p5_bajada_12-1]
        matricula1_prueba_p5_bajada_12 = fila_bajada_p5_12['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_12 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_12='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_12='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_12 = matricula1_rango12.iloc[matricula1_posicion_prueba_p95_bajada_12-1]
        matricula1_prueba_p95_bajada_12 = fila_bajada_p95_12['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_12 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_12='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_12='NO CUMPLE'



    matricula1_rango13 = matricula1[(matricula1['Rango'] == 13)]

    #subida 0.05
    num_filas_13 = (len(matricula1_rango13))
    matricula1_posicion_prueba_p5_subida_13 = math.ceil(num_filas_13*0.05)

    matricula1_prueba_p5_subida_13 = 0 # Initialize
    matricula1_resultado_p5_subida_13 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_13 = math.ceil(num_filas_13*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_13 = 0 # Initialize
    matricula1_resultado_p95_subida_13 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_13 = math.ceil(num_filas_13*0.05)
    matricula1_prueba_p5_bajada_13 = 0 # Initialize
    matricula1_resultado_p5_bajada_13 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_13 = math.ceil(num_filas_13*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_13 = 0 # Initialize
    matricula1_resultado_p95_bajada_13 = 'NO APLICA' # Initialize

    if num_filas_13 > 0:
        matricula1_rango13=matricula1_rango13.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_13 = matricula1_rango13.iloc[matricula1_posicion_prueba_p5_subida_13-1]
        matricula1_prueba_p5_subida_13 = fila_subida_p5_13['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_13 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_13='CUMPLE'
        else:
           matricula1_resultado_p5_subida_13='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_13 = matricula1_rango13.iloc[matricula1_posicion_prueba_p95_subida_13-1]
        matricula1_prueba_p95_subida_13 = fila_subida_p95_13['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_13 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_13='CUMPLE'
        else:
           matricula1_resultado_p95_subida_13='NO CUMPLE'
    
        #bajada 0.05
        matricula1_rango13=matricula1_rango13.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_13 = matricula1_rango13.iloc[matricula1_posicion_prueba_p5_bajada_13-1]
        matricula1_prueba_p5_bajada_13 = fila_bajada_p5_13['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_13 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_13='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_13='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_13 = matricula1_rango13.iloc[matricula1_posicion_prueba_p95_bajada_13-1]
        matricula1_prueba_p95_bajada_13 = fila_bajada_p95_13['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_13 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_13='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_13='NO CUMPLE'



    matricula1_rango14 = matricula1[(matricula1['Rango'] == 14)]

    #subida 0.05
    num_filas_14 = (len(matricula1_rango14))
    matricula1_posicion_prueba_p5_subida_14 = math.ceil(num_filas_14*0.05)

    matricula1_prueba_p5_subida_14 = 0 # Initialize
    matricula1_resultado_p5_subida_14 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_14 = math.ceil(num_filas_14*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_14 = 0 # Initialize
    matricula1_resultado_p95_subida_14 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_14 = math.ceil(num_filas_14*0.05)
    matricula1_prueba_p5_bajada_14 = 0 # Initialize
    matricula1_resultado_p5_bajada_14 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_14 = math.ceil(num_filas_14*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_14 = 0 # Initialize
    matricula1_resultado_p95_bajada_14 = 'NO APLICA' # Initialize

    if num_filas_14 > 0:
        matricula1_rango14=matricula1_rango14.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_14 = matricula1_rango14.iloc[matricula1_posicion_prueba_p5_subida_14-1]
        matricula1_prueba_p5_subida_14 = fila_subida_p5_14['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_14 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_14='CUMPLE'
        else:
           matricula1_resultado_p5_subida_14='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_14 = matricula1_rango14.iloc[matricula1_posicion_prueba_p95_subida_14-1]
        matricula1_prueba_p95_subida_14 = fila_subida_p95_14['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_14 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_14='CUMPLE'
        else:
           matricula1_resultado_p95_subida_14='NO CUMPLE'
    
        #bajada 0.05
        matricula1_rango14=matricula1_rango14.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_14 = matricula1_rango14.iloc[matricula1_posicion_prueba_p5_bajada_14-1]
        matricula1_prueba_p5_bajada_14 = fila_bajada_p5_14['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_14 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_14='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_14='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_14 = matricula1_rango14.iloc[matricula1_posicion_prueba_p95_bajada_14-1]
        matricula1_prueba_p95_bajada_14 = fila_bajada_p95_14['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_14 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_14='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_14='NO CUMPLE'



    matricula1_rango15 = matricula1[(matricula1['Rango'] == 15)]

    #subida 0.05
    num_filas_15 = (len(matricula1_rango15))
    matricula1_posicion_prueba_p5_subida_15 = math.ceil(num_filas_15*0.05)

    matricula1_prueba_p5_subida_15 = 0 # Initialize
    matricula1_resultado_p5_subida_15 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_15 = math.ceil(num_filas_15*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_15 = 0 # Initialize
    matricula1_resultado_p95_subida_15 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_15 = math.ceil(num_filas_15*0.05)
    matricula1_prueba_p5_bajada_15 = 0 # Initialize
    matricula1_resultado_p5_bajada_15 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_15 = math.ceil(num_filas_15*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_15 = 0 # Initialize
    matricula1_resultado_p95_bajada_15 = 'NO APLICA' # Initialize

    if num_filas_15 > 0:
        matricula1_rango15=matricula1_rango15.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_15 = matricula1_rango15.iloc[matricula1_posicion_prueba_p5_subida_15-1]
        matricula1_prueba_p5_subida_15 = fila_subida_p5_15['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_15 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_15='CUMPLE'
        else:
           matricula1_resultado_p5_subida_15='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_15 = matricula1_rango15.iloc[matricula1_posicion_prueba_p95_subida_15-1]
        matricula1_prueba_p95_subida_15 = fila_subida_p95_15['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_15 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_15='CUMPLE'
        else:
           matricula1_resultado_p95_subida_15='NO CUMPLE'
    
        #bajada 0.05
        matricula1_rango15=matricula1_rango15.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_15 = matricula1_rango15.iloc[matricula1_posicion_prueba_p5_bajada_15-1]
        matricula1_prueba_p5_bajada_15 = fila_bajada_p5_15['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_15 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_15='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_15='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_15 = matricula1_rango15.iloc[matricula1_posicion_prueba_p95_bajada_15-1]
        matricula1_prueba_p95_bajada_15 = fila_bajada_p95_15['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_15 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_15='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_15='NO CUMPLE'



    matricula1_rango16 = matricula1[(matricula1['Rango'] == 16)]


    #subida 0.05
    num_filas_16 = (len(matricula1_rango16))
    matricula1_posicion_prueba_p5_subida_16 = math.ceil(num_filas_16*0.05)

    matricula1_prueba_p5_subida_16 = 0 # Initialize
    matricula1_resultado_p5_subida_16 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_16 = math.ceil(num_filas_16*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_16 = 0 # Initialize
    matricula1_resultado_p95_subida_16 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_16 = math.ceil(num_filas_16*0.05)
    matricula1_prueba_p5_bajada_16 = 0 # Initialize
    matricula1_resultado_p5_bajada_16 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_16 = math.ceil(num_filas_16*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_16 = 0 # Initialize
    matricula1_resultado_p95_bajada_16 = 'NO APLICA' # Initialize

    if num_filas_16 > 0:
        matricula1_rango16=matricula1_rango16.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_16 = matricula1_rango16.iloc[matricula1_posicion_prueba_p5_subida_16-1]
        matricula1_prueba_p5_subida_16 = fila_subida_p5_16['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_16 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_16='CUMPLE'
        else:
           matricula1_resultado_p5_subida_16='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_16 = matricula1_rango16.iloc[matricula1_posicion_prueba_p95_subida_16-1]
        matricula1_prueba_p95_subida_16 = fila_subida_p95_16['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_16 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_16='CUMPLE'
        else:
           matricula1_resultado_p95_subida_16='NO CUMPLE'
    
        #bajada 0.05
        matricula1_rango16=matricula1_rango16.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_16 = matricula1_rango16.iloc[matricula1_posicion_prueba_p5_bajada_16-1]
        matricula1_prueba_p5_bajada_16 = fila_bajada_p5_16['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_16 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_16='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_16='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_16 = matricula1_rango16.iloc[matricula1_posicion_prueba_p95_bajada_16-1]
        matricula1_prueba_p95_bajada_16 = fila_bajada_p95_16['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_16 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_16='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_16='NO CUMPLE'



    matricula1_rango17 = matricula1[(matricula1['Rango'] == 17)]


    #subida 0.05
    num_filas_17 = (len(matricula1_rango17))
    matricula1_posicion_prueba_p5_subida_17 = math.ceil(num_filas_17*0.05)

    matricula1_prueba_p5_subida_17 = 0 # Initialize
    matricula1_resultado_p5_subida_17 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_17 = math.ceil(num_filas_17*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_17 = 0 # Initialize
    matricula1_resultado_p95_subida_17 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_17 = math.ceil(num_filas_17*0.05)
    matricula1_prueba_p5_bajada_17 = 0 # Initialize
    matricula1_resultado_p5_bajada_17 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_17 = math.ceil(num_filas_17*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_17 = 0 # Initialize
    matricula1_resultado_p95_bajada_17 = 'NO APLICA' # Initialize

    if num_filas_17 > 0:
        matricula1_rango17=matricula1_rango17.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_17 = matricula1_rango17.iloc[matricula1_posicion_prueba_p5_subida_17-1]
        matricula1_prueba_p5_subida_17 = fila_subida_p5_17['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_17 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_17='CUMPLE'
        else:
           matricula1_resultado_p5_subida_17='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_17 = matricula1_rango17.iloc[matricula1_posicion_prueba_p95_subida_17-1]
        matricula1_prueba_p95_subida_17 = fila_subida_p95_17['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_17 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_17='CUMPLE'
        else:
           matricula1_resultado_p95_subida_17='NO CUMPLE'

        #bajada 0.05
        matricula1_rango17=matricula1_rango17.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_17 = matricula1_rango17.iloc[matricula1_posicion_prueba_p5_bajada_17-1]
        matricula1_prueba_p5_bajada_17 = fila_bajada_p5_17['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_17 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_17='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_17='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_17 = matricula1_rango17.iloc[matricula1_posicion_prueba_p95_bajada_17-1]
        matricula1_prueba_p95_bajada_17 = fila_bajada_p95_17['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_17 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_17='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_17='NO CUMPLE'



    matricula1_rango18 = matricula1[(matricula1['Rango'] == 18)]

    #subida 0.05
    num_filas_18 = (len(matricula1_rango18))
    matricula1_posicion_prueba_p5_subida_18 = math.ceil(num_filas_18*0.05)

    matricula1_prueba_p5_subida_18 = 0 # Initialize
    matricula1_resultado_p5_subida_18 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_18 = math.ceil(num_filas_18*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_18 = 0 # Initialize
    matricula1_resultado_p95_subida_18 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_18 = math.ceil(num_filas_18*0.05)
    matricula1_prueba_p5_bajada_18 = 0 # Initialize
    matricula1_resultado_p5_bajada_18 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_18 = math.ceil(num_filas_18*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_18 = 0 # Initialize
    matricula1_resultado_p95_bajada_18 = 'NO APLICA' # Initialize

    if num_filas_18 > 0:
        matricula1_rango18=matricula1_rango18.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_18 = matricula1_rango18.iloc[matricula1_posicion_prueba_p5_subida_18-1]
        matricula1_prueba_p5_subida_18 = fila_subida_p5_18['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_18 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_18='CUMPLE'
        else:
           matricula1_resultado_p5_subida_18='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_18 = matricula1_rango18.iloc[matricula1_posicion_prueba_p95_subida_18-1]
        matricula1_prueba_p95_subida_18 = fila_subida_p95_18['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_18 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_18='CUMPLE'
        else:
           matricula1_resultado_p95_subida_18='NO CUMPLE'

        #bajada 0.05
        matricula1_rango18=matricula1_rango18.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_18 = matricula1_rango18.iloc[matricula1_posicion_prueba_p5_bajada_18-1]
        matricula1_prueba_p5_bajada_18 = fila_bajada_p5_18['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_18 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_18='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_18='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_18 = matricula1_rango18.iloc[matricula1_posicion_prueba_p95_bajada_18-1]
        matricula1_prueba_p95_bajada_18 = fila_bajada_p95_18['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_18 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_18='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_18='NO CUMPLE'



    matricula1_rango19 = matricula1[(matricula1['Rango'] == 19)]


    #subida 0.05
    num_filas_19 = (len(matricula1_rango19))
    matricula1_posicion_prueba_p5_subida_19 = math.ceil(num_filas_19*0.05)

    matricula1_prueba_p5_subida_19 = 0 # Initialize
    matricula1_resultado_p5_subida_19 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_19 = math.ceil(num_filas_19*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_19 = 0 # Initialize
    matricula1_resultado_p95_subida_19 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_19 = math.ceil(num_filas_19*0.05)
    matricula1_prueba_p5_bajada_19 = 0 # Initialize
    matricula1_resultado_p5_bajada_19 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_19 = math.ceil(num_filas_19*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_19 = 0 # Initialize
    matricula1_resultado_p95_bajada_19 = 'NO APLICA' # Initialize

    if num_filas_19 > 0:
        matricula1_rango19=matricula1_rango19.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_19 = matricula1_rango19.iloc[matricula1_posicion_prueba_p5_subida_19-1]
        matricula1_prueba_p5_subida_19 = fila_subida_p5_19['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_19 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_19='CUMPLE'
        else:
           matricula1_resultado_p5_subida_19='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_19 = matricula1_rango19.iloc[matricula1_posicion_prueba_p95_subida_19-1]
        matricula1_prueba_p95_subida_19 = fila_subida_p95_19['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_19 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_19='CUMPLE'
        else:
           matricula1_resultado_p95_subida_19='NO CUMPLE'

        #bajada 0.05
        matricula1_rango19=matricula1_rango19.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_19 = matricula1_rango19.iloc[matricula1_posicion_prueba_p5_bajada_19-1]
        matricula1_prueba_p5_bajada_19 = fila_bajada_p5_19['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_19 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_19='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_19='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_19 = matricula1_rango19.iloc[matricula1_posicion_prueba_p95_bajada_19-1]
        matricula1_prueba_p95_bajada_19 = fila_bajada_p95_19['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_19 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_19='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_19='NO CUMPLE'



    matricula1_rango20 = matricula1[(matricula1['Rango'] == 20)]

    #subida 0.05
    num_filas_20 = (len(matricula1_rango20))
    matricula1_posicion_prueba_p5_subida_20 = math.ceil(num_filas_20*0.05)
    
    matricula1_prueba_p5_subida_20 = 0 # Initialize
    matricula1_resultado_p5_subida_20 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_subida_20 = math.ceil(num_filas_20*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_subida_20 = 0 # Initialize
    matricula1_resultado_p95_subida_20 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula1_posicion_prueba_p5_bajada_20 = math.ceil(num_filas_20*0.05)
    matricula1_prueba_p5_bajada_20 = 0 # Initialize
    matricula1_resultado_p5_bajada_20 = 'NO APLICA' # Initialize
    matricula1_posicion_prueba_p95_bajada_20 = math.ceil(num_filas_20*0.95) # Keep outside if as it's just a calculation
    matricula1_prueba_p95_bajada_20 = 0 # Initialize
    matricula1_resultado_p95_bajada_20 = 'NO APLICA' # Initialize

    if num_filas_20 > 0:
        matricula1_rango20=matricula1_rango20.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_20 = matricula1_rango20.iloc[matricula1_posicion_prueba_p5_subida_20-1]
        matricula1_prueba_p5_subida_20 = fila_subida_p5_20['Velocidad de subida [Kbps]']
        if matricula1_prueba_p5_subida_20 >= matricula1_velocidad_subida:
           matricula1_resultado_p5_subida_20='CUMPLE'
        else:
           matricula1_resultado_p5_subida_20='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_20 = matricula1_rango20.iloc[matricula1_posicion_prueba_p95_subida_20-1]
        matricula1_prueba_p95_subida_20 = fila_subida_p95_20['Velocidad de subida [Kbps]']
        if matricula1_prueba_p95_subida_20 >= matricula1_velocidad_subida:
           matricula1_resultado_p95_subida_20='CUMPLE'
        else:
           matricula1_resultado_p95_subida_20='NO CUMPLE'

        #bajada 0.05
        matricula1_rango20=matricula1_rango20.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_20 = matricula1_rango20.iloc[matricula1_posicion_prueba_p5_bajada_20-1]
        matricula1_prueba_p5_bajada_20 = fila_bajada_p5_20['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p5_bajada_20 >= matricula1_velocidad_bajada:
           matricula1_resultado_p5_bajada_20='CUMPLE'
        else:
           matricula1_resultado_p5_bajada_20='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_20 = matricula1_rango20.iloc[matricula1_posicion_prueba_p95_bajada_20-1]
        matricula1_prueba_p95_bajada_20 = fila_bajada_p95_20['Velocidad de bajada [Kbps]']
        if matricula1_prueba_p95_bajada_20 >= matricula1_velocidad_bajada:
           matricula1_resultado_p95_bajada_20='CUMPLE'
        else:
           matricula1_resultado_p95_bajada_20='NO CUMPLE'
    
    #matricula2

    matricula2_rango6 = matricula2[(matricula2['Rango'] == 6)]

    #subida 0.05
    num_filas_6 = (len(matricula2_rango6))
    matricula2_posicion_prueba_p5_subida_6 = math.ceil(num_filas_6*0.05)

    matricula2_prueba_p5_subida_6 = 0 # Initialize
    matricula2_resultado_p5_subida_6 = 'NO APLICA' # Initialize
    matricula2_posicion_prueba_p95_subida_6 = math.ceil(num_filas_6*0.95) # Keep outside if as it's just a calculation
    matricula2_prueba_p95_subida_6 = 0 # Initialize
    matricula2_resultado_p95_subida_6 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula2_posicion_prueba_p5_bajada_6 = math.ceil(num_filas_6*0.05)
    matricula2_prueba_p5_bajada_6 = 0 # Initialize
    matricula2_resultado_p5_bajada_6 = 'NO APLICA' # Initialize
    matricula2_posicion_prueba_p95_bajada_6 = math.ceil(num_filas_6*0.95) # Keep outside if as it's just a calculation
    matricula2_prueba_p95_bajada_6 = 0 # Initialize
    matricula2_resultado_p95_bajada_6 = 'NO APLICA' # Initialize

    if num_filas_6 > 0:
        matricula2_rango6=matricula2_rango6.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_6 = matricula2_rango6.iloc[matricula2_posicion_prueba_p5_subida_6-1]
        matricula2_prueba_p5_subida_6 = fila_subida_p5_6['Velocidad de subida [Kbps]']
        if matricula2_prueba_p5_subida_6 >= matricula2_velocidad_subida:
           matricula2_resultado_p5_subida_6='CUMPLE'
        else:
           matricula2_resultado_p5_subida_6='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_6 = matricula2_rango6.iloc[matricula2_posicion_prueba_p95_subida_6-1]
        matricula2_prueba_p95_subida_6 = fila_subida_p95_6['Velocidad de subida [Kbps]']
        if matricula2_prueba_p95_subida_6 >= matricula2_velocidad_subida:
           matricula2_resultado_p95_subida_6='CUMPLE'
        else:
           matricula2_resultado_p95_subida_6='NO CUMPLE'

        #bajada 0.05
        matricula2_rango6=matricula2_rango6.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_6 = matricula2_rango6.iloc[matricula2_posicion_prueba_p5_bajada_6-1]
        matricula2_prueba_p5_bajada_6 = fila_bajada_p5_6['Velocidad de bajada [Kbps]']
        if matricula2_prueba_p5_bajada_6 >= matricula2_velocidad_bajada:
           matricula2_resultado_p5_bajada_6='CUMPLE'
        else:
           matricula2_resultado_p5_bajada_6='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_6 = matricula2_rango6.iloc[matricula2_posicion_prueba_p95_bajada_6-1]
        matricula2_prueba_p95_bajada_6 = fila_bajada_p95_6['Velocidad de bajada [Kbps]']
        if matricula2_prueba_p95_bajada_6 >= matricula2_velocidad_bajada:
           matricula2_resultado_p95_bajada_6='CUMPLE'
        else:
           matricula2_resultado_p95_bajada_6='NO CUMPLE'


    matricula2_rango7 = matricula2[(matricula2['Rango'] == 7)]


    #subida 0.05
    num_filas_7 = (len(matricula2_rango7))
    matricula2_posicion_prueba_p5_subida_7 = math.ceil(num_filas_7*0.05)

    matricula2_prueba_p5_subida_7 = 0 # Initialize
    matricula2_resultado_p5_subida_7 = 'NO APLICA' # Initialize
    matricula2_posicion_prueba_p95_subida_7 = math.ceil(num_filas_7*0.95) # Keep outside if as it's just a calculation
    matricula2_prueba_p95_subida_7 = 0 # Initialize
    matricula2_resultado_p95_subida_7 = 'NO APLICA' # Initialize

    #bajada 0.05
    matricula2_posicion_prueba_p5_bajada_7 = math.ceil(num_filas_7*0.05)
    matricula2_prueba_p5_bajada_7 = 0 # Initialize
    matricula2_resultado_p5_bajada_7 = 'NO APLICA' # Initialize
    matricula2_posicion_prueba_p95_bajada_7 = math.ceil(num_filas_7*0.95) # Keep outside if as it's just a calculation
    matricula2_prueba_p95_bajada_7 = 0 # Initialize
    matricula2_resultado_p95_bajada_7 = 'NO APLICA' # Initialize

    if num_filas_7 > 0:
        matricula2_rango7=matricula2_rango7.sort_values(by='Velocidad de subida [Kbps]')
        fila_subida_p5_7 = matricula2_rango7.iloc[matricula2_posicion_prueba_p5_subida_7-1]
        matricula2_prueba_p5_subida_7 = fila_subida_p5_7['Velocidad de subida [Kbps]']
        if matricula2_prueba_p5_subida_7 >= matricula2_velocidad_subida:
           matricula2_resultado_p5_subida_7='CUMPLE'
        else:
           matricula2_resultado_p5_subida_7='NO CUMPLE'
        #subida 0.95
        fila_subida_p95_7 = matricula2_rango7.iloc[matricula2_posicion_prueba_p95_subida_7-1]
        matricula2_prueba_p95_subida_7 = fila_subida_p95_7['Velocidad de subida [Kbps]']
        if matricula2_prueba_p95_subida_7 >= matricula2_velocidad_subida:
           matricula2_resultado_p95_subida_7='CUMPLE'
        else:
           matricula2_resultado_p95_subida_7='NO CUMPLE'

        #bajada 0.05
        matricula2_rango7=matricula2_rango7.sort_values(by='Velocidad de bajada [Kbps]')
        fila_bajada_p5_7 = matricula2_rango7.iloc[matricula2_posicion_prueba_p5_bajada_7-1]
        matricula2_prueba_p5_bajada_7 = fila_bajada_p5_7['Velocidad de bajada [Kbps]']
        if matricula2_prueba_p5_bajada_7 >= matricula2_velocidad_bajada:
           matricula2_resultado_p5_bajada_7='CUMPLE'
        else:
           matricula2_resultado_p5_bajada_7='NO CUMPLE'
        #bajada 0.95
        fila_bajada_p95_7 = matricula2_rango7.iloc[matricula2_posicion_prueba_p95_bajada_7-1]
        matricula2_prueba_p95_bajada_7 = fila_bajada_p95_7['Velocidad de bajada [Kbps]']
        if matricula2_prueba_p95_bajada_7 >= matricula2_velocidad_bajada:
           matricula2_resultado_p95_bajada_7='CUMPLE'
        else:
           matricula2_resultado_p95_bajada_7='NO CUMPLE'
    
######################voy aca

matricula2_rango8 = matricula2[(matricula2['Rango'] == 8)]


#subida 0.05
num_filas_8 = (len(matricula2_rango8))
matricula2_posicion_prueba_p5_subida_8 = math.ceil(num_filas_8*0.05)

matricula2_prueba_p5_subida_8 = 0 # Initialize
matricula2_resultado_p5_subida_8 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_8 = math.ceil(num_filas_8*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_8 = 0 # Initialize
matricula2_resultado_p95_subida_8 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_8 = math.ceil(num_filas_8*0.05)
matricula2_prueba_p5_bajada_8 = 0 # Initialize
matricula2_resultado_p5_bajada_8 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_8 = math.ceil(num_filas_8*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_8 = 0 # Initialize
matricula2_resultado_p95_bajada_8 = 'NO APLICA' # Initialize

if num_filas_8 > 0:
    matricula2_rango8=matricula2_rango8.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_8 = matricula2_rango8.iloc[matricula2_posicion_prueba_p5_subida_8-1]
    matricula2_prueba_p5_subida_8 = fila_subida_p5_8['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_8 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_8='CUMPLE'
    else:
       matricula2_resultado_p5_subida_8='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_8 = matricula2_rango8.iloc[matricula2_posicion_prueba_p95_subida_8-1]
    matricula2_prueba_p95_subida_8 = fila_subida_p95_8['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_8 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_8='CUMPLE'
    else:
       matricula2_resultado_p95_subida_8='NO CUMPLE'

    #bajada 0.05
    matricula2_rango8=matricula2_rango8.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_8 = matricula2_rango8.iloc[matricula2_posicion_prueba_p5_bajada_8-1]
    matricula2_prueba_p5_bajada_8 = fila_bajada_p5_8['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_8 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_8='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_8='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_8 = matricula2_rango8.iloc[matricula2_posicion_prueba_p95_bajada_8-1]
    matricula2_prueba_p95_bajada_8 = fila_bajada_p95_8['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_8 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_8='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_8='NO CUMPLE'



matricula2_rango9 = matricula2[(matricula2['Rango'] == 9)]


#subida 0.05
num_filas_9 = (len(matricula2_rango9))
matricula2_posicion_prueba_p5_subida_9 = math.ceil(num_filas_9*0.05)

matricula2_prueba_p5_subida_9 = 0 # Initialize
matricula2_resultado_p5_subida_9 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_9 = math.ceil(num_filas_9*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_9 = 0 # Initialize
matricula2_resultado_p95_subida_9 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_9 = math.ceil(num_filas_9*0.05)
matricula2_prueba_p5_bajada_9 = 0 # Initialize
matricula2_resultado_p5_bajada_9 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_9 = math.ceil(num_filas_9*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_9 = 0 # Initialize
matricula2_resultado_p95_bajada_9 = 'NO APLICA' # Initialize

if num_filas_9 > 0:
    matricula2_rango9=matricula2_rango9.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_9 = matricula2_rango9.iloc[matricula2_posicion_prueba_p5_subida_9-1]
    matricula2_prueba_p5_subida_9 = fila_subida_p5_9['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_9 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_9='CUMPLE'
    else:
       matricula2_resultado_p5_subida_9='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_9 = matricula2_rango9.iloc[matricula2_posicion_prueba_p95_subida_9-1]
    matricula2_prueba_p95_subida_9 = fila_subida_p95_9['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_9 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_9='CUMPLE'
    else:
       matricula2_resultado_p95_subida_9='NO CUMPLE'

    #bajada 0.05
    matricula2_rango9=matricula2_rango9.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_9 = matricula2_rango9.iloc[matricula2_posicion_prueba_p5_bajada_9-1]
    matricula2_prueba_p5_bajada_9 = fila_bajada_p5_9['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_9 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_9='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_9='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_9 = matricula2_rango9.iloc[matricula2_posicion_prueba_p95_bajada_9-1]
    matricula2_prueba_p95_bajada_9 = fila_bajada_p95_9['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_9 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_9='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_9='NO CUMPLE'



matricula2_rango10 = matricula2[(matricula2['Rango'] == 10)]


#subida 0.05
num_filas_10 = (len(matricula2_rango10))
matricula2_posicion_prueba_p5_subida_10 = math.ceil(num_filas_10*0.05)

matricula2_prueba_p5_subida_10 = 0 # Initialize
matricula2_resultado_p5_subida_10 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_10 = math.ceil(num_filas_10*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_10 = 0 # Initialize
matricula2_resultado_p95_subida_10 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_10 = math.ceil(num_filas_10*0.05)
matricula2_prueba_p5_bajada_10 = 0 # Initialize
matricula2_resultado_p5_bajada_10 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_10 = math.ceil(num_filas_10*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_10 = 0 # Initialize
matricula2_resultado_p95_bajada_10 = 'NO APLICA' # Initialize

if num_filas_10 > 0:
    matricula2_rango10=matricula2_rango10.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_10 = matricula2_rango10.iloc[matricula2_posicion_prueba_p5_subida_10-1]
    matricula2_prueba_p5_subida_10 = fila_subida_p5_10['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_10 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_10='CUMPLE'
    else:
       matricula2_resultado_p5_subida_10='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_10 = matricula2_rango10.iloc[matricula2_posicion_prueba_p95_subida_10-1]
    matricula2_prueba_p95_subida_10 = fila_subida_p95_10['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_10 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_10='CUMPLE'
    else:
       matricula2_resultado_p105_subida_10='NO CUMPLE'

    #bajada 0.05
    matricula2_rango10=matricula2_rango10.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_10 = matricula2_rango10.iloc[matricula2_posicion_prueba_p5_bajada_10-1]
    matricula2_prueba_p5_bajada_10 = fila_bajada_p5_10['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_10 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_10='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_10='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_10 = matricula2_rango10.iloc[matricula2_posicion_prueba_p95_bajada_10-1]
    matricula2_prueba_p95_bajada_10 = fila_bajada_p95_10['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_10 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_10='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_10='NO CUMPLE'



matricula2_rango11 = matricula2[(matricula2['Rango'] == 11)]


#subida 0.05
num_filas_11 = (len(matricula2_rango11))
matricula2_posicion_prueba_p5_subida_11 = math.ceil(num_filas_11*0.05)

matricula2_prueba_p5_subida_11 = 0 # Initialize
matricula2_resultado_p5_subida_11 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_11 = math.ceil(num_filas_11*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_11 = 0 # Initialize
matricula2_resultado_p95_subida_11 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_11 = math.ceil(num_filas_11*0.05)
matricula2_prueba_p5_bajada_11 = 0 # Initialize
matricula2_resultado_p5_bajada_11 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_11 = math.ceil(num_filas_11*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_11 = 0 # Initialize
matricula2_resultado_p95_bajada_11 = 'NO APLICA' # Initialize

if num_filas_11 > 0:
    matricula2_rango11=matricula2_rango11.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_11 = matricula2_rango11.iloc[matricula2_posicion_prueba_p5_subida_11-1]
    matricula2_prueba_p5_subida_11 = fila_subida_p5_11['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_11 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_11='CUMPLE'
    else:
       matricula2_resultado_p5_subida_11='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_11 = matricula2_rango11.iloc[matricula2_posicion_prueba_p95_subida_11-1]
    matricula2_prueba_p95_subida_11 = fila_subida_p95_11['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_11 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_11='CUMPLE'
    else:
       matricula2_resultado_p95_subida_11='NO CUMPLE'

    #bajada 0.05
    matricula2_rango11=matricula2_rango11.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_11 = matricula2_rango11.iloc[matricula2_posicion_prueba_p5_bajada_11-1]
    matricula2_prueba_p5_bajada_11 = fila_bajada_p5_11['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_11 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_11='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_11='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_11 = matricula2_rango11.iloc[matricula2_posicion_prueba_p95_bajada_11-1]
    matricula2_prueba_p95_bajada_11 = fila_bajada_p95_11['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_11 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_11='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_11='NO CUMPLE'



matricula2_rango12 = matricula2[(matricula2['Rango'] == 12)]

#subida 0.05
num_filas_12 = (len(matricula2_rango12))
matricula2_posicion_prueba_p5_subida_12 = math.ceil(num_filas_12*0.05)

matricula2_prueba_p5_subida_12 = 0 # Initialize
matricula2_resultado_p5_subida_12 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_12 = math.ceil(num_filas_12*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_12 = 0 # Initialize
matricula2_resultado_p95_subida_12 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_12 = math.ceil(num_filas_12*0.05)
matricula2_prueba_p5_bajada_12 = 0 # Initialize
matricula2_resultado_p5_bajada_12 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_12 = math.ceil(num_filas_12*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_12 = 0 # Initialize
matricula2_resultado_p95_bajada_12 = 'NO APLICA' # Initialize

if num_filas_12 > 0:
    matricula2_rango12=matricula2_rango12.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_12 = matricula2_rango12.iloc[matricula2_posicion_prueba_p5_subida_12-1]
    matricula2_prueba_p5_subida_12 = fila_subida_p5_12['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_12 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_12='CUMPLE'
    else:
       matricula2_resultado_p5_subida_12='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_12 = matricula2_rango12.iloc[matricula2_posicion_prueba_p95_subida_12-1]
    matricula2_prueba_p95_subida_12 = fila_subida_p95_12['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_12 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_12='CUMPLE'
    else:
       matricula2_resultado_p95_subida_12='NO CUMPLE'

    #bajada 0.05
    matricula2_rango12=matricula2_rango12.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_12 = matricula2_rango12.iloc[matricula2_posicion_prueba_p5_bajada_12-1]
    matricula2_prueba_p5_bajada_12 = fila_bajada_p5_12['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_12 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_12='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_12='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_12 = matricula2_rango12.iloc[matricula2_posicion_prueba_p95_bajada_12-1]
    matricula2_prueba_p95_bajada_12 = fila_bajada_p95_12['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_12 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_12='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_12='NO CUMPLE'



matricula2_rango13 = matricula2[(matricula2['Rango'] == 13)]

#subida 0.05
num_filas_13 = (len(matricula2_rango13))
matricula2_posicion_prueba_p5_subida_13 = math.ceil(num_filas_13*0.05)

matricula2_prueba_p5_subida_13 = 0 # Initialize
matricula2_resultado_p5_subida_13 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_13 = math.ceil(num_filas_13*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_13 = 0 # Initialize
matricula2_resultado_p95_subida_13 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_13 = math.ceil(num_filas_13*0.05)
matricula2_prueba_p5_bajada_13 = 0 # Initialize
matricula2_resultado_p5_bajada_13 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_13 = math.ceil(num_filas_13*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_13 = 0 # Initialize
matricula2_resultado_p95_bajada_13 = 'NO APLICA' # Initialize

if num_filas_13 > 0:
    matricula2_rango13=matricula2_rango13.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_13 = matricula2_rango13.iloc[matricula2_posicion_prueba_p5_subida_13-1]
    matricula2_prueba_p5_subida_13 = fila_subida_p5_13['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_13 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_13='CUMPLE'
    else:
       matricula2_resultado_p5_subida_13='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_13 = matricula2_rango13.iloc[matricula2_posicion_prueba_p95_subida_13-1]
    matricula2_prueba_p95_subida_13 = fila_subida_p95_13['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_13 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_13='CUMPLE'
    else:
       matricula2_resultado_p95_subida_13='NO CUMPLE'

    #bajada 0.05
    matricula2_rango13=matricula2_rango13.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_13 = matricula2_rango13.iloc[matricula2_posicion_prueba_p5_bajada_13-1]
    matricula2_prueba_p5_bajada_13 = fila_bajada_p5_13['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_13 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_13='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_13='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_13 = matricula2_rango13.iloc[matricula2_posicion_prueba_p95_bajada_13-1]
    matricula2_prueba_p95_bajada_13 = fila_bajada_p95_13['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_13 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_13='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_13='NO CUMPLE'



matricula2_rango14 = matricula2[(matricula2['Rango'] == 14)]

#subida 0.05
num_filas_14 = (len(matricula2_rango14))
matricula2_posicion_prueba_p5_subida_14 = math.ceil(num_filas_14*0.05)

matricula2_prueba_p5_subida_14 = 0 # Initialize
matricula2_resultado_p5_subida_14 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_14 = math.ceil(num_filas_14*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_14 = 0 # Initialize
matricula2_resultado_p95_subida_14 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_14 = math.ceil(num_filas_14*0.05)
matricula2_prueba_p5_bajada_14 = 0 # Initialize
matricula2_resultado_p5_bajada_14 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_14 = math.ceil(num_filas_14*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_14 = 0 # Initialize
matricula2_resultado_p95_bajada_14 = 'NO APLICA' # Initialize

if num_filas_14 > 0:
    matricula2_rango14=matricula2_rango14.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_14 = matricula2_rango14.iloc[matricula2_posicion_prueba_p5_subida_14-1]
    matricula2_prueba_p5_subida_14 = fila_subida_p5_14['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_14 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_14='CUMPLE'
    else:
       matricula2_resultado_p5_subida_14='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_14 = matricula2_rango14.iloc[matricula2_posicion_prueba_p95_subida_14-1]
    matricula2_prueba_p95_subida_14 = fila_subida_p95_14['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_14 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_14='CUMPLE'
    else:
       matricula2_resultado_p95_subida_14='NO CUMPLE'

    #bajada 0.05
    matricula2_rango14=matricula2_rango14.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_14 = matricula2_rango14.iloc[matricula2_posicion_prueba_p5_bajada_14-1]
    matricula2_prueba_p5_bajada_14 = fila_bajada_p5_14['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_14 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_14='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_14='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_14 = matricula2_rango14.iloc[matricula2_posicion_prueba_p95_bajada_14-1]
    matricula2_prueba_p95_bajada_14 = fila_bajada_p95_14['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_14 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_14='CUMPLE'
    else:
       matricula2_resultado_p145_bajada_14='NO CUMPLE'




matricula2_rango15 = matricula2[(matricula2['Rango'] == 15)]

#subida 0.05
num_filas_15 = (len(matricula2_rango15))
matricula2_posicion_prueba_p5_subida_15 = math.ceil(num_filas_15*0.05)

matricula2_prueba_p5_subida_15 = 0 # Initialize
matricula2_resultado_p5_subida_15 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_15 = math.ceil(num_filas_15*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_15 = 0 # Initialize
matricula2_resultado_p95_subida_15 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_15 = math.ceil(num_filas_15*0.05)
matricula2_prueba_p5_bajada_15 = 0 # Initialize
matricula2_resultado_p5_bajada_15 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_15 = math.ceil(num_filas_15*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_15 = 0 # Initialize
matricula2_resultado_p95_bajada_15 = 'NO APLICA' # Initialize

if num_filas_15 > 0:
    matricula2_rango15=matricula2_rango15.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_15 = matricula2_rango15.iloc[matricula2_posicion_prueba_p5_subida_15-1]
    matricula2_prueba_p5_subida_15 = fila_subida_p5_15['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_15 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_15='CUMPLE'
    else:
       matricula2_resultado_p5_subida_15='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_15 = matricula2_rango15.iloc[matricula2_posicion_prueba_p95_subida_15-1]
    matricula2_prueba_p95_subida_15 = fila_subida_p95_15['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_15 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_15='CUMPLE'
    else:
       matricula2_resultado_p155_subida_15='NO CUMPLE'

    #bajada 0.05
    matricula2_rango15=matricula2_rango15.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_15 = matricula2_rango15.iloc[matricula2_posicion_prueba_p5_bajada_15-1]
    matricula2_prueba_p5_bajada_15 = fila_bajada_p5_15['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_15 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_15='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_15='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_15 = matricula2_rango15.iloc[matricula2_posicion_prueba_p95_bajada_15-1]
    matricula2_prueba_p95_bajada_15 = fila_bajada_p95_15['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_15 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_15='CUMPLE'
    else:
       matricula2_resultado_p155_bajada_15='NO CUMPLE'




matricula2_rango16 = matricula2[(matricula2['Rango'] == 16)]


#subida 0.05
num_filas_16 = (len(matricula2_rango16))
matricula2_posicion_prueba_p5_subida_16 = math.ceil(num_filas_16*0.05)

matricula2_prueba_p5_subida_16 = 0 # Initialize
matricula2_resultado_p5_subida_16 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_16 = math.ceil(num_filas_16*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_16 = 0 # Initialize
matricula2_resultado_p95_subida_16 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_16 = math.ceil(num_filas_16*0.05)
matricula2_prueba_p5_bajada_16 = 0 # Initialize
matricula2_resultado_p5_bajada_16 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_16 = math.ceil(num_filas_16*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_16 = 0 # Initialize
matricula2_resultado_p95_bajada_16 = 'NO APLICA' # Initialize

if num_filas_16 > 0:
    matricula2_rango16=matricula2_rango16.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_16 = matricula2_rango16.iloc[matricula2_posicion_prueba_p5_subida_16-1]
    matricula2_prueba_p5_subida_16 = fila_subida_p5_16['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_16 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_16='CUMPLE'
    else:
       matricula2_resultado_p5_subida_16='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_16 = matricula2_rango16.iloc[matricula2_posicion_prueba_p95_subida_16-1]
    matricula2_prueba_p95_subida_16 = fila_subida_p95_16['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_16 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_16='CUMPLE'
    else:
       matricula2_resultado_p95_subida_16='NO CUMPLE'

    #bajada 0.05
    matricula2_rango16=matricula2_rango16.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_16 = matricula2_rango16.iloc[matricula2_posicion_prueba_p5_bajada_16-1]
    matricula2_prueba_p5_bajada_16 = fila_bajada_p5_16['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_16 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_16='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_16='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_16 = matricula2_rango16.iloc[matricula2_posicion_prueba_p95_bajada_16-1]
    matricula2_prueba_p95_bajada_16 = fila_bajada_p95_16['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_16 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_16='CUMPLE'
    else:
       matricula2_resultado_p165_bajada_16='NO CUMPLE'




matricula2_rango17 = matricula2[(matricula2['Rango'] == 17)]


#subida 0.05
num_filas_17 = (len(matricula2_rango17))
matricula2_posicion_prueba_p5_subida_17 = math.ceil(num_filas_17*0.05)

matricula2_prueba_p5_subida_17 = 0 # Initialize
matricula2_resultado_p5_subida_17 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_17 = math.ceil(num_filas_17*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_17 = 0 # Initialize
matricula2_resultado_p95_subida_17 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_17 = math.ceil(num_filas_17*0.05)
matricula2_prueba_p5_bajada_17 = 0 # Initialize
matricula2_resultado_p5_bajada_17 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_17 = math.ceil(num_filas_17*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_17 = 0 # Initialize
matricula2_resultado_p95_bajada_17 = 'NO APLICA' # Initialize

if num_filas_17 > 0:
    matricula2_rango17=matricula2_rango17.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_17 = matricula2_rango17.iloc[matricula2_posicion_prueba_p5_subida_17-1]
    matricula2_prueba_p5_subida_17 = fila_subida_p5_17['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_17 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_17='CUMPLE'
    else:
       matricula2_resultado_p5_subida_17='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_17 = matricula2_rango17.iloc[matricula2_posicion_prueba_p95_subida_17-1]
    matricula2_prueba_p95_subida_17 = fila_subida_p95_17['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_17 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_17='CUMPLE'
    else:
       matricula2_resultado_p95_subida_17='NO CUMPLE'

    #bajada 0.05
    matricula2_rango17=matricula2_rango17.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_17 = matricula2_rango17.iloc[matricula2_posicion_prueba_p5_bajada_17-1]
    matricula2_prueba_p5_bajada_17 = fila_bajada_p5_17['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_17 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_17='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_17='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_17 = matricula2_rango17.iloc[matricula2_posicion_prueba_p95_bajada_17-1]
    matricula2_prueba_p95_bajada_17 = fila_bajada_p95_17['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_17 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_17='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_17='NO CUMPLE'



matricula2_rango18 = matricula2[(matricula2['Rango'] == 18)]

#subida 0.05
num_filas_18 = (len(matricula2_rango18))
matricula2_posicion_prueba_p5_subida_18 = math.ceil(num_filas_18*0.05)

matricula2_prueba_p5_subida_18 = 0 # Initialize
matricula2_resultado_p5_subida_18 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_18 = math.ceil(num_filas_18*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_18 = 0 # Initialize
matricula2_resultado_p95_subida_18 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_18 = math.ceil(num_filas_18*0.05)
matricula2_prueba_p5_bajada_18 = 0 # Initialize
matricula2_resultado_p5_bajada_18 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_18 = math.ceil(num_filas_18*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_18 = 0 # Initialize
matricula2_resultado_p95_bajada_18 = 'NO APLICA' # Initialize

if num_filas_18 > 0:
    matricula2_rango18=matricula2_rango18.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_18 = matricula2_rango18.iloc[matricula2_posicion_prueba_p5_subida_18-1]
    matricula2_prueba_p5_subida_18 = fila_subida_p5_18['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_18 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_18='CUMPLE'
    else:
       matricula2_resultado_p5_subida_18='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_18 = matricula2_rango18.iloc[matricula2_posicion_prueba_p95_subida_18-1]
    matricula2_prueba_p95_subida_18 = fila_subida_p95_18['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_18 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_18='CUMPLE'
    else:
       matricula2_resultado_p95_subida_18='NO CUMPLE'

    #bajada 0.05
    matricula2_rango18=matricula2_rango18.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_18 = matricula2_rango18.iloc[matricula2_posicion_prueba_p5_bajada_18-1]
    matricula2_prueba_p5_bajada_18 = fila_bajada_p5_18['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_18 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_18='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_18='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_18 = matricula2_rango18.iloc[matricula2_posicion_prueba_p95_bajada_18-1]
    matricula2_prueba_p95_bajada_18 = fila_bajada_p95_18['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_18 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_18='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_18='NO CUMPLE'



matricula2_rango19 = matricula2[(matricula2['Rango'] == 19)]


#subida 0.05
num_filas_19 = (len(matricula2_rango19))
matricula2_posicion_prueba_p5_subida_19 = math.ceil(num_filas_19*0.05)

matricula2_prueba_p5_subida_19 = 0 # Initialize
matricula2_resultado_p5_subida_19 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_19 = math.ceil(num_filas_19*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_19 = 0 # Initialize
matricula2_resultado_p95_subida_19 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_19 = math.ceil(num_filas_19*0.05)
matricula2_prueba_p5_bajada_19 = 0 # Initialize
matricula2_resultado_p5_bajada_19 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_19 = math.ceil(num_filas_19*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_19 = 0 # Initialize
matricula2_resultado_p95_bajada_19 = 'NO APLICA' # Initialize

if num_filas_19 > 0:
    matricula2_rango19=matricula2_rango19.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_19 = matricula2_rango19.iloc[matricula2_posicion_prueba_p5_subida_19-1]
    matricula2_prueba_p5_subida_19 = fila_subida_p5_19['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_19 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_19='CUMPLE'
    else:
       matricula2_resultado_p5_subida_19='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_19 = matricula2_rango19.iloc[matricula2_posicion_prueba_p95_subida_19-1]
    matricula2_prueba_p95_subida_19 = fila_subida_p95_19['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_19 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_19='CUMPLE'
    else:
       matricula2_resultado_p95_subida_19='NO CUMPLE'

    #bajada 0.05
    matricula2_rango19=matricula2_rango19.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_19 = matricula2_rango19.iloc[matricula2_posicion_prueba_p5_bajada_19-1]
    matricula2_prueba_p5_bajada_19 = fila_bajada_p5_19['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_19 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_19='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_19='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_19 = matricula2_rango19.iloc[matricula2_posicion_prueba_p95_bajada_19-1]
    matricula2_prueba_p95_bajada_19 = fila_bajada_p95_19['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_19 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_19='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_19='NO CUMPLE'



matricula2_rango20 = matricula2[(matricula2['Rango'] == 20)]

#subida 0.05
num_filas_20 = (len(matricula2_rango20))
matricula2_posicion_prueba_p5_subida_20 = math.ceil(num_filas_20*0.05)

matricula2_prueba_p5_subida_20 = 0 # Initialize
matricula2_resultado_p5_subida_20 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_subida_20 = math.ceil(num_filas_20*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_subida_20 = 0 # Initialize
matricula2_resultado_p95_subida_20 = 'NO APLICA' # Initialize

#bajada 0.05
matricula2_posicion_prueba_p5_bajada_20 = math.ceil(num_filas_20*0.05)
matricula2_prueba_p5_bajada_20 = 0 # Initialize
matricula2_resultado_p5_bajada_20 = 'NO APLICA' # Initialize
matricula2_posicion_prueba_p95_bajada_20 = math.ceil(num_filas_20*0.95) # Keep outside if as it's just a calculation
matricula2_prueba_p95_bajada_20 = 0 # Initialize
matricula2_resultado_p95_bajada_20 = 'NO APLICA' # Initialize

if num_filas_20 > 0:
    matricula2_rango20=matricula2_rango20.sort_values(by='Velocidad de subida [Kbps]')
    fila_subida_p5_20 = matricula2_rango20.iloc[matricula2_posicion_prueba_p5_subida_20-1]
    matricula2_prueba_p5_subida_20 = fila_subida_p5_20['Velocidad de subida [Kbps]']
    if matricula2_prueba_p5_subida_20 >= matricula2_velocidad_subida:
       matricula2_resultado_p5_subida_20='CUMPLE'
    else:
       matricula2_resultado_p5_subida_20='NO CUMPLE'
    #subida 0.95
    fila_subida_p95_20 = matricula2_rango20.iloc[matricula2_posicion_prueba_p95_subida_20-1]
    matricula2_prueba_p95_subida_20 = fila_subida_p95_20['Velocidad de subida [Kbps]']
    if matricula2_prueba_p95_subida_20 >= matricula2_velocidad_subida:
       matricula2_resultado_p95_subida_20='CUMPLE'
    else:
       matricula2_resultado_p95_subida_20='NO CUMPLE'

    #bajada 0.05
    matricula2_rango20=matricula2_rango20.sort_values(by='Velocidad de bajada [Kbps]')
    fila_bajada_p5_20 = matricula2_rango20.iloc[matricula2_posicion_prueba_p5_bajada_20-1]
    matricula2_prueba_p5_bajada_20 = fila_bajada_p5_20['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p5_bajada_20 >= matricula2_velocidad_bajada:
       matricula2_resultado_p5_bajada_20='CUMPLE'
    else:
       matricula2_resultado_p5_bajada_20='NO CUMPLE'
    #bajada 0.95
    fila_bajada_p95_20 = matricula2_rango20.iloc[matricula2_posicion_prueba_p95_bajada_20-1]
    matricula2_prueba_p95_bajada_20 = fila_bajada_p95_20['Velocidad de bajada [Kbps]']
    if matricula2_prueba_p95_bajada_20 >= matricula2_velocidad_bajada:
       matricula2_resultado_p95_bajada_20='CUMPLE'
    else:
       matricula2_resultado_p95_bajada_20='NO CUMPLE'

#matricula3

matricula3_rango6 = matricula3[(matricula3['Rango'] == 6)]

#subida 0.05
num_filas_6 = (len(matricula3_rango6))
matricula3_posicion_prueba_p5_subida_6 = math.ceil(num_filas_6*0.05)

matricula3_rango6=matricula3_rango6.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_6 = matricula3_rango6.iloc[matricula3_posicion_prueba_p5_subida_6-1]
matricula3_prueba_p5_subida_6 = fila_subida_p5_6['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_6 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_6='CUMPLE'
else:
   matricula3_resultado_p5_subida_6='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_6 = math.ceil(num_filas_6*0.95)
fila_subida_p95_6 = matricula3_rango6.iloc[matricula3_posicion_prueba_p95_subida_6-1]
matricula3_prueba_p95_subida_6 = fila_subida_p95_6['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_6 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_6='CUMPLE'
else:
   matricula3_resultado_p95_subida_6='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_6 = math.ceil(num_filas_6*0.05)
matricula3_rango6=matricula3_rango6.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_6 = matricula3_rango6.iloc[matricula3_posicion_prueba_p5_bajada_6-1]
matricula3_prueba_p5_bajada_6 = fila_bajada_p5_6['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_6 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_6='CUMPLE'
else:
   matricula3_resultado_p5_bajada_6='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_6 = math.ceil(num_filas_6*0.95)
fila_bajada_p95_6 = matricula3_rango6.iloc[matricula3_posicion_prueba_p95_bajada_6-1]
matricula3_prueba_p95_bajada_6 = fila_bajada_p95_6['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_6 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_6='CUMPLE'
else:
   matricula3_resultado_p95_bajada_6='NO CUMPLE'


matricula3_rango7 = matricula3[(matricula3['Rango'] == 7)]


#subida 0.05
num_filas_7 = (len(matricula3_rango7))
matricula3_posicion_prueba_p5_subida_7 = math.ceil(num_filas_7*0.05)
matricula3_rango7=matricula3_rango7.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_7 = matricula3_rango7.iloc[matricula3_posicion_prueba_p5_subida_7-1]
matricula3_prueba_p5_subida_7 = fila_subida_p5_7['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_7 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_7='CUMPLE'
else:
   matricula3_resultado_p5_subida_7='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_7 = math.ceil(num_filas_7*0.95)
fila_subida_p95_7 = matricula3_rango7.iloc[matricula3_posicion_prueba_p95_subida_7-1]
matricula3_prueba_p95_subida_7 = fila_subida_p95_7['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_7 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_7='CUMPLE'
else:
   matricula3_resultado_p95_subida_7='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_7 = math.ceil(num_filas_7*0.05)
matricula3_rango7=matricula3_rango7.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_7 = matricula3_rango7.iloc[matricula3_posicion_prueba_p5_bajada_7-1]
matricula3_prueba_p5_bajada_7 = fila_bajada_p5_7['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_7 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_7='CUMPLE'
else:
   matricula3_resultado_p5_bajada_7='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_7 = math.ceil(num_filas_7*0.95)
fila_bajada_p95_7 = matricula3_rango7.iloc[matricula3_posicion_prueba_p95_bajada_7-1]
matricula3_prueba_p95_bajada_7 = fila_bajada_p95_7['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_7 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_7='CUMPLE'
else:
   matricula3_resultado_p95_bajada_7='NO CUMPLE'



matricula3_rango8 = matricula3[(matricula3['Rango'] == 8)]


#subida 0.05
num_filas_8 = (len(matricula3_rango8))
matricula3_posicion_prueba_p5_subida_8 = math.ceil(num_filas_8*0.05)
matricula3_rango8=matricula3_rango8.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_8 = matricula3_rango8.iloc[matricula3_posicion_prueba_p5_subida_8-1]
matricula3_prueba_p5_subida_8 = fila_subida_p5_8['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_8 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_8='CUMPLE'
else:
   matricula3_resultado_p5_subida_8='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_8 = math.ceil(num_filas_8*0.95)
fila_subida_p95_8 = matricula3_rango8.iloc[matricula3_posicion_prueba_p95_subida_8-1]
matricula3_prueba_p95_subida_8 = fila_subida_p95_8['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_8 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_8='CUMPLE'
else:
   matricula3_resultado_p95_subida_8='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_8 = math.ceil(num_filas_8*0.05)
matricula3_rango8=matricula3_rango8.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_8 = matricula3_rango8.iloc[matricula3_posicion_prueba_p5_bajada_8-1]
matricula3_prueba_p5_bajada_8 = fila_bajada_p5_8['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_8 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_8='CUMPLE'
else:
   matricula3_resultado_p5_bajada_8='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_8 = math.ceil(num_filas_8*0.95)
fila_bajada_p95_8 = matricula3_rango8.iloc[matricula3_posicion_prueba_p95_bajada_8-1]
matricula3_prueba_p95_bajada_8 = fila_bajada_p95_8['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_8 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_8='CUMPLE'
else:
   matricula3_resultado_p95_bajada_8='NO CUMPLE'



matricula3_rango9 = matricula3[(matricula3['Rango'] == 9)]


#subida 0.05
num_filas_9 = (len(matricula3_rango9))
matricula3_posicion_prueba_p5_subida_9 = math.ceil(num_filas_9*0.05)
matricula3_rango9=matricula3_rango9.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_9 = matricula3_rango9.iloc[matricula3_posicion_prueba_p5_subida_9-1]
matricula3_prueba_p5_subida_9 = fila_subida_p5_9['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_9 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_9='CUMPLE'
else:
   matricula3_resultado_p5_subida_9='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_9 = math.ceil(num_filas_9*0.95)
fila_subida_p95_9 = matricula3_rango9.iloc[matricula3_posicion_prueba_p95_subida_9-1]
matricula3_prueba_p95_subida_9 = fila_subida_p95_9['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_9 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_9='CUMPLE'
else:
   matricula3_resultado_p95_subida_9='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_9 = math.ceil(num_filas_9*0.05)
matricula3_rango9=matricula3_rango9.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_9 = matricula3_rango9.iloc[matricula3_posicion_prueba_p5_bajada_9-1]
matricula3_prueba_p5_bajada_9 = fila_bajada_p5_9['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_9 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_9='CUMPLE'
else:
   matricula3_resultado_p5_bajada_9='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_9 = math.ceil(num_filas_9*0.95)
fila_bajada_p95_9 = matricula3_rango9.iloc[matricula3_posicion_prueba_p95_bajada_9-1]
matricula3_prueba_p95_bajada_9 = fila_bajada_p95_9['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_9 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_9='CUMPLE'
else:
   matricula3_resultado_p95_bajada_9='NO CUMPLE'



matricula3_rango10 = matricula3[(matricula3['Rango'] == 10)]


#subida 0.05
num_filas_10 = (len(matricula3_rango10))
matricula3_posicion_prueba_p5_subida_10 = math.ceil(num_filas_10*0.05)
matricula3_rango10=matricula3_rango10.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_10 = matricula3_rango10.iloc[matricula3_posicion_prueba_p5_subida_10-1]
matricula3_prueba_p5_subida_10 = fila_subida_p5_10['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_10 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_10='CUMPLE'
else:
   matricula3_resultado_p5_subida_10='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_10 = math.ceil(num_filas_10*0.95)
fila_subida_p95_10 = matricula3_rango10.iloc[matricula3_posicion_prueba_p95_subida_10-1]
matricula3_prueba_p95_subida_10 = fila_subida_p95_10['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_10 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_10='CUMPLE'
else:
   matricula3_resultado_p105_subida_10='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_10 = math.ceil(num_filas_10*0.05)
matricula3_rango10=matricula3_rango10.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_10 = matricula3_rango10.iloc[matricula3_posicion_prueba_p5_bajada_10-1]
matricula3_prueba_p5_bajada_10 = fila_bajada_p5_10['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_10 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_10='CUMPLE'
else:
   matricula3_resultado_p5_bajada_10='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_10 = math.ceil(num_filas_10*0.95)
fila_bajada_p95_10 = matricula3_rango10.iloc[matricula3_posicion_prueba_p95_bajada_10-1]
matricula3_prueba_p95_bajada_10 = fila_bajada_p95_10['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_10 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_10='CUMPLE'
else:
   matricula3_resultado_p95_bajada_10='NO CUMPLE'



matricula3_rango11 = matricula3[(matricula3['Rango'] == 11)]


#subida 0.05
num_filas_11 = (len(matricula3_rango11))
matricula3_posicion_prueba_p5_subida_11 = math.ceil(num_filas_11*0.05)
matricula3_rango11=matricula3_rango11.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_11 = matricula3_rango11.iloc[matricula3_posicion_prueba_p5_subida_11-1]
matricula3_prueba_p5_subida_11 = fila_subida_p5_11['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_11 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_11='CUMPLE'
else:
   matricula3_resultado_p5_subida_11='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_11 = math.ceil(num_filas_11*0.95)
fila_subida_p95_11 = matricula3_rango11.iloc[matricula3_posicion_prueba_p95_subida_11-1]
matricula3_prueba_p95_subida_11 = fila_subida_p95_11['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_11 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_11='CUMPLE'
else:
   matricula3_resultado_p95_subida_11='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_11 = math.ceil(num_filas_11*0.05)
matricula3_rango11=matricula3_rango11.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_11 = matricula3_rango11.iloc[matricula3_posicion_prueba_p5_bajada_11-1]
matricula3_prueba_p5_bajada_11 = fila_bajada_p5_11['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_11 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_11='CUMPLE'
else:
   matricula3_resultado_p5_bajada_11='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_11 = math.ceil(num_filas_11*0.95)
fila_bajada_p95_11 = matricula3_rango11.iloc[matricula3_posicion_prueba_p95_bajada_11-1]
matricula3_prueba_p95_bajada_11 = fila_bajada_p95_11['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_11 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_11='CUMPLE'
else:
   matricula3_resultado_p95_bajada_11='NO CUMPLE'



matricula3_rango12 = matricula3[(matricula3['Rango'] == 12)]

#subida 0.05
num_filas_12 = (len(matricula3_rango12))
matricula3_posicion_prueba_p5_subida_12 = math.ceil(num_filas_12*0.05)
matricula3_rango12=matricula3_rango12.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_12 = matricula3_rango12.iloc[matricula3_posicion_prueba_p5_subida_12-1]
matricula3_prueba_p5_subida_12 = fila_subida_p5_12['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_12 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_12='CUMPLE'
else:
   matricula3_resultado_p5_subida_12='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_12 = math.ceil(num_filas_12*0.95)
fila_subida_p95_12 = matricula3_rango12.iloc[matricula3_posicion_prueba_p95_subida_12-1]
matricula3_prueba_p95_subida_12 = fila_subida_p95_12['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_12 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_12='CUMPLE'
else:
   matricula3_resultado_p95_subida_12='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_12 = math.ceil(num_filas_12*0.05)
matricula3_rango12=matricula3_rango12.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_12 = matricula3_rango12.iloc[matricula3_posicion_prueba_p5_bajada_12-1]
matricula3_prueba_p5_bajada_12 = fila_bajada_p5_12['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_12 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_12='CUMPLE'
else:
   matricula3_resultado_p5_bajada_12='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_12 = math.ceil(num_filas_12*0.95)
fila_bajada_p95_12 = matricula3_rango12.iloc[matricula3_posicion_prueba_p95_bajada_12-1]
matricula3_prueba_p95_bajada_12 = fila_bajada_p95_12['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_12 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_12='CUMPLE'
else:
   matricula3_resultado_p95_bajada_12='NO CUMPLE'



matricula3_rango13 = matricula3[(matricula3['Rango'] == 13)]

#subida 0.05
num_filas_13 = (len(matricula3_rango13))
matricula3_posicion_prueba_p5_subida_13 = math.ceil(num_filas_13*0.05)
matricula3_rango13=matricula3_rango13.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_13 = matricula3_rango13.iloc[matricula3_posicion_prueba_p5_subida_13-1]
matricula3_prueba_p5_subida_13 = fila_subida_p5_13['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_13 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_13='CUMPLE'
else:
   matricula3_resultado_p5_subida_13='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_13 = math.ceil(num_filas_13*0.95)
fila_subida_p95_13 = matricula3_rango13.iloc[matricula3_posicion_prueba_p95_subida_13-1]
matricula3_prueba_p95_subida_13 = fila_subida_p95_13['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_13 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_13='CUMPLE'
else:
   matricula3_resultado_p95_subida_13='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_13 = math.ceil(num_filas_13*0.05)
matricula3_rango13=matricula3_rango13.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_13 = matricula3_rango13.iloc[matricula3_posicion_prueba_p5_bajada_13-1]
matricula3_prueba_p5_bajada_13 = fila_bajada_p5_13['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_13 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_13='CUMPLE'
else:
   matricula3_resultado_p5_bajada_13='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_13 = math.ceil(num_filas_13*0.95)
fila_bajada_p95_13 = matricula3_rango13.iloc[matricula3_posicion_prueba_p95_bajada_13-1]
matricula3_prueba_p95_bajada_13 = fila_bajada_p95_13['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_13 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_13='CUMPLE'
else:
   matricula3_resultado_p95_bajada_13='NO CUMPLE'



matricula3_rango14 = matricula3[(matricula3['Rango'] == 14)]

#subida 0.05
num_filas_14 = (len(matricula3_rango14))
matricula3_posicion_prueba_p5_subida_14 = math.ceil(num_filas_14*0.05)
matricula3_rango14=matricula3_rango14.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_14 = matricula3_rango14.iloc[matricula3_posicion_prueba_p5_subida_14-1]
matricula3_prueba_p5_subida_14 = fila_subida_p5_14['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_14 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_14='CUMPLE'
else:
   matricula3_resultado_p5_subida_14='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_14 = math.ceil(num_filas_14*0.95)
fila_subida_p95_14 = matricula3_rango14.iloc[matricula3_posicion_prueba_p95_subida_14-1]
matricula3_prueba_p95_subida_14 = fila_subida_p95_14['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_14 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_14='CUMPLE'
else:
   matricula3_resultado_p95_subida_14='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_14 = math.ceil(num_filas_14*0.05)
matricula3_rango14=matricula3_rango14.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_14 = matricula3_rango14.iloc[matricula3_posicion_prueba_p5_bajada_14-1]
matricula3_prueba_p5_bajada_14 = fila_bajada_p5_14['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_14 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_14='CUMPLE'
else:
   matricula3_resultado_p5_bajada_14='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_14 = math.ceil(num_filas_14*0.95)
fila_bajada_p95_14 = matricula3_rango14.iloc[matricula3_posicion_prueba_p95_bajada_14-1]
matricula3_prueba_p95_bajada_14 = fila_bajada_p95_14['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_14 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_14='CUMPLE'
else:
   matricula3_resultado_p145_bajada_14='NO CUMPLE'




matricula3_rango15 = matricula3[(matricula3['Rango'] == 15)]

#subida 0.05
num_filas_15 = (len(matricula3_rango15))
matricula3_posicion_prueba_p5_subida_15 = math.ceil(num_filas_15*0.05)
matricula3_rango15=matricula3_rango15.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_15 = matricula3_rango15.iloc[matricula3_posicion_prueba_p5_subida_15-1]
matricula3_prueba_p5_subida_15 = fila_subida_p5_15['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_15 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_15='CUMPLE'
else:
   matricula3_resultado_p5_subida_15='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_15 = math.ceil(num_filas_15*0.95)
fila_subida_p95_15 = matricula3_rango15.iloc[matricula3_posicion_prueba_p95_subida_15-1]
matricula3_prueba_p95_subida_15 = fila_subida_p95_15['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_15 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_15='CUMPLE'
else:
   matricula3_resultado_p155_subida_15='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_15 = math.ceil(num_filas_15*0.05)
matricula3_rango15=matricula3_rango15.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_15 = matricula3_rango15.iloc[matricula3_posicion_prueba_p5_bajada_15-1]
matricula3_prueba_p5_bajada_15 = fila_bajada_p5_15['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_15 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_15='CUMPLE'
else:
   matricula3_resultado_p5_bajada_15='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_15 = math.ceil(num_filas_15*0.95)
fila_bajada_p95_15 = matricula3_rango15.iloc[matricula3_posicion_prueba_p95_bajada_15-1]
matricula3_prueba_p95_bajada_15 = fila_bajada_p95_15['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_15 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_15='CUMPLE'
else:
   matricula3_resultado_p155_bajada_15='NO CUMPLE'




matricula3_rango16 = matricula3[(matricula3['Rango'] == 16)]


#subida 0.05
num_filas_16 = (len(matricula3_rango16))
matricula3_posicion_prueba_p5_subida_16 = math.ceil(num_filas_16*0.05)
matricula3_rango16=matricula3_rango16.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_16 = matricula3_rango16.iloc[matricula3_posicion_prueba_p5_subida_16-1]
matricula3_prueba_p5_subida_16 = fila_subida_p5_16['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_16 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_16='CUMPLE'
else:
   matricula3_resultado_p5_subida_16='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_16 = math.ceil(num_filas_16*0.95)
fila_subida_p95_16 = matricula3_rango16.iloc[matricula3_posicion_prueba_p95_subida_16-1]
matricula3_prueba_p95_subida_16 = fila_subida_p95_16['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_16 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_16='CUMPLE'
else:
   matricula3_resultado_p95_subida_16='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_16 = math.ceil(num_filas_16*0.05)
matricula3_rango16=matricula3_rango16.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_16 = matricula3_rango16.iloc[matricula3_posicion_prueba_p5_bajada_16-1]
matricula3_prueba_p5_bajada_16 = fila_bajada_p5_16['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_16 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_16='CUMPLE'
else:
   matricula3_resultado_p5_bajada_16='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_16 = math.ceil(num_filas_16*0.95)
fila_bajada_p95_16 = matricula3_rango16.iloc[matricula3_posicion_prueba_p95_bajada_16-1]
matricula3_prueba_p95_bajada_16 = fila_bajada_p95_16['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_16 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_16='CUMPLE'
else:
   matricula3_resultado_p165_bajada_16='NO CUMPLE'




matricula3_rango17 = matricula3[(matricula3['Rango'] == 17)]


#subida 0.05
num_filas_17 = (len(matricula3_rango17))
matricula3_posicion_prueba_p5_subida_17 = math.ceil(num_filas_17*0.05)
matricula3_rango17=matricula3_rango17.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_17 = matricula3_rango17.iloc[matricula3_posicion_prueba_p5_subida_17-1]
matricula3_prueba_p5_subida_17 = fila_subida_p5_17['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_17 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_17='CUMPLE'
else:
   matricula3_resultado_p5_subida_17='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_17 = math.ceil(num_filas_17*0.95)
fila_subida_p95_17 = matricula3_rango17.iloc[matricula3_posicion_prueba_p95_subida_17-1]
matricula3_prueba_p95_subida_17 = fila_subida_p95_17['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_17 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_17='CUMPLE'
else:
   matricula3_resultado_p95_subida_17='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_17 = math.ceil(num_filas_17*0.05)
matricula3_rango17=matricula3_rango17.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_17 = matricula3_rango17.iloc[matricula3_posicion_prueba_p5_bajada_17-1]
matricula3_prueba_p5_bajada_17 = fila_bajada_p5_17['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_17 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_17='CUMPLE'
else:
   matricula3_resultado_p5_bajada_17='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_17 = math.ceil(num_filas_17*0.95)
fila_bajada_p95_17 = matricula3_rango17.iloc[matricula3_posicion_prueba_p95_bajada_17-1]
matricula3_prueba_p95_bajada_17 = fila_bajada_p95_17['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_17 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_17='CUMPLE'
else:
   matricula3_resultado_p95_bajada_17='NO CUMPLE'



matricula3_rango18 = matricula3[(matricula3['Rango'] == 18)]

#subida 0.05
num_filas_18 = (len(matricula3_rango18))
matricula3_posicion_prueba_p5_subida_18 = math.ceil(num_filas_18*0.05)
matricula3_rango18=matricula3_rango18.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_18 = matricula3_rango18.iloc[matricula3_posicion_prueba_p5_subida_18-1]
matricula3_prueba_p5_subida_18 = fila_subida_p5_18['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_18 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_18='CUMPLE'
else:
   matricula3_resultado_p5_subida_18='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_18 = math.ceil(num_filas_18*0.95)
fila_subida_p95_18 = matricula3_rango18.iloc[matricula3_posicion_prueba_p95_subida_18-1]
matricula3_prueba_p95_subida_18 = fila_subida_p95_18['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_18 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_18='CUMPLE'
else:
   matricula3_resultado_p95_subida_18='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_18 = math.ceil(num_filas_18*0.05)
matricula3_rango18=matricula3_rango18.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_18 = matricula3_rango18.iloc[matricula3_posicion_prueba_p5_bajada_18-1]
matricula3_prueba_p5_bajada_18 = fila_bajada_p5_18['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_18 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_18='CUMPLE'
else:
   matricula3_resultado_p5_bajada_18='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_18 = math.ceil(num_filas_18*0.95)
fila_bajada_p95_18 = matricula3_rango18.iloc[matricula3_posicion_prueba_p95_bajada_18-1]
matricula3_prueba_p95_bajada_18 = fila_bajada_p95_18['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_18 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_18='CUMPLE'
else:
   matricula3_resultado_p95_bajada_18='NO CUMPLE'



matricula3_rango19 = matricula3[(matricula3['Rango'] == 19)]


#subida 0.05
num_filas_19 = (len(matricula3_rango19))
matricula3_posicion_prueba_p5_subida_19 = math.ceil(num_filas_19*0.05)
matricula3_rango19=matricula3_rango19.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_19 = matricula3_rango19.iloc[matricula3_posicion_prueba_p5_subida_19-1]
matricula3_prueba_p5_subida_19 = fila_subida_p5_19['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_19 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_19='CUMPLE'
else:
   matricula3_resultado_p5_subida_19='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_19 = math.ceil(num_filas_19*0.95)
fila_subida_p95_19 = matricula3_rango19.iloc[matricula3_posicion_prueba_p95_subida_19-1]
matricula3_prueba_p95_subida_19 = fila_subida_p95_19['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_19 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_19='CUMPLE'
else:
   matricula3_resultado_p95_subida_19='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_19 = math.ceil(num_filas_19*0.05)
matricula3_rango19=matricula3_rango19.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_19 = matricula3_rango19.iloc[matricula3_posicion_prueba_p5_bajada_19-1]
matricula3_prueba_p5_bajada_19 = fila_bajada_p5_19['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_19 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_19='CUMPLE'
else:
   matricula3_resultado_p5_bajada_19='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_19 = math.ceil(num_filas_19*0.95)
fila_bajada_p95_19 = matricula3_rango19.iloc[matricula3_posicion_prueba_p95_bajada_19-1]
matricula3_prueba_p95_bajada_19 = fila_bajada_p95_19['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_19 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_19='CUMPLE'
else:
   matricula3_resultado_p95_bajada_19='NO CUMPLE'



matricula3_rango20 = matricula3[(matricula3['Rango'] == 20)]

#subida 0.05
num_filas_20 = (len(matricula3_rango20))
matricula3_posicion_prueba_p5_subida_20 = math.ceil(num_filas_20*0.05)
matricula3_rango20=matricula3_rango20.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_20 = matricula3_rango20.iloc[matricula3_posicion_prueba_p5_subida_20-1]
matricula3_prueba_p5_subida_20 = fila_subida_p5_20['Velocidad de subida [Kbps]']
if matricula3_prueba_p5_subida_20 >= matricula3_velocidad_subida:
   matricula3_resultado_p5_subida_20='CUMPLE'
else:
   matricula3_resultado_p5_subida_20='NO CUMPLE'
#subida 0.95
matricula3_posicion_prueba_p95_subida_20 = math.ceil(num_filas_20*0.95)
fila_subida_p95_20 = matricula3_rango20.iloc[matricula3_posicion_prueba_p95_subida_20-1]
matricula3_prueba_p95_subida_20 = fila_subida_p95_20['Velocidad de subida [Kbps]']
if matricula3_prueba_p95_subida_20 >= matricula3_velocidad_subida:
   matricula3_resultado_p95_subida_20='CUMPLE'
else:
   matricula3_resultado_p95_subida_20='NO CUMPLE'

#bajada 0.05
matricula3_posicion_prueba_p5_bajada_20 = math.ceil(num_filas_20*0.05)
matricula3_rango20=matricula3_rango20.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_20 = matricula3_rango20.iloc[matricula3_posicion_prueba_p5_bajada_20-1]
matricula3_prueba_p5_bajada_20 = fila_bajada_p5_20['Velocidad de bajada [Kbps]']
if matricula3_prueba_p5_bajada_20 >= matricula3_velocidad_bajada:
   matricula3_resultado_p5_bajada_20='CUMPLE'
else:
   matricula3_resultado_p5_bajada_20='NO CUMPLE'
#bajada 0.95
matricula3_posicion_prueba_p95_bajada_20 = math.ceil(num_filas_20*0.95)
fila_bajada_p95_20 = matricula3_rango20.iloc[matricula3_posicion_prueba_p95_bajada_20-1]
matricula3_prueba_p95_bajada_20 = fila_bajada_p95_20['Velocidad de bajada [Kbps]']
if matricula3_prueba_p95_bajada_20 >= matricula3_velocidad_bajada:
   matricula3_resultado_p95_bajada_20='CUMPLE'
else:
   matricula3_resultado_p95_bajada_20='NO CUMPLE'

#matricula4

matricula4_rango6 = matricula4[(matricula4['Rango'] == 6)]

#subida 0.05
num_filas_6 = (len(matricula4_rango6))
matricula4_posicion_prueba_p5_subida_6 = math.ceil(num_filas_6*0.05)

matricula4_rango6=matricula4_rango6.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_6 = matricula4_rango6.iloc[matricula4_posicion_prueba_p5_subida_6-1]
matricula4_prueba_p5_subida_6 = fila_subida_p5_6['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_6 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_6='CUMPLE'
else:
   matricula4_resultado_p5_subida_6='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_6 = math.ceil(num_filas_6*0.95)
fila_subida_p95_6 = matricula4_rango6.iloc[matricula4_posicion_prueba_p95_subida_6-1]
matricula4_prueba_p95_subida_6 = fila_subida_p95_6['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_6 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_6='CUMPLE'
else:
   matricula4_resultado_p95_subida_6='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_6 = math.ceil(num_filas_6*0.05)
matricula4_rango6=matricula4_rango6.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_6 = matricula4_rango6.iloc[matricula4_posicion_prueba_p5_bajada_6-1]
matricula4_prueba_p5_bajada_6 = fila_bajada_p5_6['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_6 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_6='CUMPLE'
else:
   matricula4_resultado_p5_bajada_6='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_6 = math.ceil(num_filas_6*0.95)
fila_bajada_p95_6 = matricula4_rango6.iloc[matricula4_posicion_prueba_p95_bajada_6-1]
matricula4_prueba_p95_bajada_6 = fila_bajada_p95_6['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_6 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_6='CUMPLE'
else:
   matricula4_resultado_p95_bajada_6='NO CUMPLE'


matricula4_rango7 = matricula4[(matricula4['Rango'] == 7)]


#subida 0.05
num_filas_7 = (len(matricula4_rango7))
matricula4_posicion_prueba_p5_subida_7 = math.ceil(num_filas_7*0.05)
matricula4_rango7=matricula4_rango7.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_7 = matricula4_rango7.iloc[matricula4_posicion_prueba_p5_subida_7-1]
matricula4_prueba_p5_subida_7 = fila_subida_p5_7['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_7 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_7='CUMPLE'
else:
   matricula4_resultado_p5_subida_7='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_7 = math.ceil(num_filas_7*0.95)
fila_subida_p95_7 = matricula4_rango7.iloc[matricula4_posicion_prueba_p95_subida_7-1]
matricula4_prueba_p95_subida_7 = fila_subida_p95_7['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_7 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_7='CUMPLE'
else:
   matricula4_resultado_p95_subida_7='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_7 = math.ceil(num_filas_7*0.05)
matricula4_rango7=matricula4_rango7.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_7 = matricula4_rango7.iloc[matricula4_posicion_prueba_p5_bajada_7-1]
matricula4_prueba_p5_bajada_7 = fila_bajada_p5_7['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_7 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_7='CUMPLE'
else:
   matricula4_resultado_p5_bajada_7='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_7 = math.ceil(num_filas_7*0.95)
fila_bajada_p95_7 = matricula4_rango7.iloc[matricula4_posicion_prueba_p95_bajada_7-1]
matricula4_prueba_p95_bajada_7 = fila_bajada_p95_7['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_7 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_7='CUMPLE'
else:
   matricula4_resultado_p95_bajada_7='NO CUMPLE'



matricula4_rango8 = matricula4[(matricula4['Rango'] == 8)]


#subida 0.05
num_filas_8 = (len(matricula4_rango8))
matricula4_posicion_prueba_p5_subida_8 = math.ceil(num_filas_8*0.05)
matricula4_rango8=matricula4_rango8.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_8 = matricula4_rango8.iloc[matricula4_posicion_prueba_p5_subida_8-1]
matricula4_prueba_p5_subida_8 = fila_subida_p5_8['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_8 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_8='CUMPLE'
else:
   matricula4_resultado_p5_subida_8='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_8 = math.ceil(num_filas_8*0.95)
fila_subida_p95_8 = matricula4_rango8.iloc[matricula4_posicion_prueba_p95_subida_8-1]
matricula4_prueba_p95_subida_8 = fila_subida_p95_8['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_8 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_8='CUMPLE'
else:
   matricula4_resultado_p95_subida_8='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_8 = math.ceil(num_filas_8*0.05)
matricula4_rango8=matricula4_rango8.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_8 = matricula4_rango8.iloc[matricula4_posicion_prueba_p5_bajada_8-1]
matricula4_prueba_p5_bajada_8 = fila_bajada_p5_8['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_8 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_8='CUMPLE'
else:
   matricula4_resultado_p5_bajada_8='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_8 = math.ceil(num_filas_8*0.95)
fila_bajada_p95_8 = matricula4_rango8.iloc[matricula4_posicion_prueba_p95_bajada_8-1]
matricula4_prueba_p95_bajada_8 = fila_bajada_p95_8['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_8 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_8='CUMPLE'
else:
   matricula4_resultado_p95_bajada_8='NO CUMPLE'



matricula4_rango9 = matricula4[(matricula4['Rango'] == 9)]


#subida 0.05
num_filas_9 = (len(matricula4_rango9))
matricula4_posicion_prueba_p5_subida_9 = math.ceil(num_filas_9*0.05)
matricula4_rango9=matricula4_rango9.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_9 = matricula4_rango9.iloc[matricula4_posicion_prueba_p5_subida_9-1]
matricula4_prueba_p5_subida_9 = fila_subida_p5_9['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_9 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_9='CUMPLE'
else:
   matricula4_resultado_p5_subida_9='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_9 = math.ceil(num_filas_9*0.95)
fila_subida_p95_9 = matricula4_rango9.iloc[matricula4_posicion_prueba_p95_subida_9-1]
matricula4_prueba_p95_subida_9 = fila_subida_p95_9['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_9 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_9='CUMPLE'
else:
   matricula4_resultado_p95_subida_9='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_9 = math.ceil(num_filas_9*0.05)
matricula4_rango9=matricula4_rango9.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_9 = matricula4_rango9.iloc[matricula4_posicion_prueba_p5_bajada_9-1]
matricula4_prueba_p5_bajada_9 = fila_bajada_p5_9['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_9 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_9='CUMPLE'
else:
   matricula4_resultado_p5_bajada_9='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_9 = math.ceil(num_filas_9*0.95)
fila_bajada_p95_9 = matricula4_rango9.iloc[matricula4_posicion_prueba_p95_bajada_9-1]
matricula4_prueba_p95_bajada_9 = fila_bajada_p95_9['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_9 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_9='CUMPLE'
else:
   matricula4_resultado_p95_bajada_9='NO CUMPLE'



matricula4_rango10 = matricula4[(matricula4['Rango'] == 10)]


#subida 0.05
num_filas_10 = (len(matricula4_rango10))
matricula4_posicion_prueba_p5_subida_10 = math.ceil(num_filas_10*0.05)
matricula4_rango10=matricula4_rango10.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_10 = matricula4_rango10.iloc[matricula4_posicion_prueba_p5_subida_10-1]
matricula4_prueba_p5_subida_10 = fila_subida_p5_10['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_10 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_10='CUMPLE'
else:
   matricula4_resultado_p5_subida_10='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_10 = math.ceil(num_filas_10*0.95)
fila_subida_p95_10 = matricula4_rango10.iloc[matricula4_posicion_prueba_p95_subida_10-1]
matricula4_prueba_p95_subida_10 = fila_subida_p95_10['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_10 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_10='CUMPLE'
else:
   matricula4_resultado_p105_subida_10='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_10 = math.ceil(num_filas_10*0.05)
matricula4_rango10=matricula4_rango10.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_10 = matricula4_rango10.iloc[matricula4_posicion_prueba_p5_bajada_10-1]
matricula4_prueba_p5_bajada_10 = fila_bajada_p5_10['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_10 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_10='CUMPLE'
else:
   matricula4_resultado_p5_bajada_10='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_10 = math.ceil(num_filas_10*0.95)
fila_bajada_p95_10 = matricula4_rango10.iloc[matricula4_posicion_prueba_p95_bajada_10-1]
matricula4_prueba_p95_bajada_10 = fila_bajada_p95_10['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_10 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_10='CUMPLE'
else:
   matricula4_resultado_p95_bajada_10='NO CUMPLE'



matricula4_rango11 = matricula4[(matricula4['Rango'] == 11)]


#subida 0.05
num_filas_11 = (len(matricula4_rango11))
matricula4_posicion_prueba_p5_subida_11 = math.ceil(num_filas_11*0.05)
matricula4_rango11=matricula4_rango11.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_11 = matricula4_rango11.iloc[matricula4_posicion_prueba_p5_subida_11-1]
matricula4_prueba_p5_subida_11 = fila_subida_p5_11['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_11 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_11='CUMPLE'
else:
   matricula4_resultado_p5_subida_11='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_11 = math.ceil(num_filas_11*0.95)
fila_subida_p95_11 = matricula4_rango11.iloc[matricula4_posicion_prueba_p95_subida_11-1]
matricula4_prueba_p95_subida_11 = fila_subida_p95_11['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_11 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_11='CUMPLE'
else:
   matricula4_resultado_p95_subida_11='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_11 = math.ceil(num_filas_11*0.05)
matricula4_rango11=matricula4_rango11.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_11 = matricula4_rango11.iloc[matricula4_posicion_prueba_p5_bajada_11-1]
matricula4_prueba_p5_bajada_11 = fila_bajada_p5_11['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_11 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_11='CUMPLE'
else:
   matricula4_resultado_p5_bajada_11='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_11 = math.ceil(num_filas_11*0.95)
fila_bajada_p95_11 = matricula4_rango11.iloc[matricula4_posicion_prueba_p95_bajada_11-1]
matricula4_prueba_p95_bajada_11 = fila_bajada_p95_11['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_11 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_11='CUMPLE'
else:
   matricula4_resultado_p95_bajada_11='NO CUMPLE'



matricula4_rango12 = matricula4[(matricula4['Rango'] == 12)]

#subida 0.05
num_filas_12 = (len(matricula4_rango12))
matricula4_posicion_prueba_p5_subida_12 = math.ceil(num_filas_12*0.05)
matricula4_rango12=matricula4_rango12.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_12 = matricula4_rango12.iloc[matricula4_posicion_prueba_p5_subida_12-1]
matricula4_prueba_p5_subida_12 = fila_subida_p5_12['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_12 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_12='CUMPLE'
else:
   matricula4_resultado_p5_subida_12='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_12 = math.ceil(num_filas_12*0.95)
fila_subida_p95_12 = matricula4_rango12.iloc[matricula4_posicion_prueba_p95_subida_12-1]
matricula4_prueba_p95_subida_12 = fila_subida_p95_12['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_12 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_12='CUMPLE'
else:
   matricula4_resultado_p95_subida_12='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_12 = math.ceil(num_filas_12*0.05)
matricula4_rango12=matricula4_rango12.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_12 = matricula4_rango12.iloc[matricula4_posicion_prueba_p5_bajada_12-1]
matricula4_prueba_p5_bajada_12 = fila_bajada_p5_12['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_12 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_12='CUMPLE'
else:
   matricula4_resultado_p5_bajada_12='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_12 = math.ceil(num_filas_12*0.95)
fila_bajada_p95_12 = matricula4_rango12.iloc[matricula4_posicion_prueba_p95_bajada_12-1]
matricula4_prueba_p95_bajada_12 = fila_bajada_p95_12['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_12 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_12='CUMPLE'
else:
   matricula4_resultado_p95_bajada_12='NO CUMPLE'



matricula4_rango13 = matricula4[(matricula4['Rango'] == 13)]

#subida 0.05
num_filas_13 = (len(matricula4_rango13))
matricula4_posicion_prueba_p5_subida_13 = math.ceil(num_filas_13*0.05)
matricula4_rango13=matricula4_rango13.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_13 = matricula4_rango13.iloc[matricula4_posicion_prueba_p5_subida_13-1]
matricula4_prueba_p5_subida_13 = fila_subida_p5_13['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_13 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_13='CUMPLE'
else:
   matricula4_resultado_p5_subida_13='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_13 = math.ceil(num_filas_13*0.95)
fila_subida_p95_13 = matricula4_rango13.iloc[matricula4_posicion_prueba_p95_subida_13-1]
matricula4_prueba_p95_subida_13 = fila_subida_p95_13['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_13 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_13='CUMPLE'
else:
   matricula4_resultado_p95_subida_13='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_13 = math.ceil(num_filas_13*0.05)
matricula4_rango13=matricula4_rango13.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_13 = matricula4_rango13.iloc[matricula4_posicion_prueba_p5_bajada_13-1]
matricula4_prueba_p5_bajada_13 = fila_bajada_p5_13['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_13 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_13='CUMPLE'
else:
   matricula4_resultado_p5_bajada_13='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_13 = math.ceil(num_filas_13*0.95)
fila_bajada_p95_13 = matricula4_rango13.iloc[matricula4_posicion_prueba_p95_bajada_13-1]
matricula4_prueba_p95_bajada_13 = fila_bajada_p95_13['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_13 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_13='CUMPLE'
else:
   matricula4_resultado_p95_bajada_13='NO CUMPLE'



matricula4_rango14 = matricula4[(matricula4['Rango'] == 14)]

#subida 0.05
num_filas_14 = (len(matricula4_rango14))
matricula4_posicion_prueba_p5_subida_14 = math.ceil(num_filas_14*0.05)
matricula4_rango14=matricula4_rango14.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_14 = matricula4_rango14.iloc[matricula4_posicion_prueba_p5_subida_14-1]
matricula4_prueba_p5_subida_14 = fila_subida_p5_14['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_14 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_14='CUMPLE'
else:
   matricula4_resultado_p5_subida_14='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_14 = math.ceil(num_filas_14*0.95)
fila_subida_p95_14 = matricula4_rango14.iloc[matricula4_posicion_prueba_p95_subida_14-1]
matricula4_prueba_p95_subida_14 = fila_subida_p95_14['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_14 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_14='CUMPLE'
else:
   matricula4_resultado_p95_subida_14='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_14 = math.ceil(num_filas_14*0.05)
matricula4_rango14=matricula4_rango14.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_14 = matricula4_rango14.iloc[matricula4_posicion_prueba_p5_bajada_14-1]
matricula4_prueba_p5_bajada_14 = fila_bajada_p5_14['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_14 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_14='CUMPLE'
else:
   matricula4_resultado_p5_bajada_14='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_14 = math.ceil(num_filas_14*0.95)
fila_bajada_p95_14 = matricula4_rango14.iloc[matricula4_posicion_prueba_p95_bajada_14-1]
matricula4_prueba_p95_bajada_14 = fila_bajada_p95_14['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_14 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_14='CUMPLE'
else:
   matricula4_resultado_p145_bajada_14='NO CUMPLE'




matricula4_rango15 = matricula4[(matricula4['Rango'] == 15)]

#subida 0.05
num_filas_15 = (len(matricula4_rango15))
matricula4_posicion_prueba_p5_subida_15 = math.ceil(num_filas_15*0.05)
matricula4_rango15=matricula4_rango15.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_15 = matricula4_rango15.iloc[matricula4_posicion_prueba_p5_subida_15-1]
matricula4_prueba_p5_subida_15 = fila_subida_p5_15['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_15 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_15='CUMPLE'
else:
   matricula4_resultado_p5_subida_15='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_15 = math.ceil(num_filas_15*0.95)
fila_subida_p95_15 = matricula4_rango15.iloc[matricula4_posicion_prueba_p95_subida_15-1]
matricula4_prueba_p95_subida_15 = fila_subida_p95_15['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_15 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_15='CUMPLE'
else:
   matricula4_resultado_p155_subida_15='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_15 = math.ceil(num_filas_15*0.05)
matricula4_rango15=matricula4_rango15.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_15 = matricula4_rango15.iloc[matricula4_posicion_prueba_p5_bajada_15-1]
matricula4_prueba_p5_bajada_15 = fila_bajada_p5_15['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_15 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_15='CUMPLE'
else:
   matricula4_resultado_p5_bajada_15='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_15 = math.ceil(num_filas_15*0.95)
fila_bajada_p95_15 = matricula4_rango15.iloc[matricula4_posicion_prueba_p95_bajada_15-1]
matricula4_prueba_p95_bajada_15 = fila_bajada_p95_15['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_15 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_15='CUMPLE'
else:
   matricula4_resultado_p155_bajada_15='NO CUMPLE'




matricula4_rango16 = matricula4[(matricula4['Rango'] == 16)]


#subida 0.05
num_filas_16 = (len(matricula4_rango16))
matricula4_posicion_prueba_p5_subida_16 = math.ceil(num_filas_16*0.05)
matricula4_rango16=matricula4_rango16.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_16 = matricula4_rango16.iloc[matricula4_posicion_prueba_p5_subida_16-1]
matricula4_prueba_p5_subida_16 = fila_subida_p5_16['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_16 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_16='CUMPLE'
else:
   matricula4_resultado_p5_subida_16='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_16 = math.ceil(num_filas_16*0.95)
fila_subida_p95_16 = matricula4_rango16.iloc[matricula4_posicion_prueba_p95_subida_16-1]
matricula4_prueba_p95_subida_16 = fila_subida_p95_16['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_16 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_16='CUMPLE'
else:
   matricula4_resultado_p95_subida_16='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_16 = math.ceil(num_filas_16*0.05)
matricula4_rango16=matricula4_rango16.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_16 = matricula4_rango16.iloc[matricula4_posicion_prueba_p5_bajada_16-1]
matricula4_prueba_p5_bajada_16 = fila_bajada_p5_16['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_16 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_16='CUMPLE'
else:
   matricula4_resultado_p5_bajada_16='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_16 = math.ceil(num_filas_16*0.95)
fila_bajada_p95_16 = matricula4_rango16.iloc[matricula4_posicion_prueba_p95_bajada_16-1]
matricula4_prueba_p95_bajada_16 = fila_bajada_p95_16['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_16 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_16='CUMPLE'
else:
   matricula4_resultado_p165_bajada_16='NO CUMPLE'




matricula4_rango17 = matricula4[(matricula4['Rango'] == 17)]


#subida 0.05
num_filas_17 = (len(matricula4_rango17))
matricula4_posicion_prueba_p5_subida_17 = math.ceil(num_filas_17*0.05)
matricula4_rango17=matricula4_rango17.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_17 = matricula4_rango17.iloc[matricula4_posicion_prueba_p5_subida_17-1]
matricula4_prueba_p5_subida_17 = fila_subida_p5_17['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_17 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_17='CUMPLE'
else:
   matricula4_resultado_p5_subida_17='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_17 = math.ceil(num_filas_17*0.95)
fila_subida_p95_17 = matricula4_rango17.iloc[matricula4_posicion_prueba_p95_subida_17-1]
matricula4_prueba_p95_subida_17 = fila_subida_p95_17['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_17 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_17='CUMPLE'
else:
   matricula4_resultado_p95_subida_17='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_17 = math.ceil(num_filas_17*0.05)
matricula4_rango17=matricula4_rango17.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_17 = matricula4_rango17.iloc[matricula4_posicion_prueba_p5_bajada_17-1]
matricula4_prueba_p5_bajada_17 = fila_bajada_p5_17['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_17 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_17='CUMPLE'
else:
   matricula4_resultado_p5_bajada_17='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_17 = math.ceil(num_filas_17*0.95)
fila_bajada_p95_17 = matricula4_rango17.iloc[matricula4_posicion_prueba_p95_bajada_17-1]
matricula4_prueba_p95_bajada_17 = fila_bajada_p95_17['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_17 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_17='CUMPLE'
else:
   matricula4_resultado_p95_bajada_17='NO CUMPLE'



matricula4_rango18 = matricula4[(matricula4['Rango'] == 18)]

#subida 0.05
num_filas_18 = (len(matricula4_rango18))
matricula4_posicion_prueba_p5_subida_18 = math.ceil(num_filas_18*0.05)
matricula4_rango18=matricula4_rango18.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_18 = matricula4_rango18.iloc[matricula4_posicion_prueba_p5_subida_18-1]
matricula4_prueba_p5_subida_18 = fila_subida_p5_18['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_18 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_18='CUMPLE'
else:
   matricula4_resultado_p5_subida_18='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_18 = math.ceil(num_filas_18*0.95)
fila_subida_p95_18 = matricula4_rango18.iloc[matricula4_posicion_prueba_p95_subida_18-1]
matricula4_prueba_p95_subida_18 = fila_subida_p95_18['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_18 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_18='CUMPLE'
else:
   matricula4_resultado_p95_subida_18='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_18 = math.ceil(num_filas_18*0.05)
matricula4_rango18=matricula4_rango18.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_18 = matricula4_rango18.iloc[matricula4_posicion_prueba_p5_bajada_18-1]
matricula4_prueba_p5_bajada_18 = fila_bajada_p5_18['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_18 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_18='CUMPLE'
else:
   matricula4_resultado_p5_bajada_18='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_18 = math.ceil(num_filas_18*0.95)
fila_bajada_p95_18 = matricula4_rango18.iloc[matricula4_posicion_prueba_p95_bajada_18-1]
matricula4_prueba_p95_bajada_18 = fila_bajada_p95_18['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_18 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_18='CUMPLE'
else:
   matricula4_resultado_p95_bajada_18='NO CUMPLE'



matricula4_rango19 = matricula4[(matricula4['Rango'] == 19)]


#subida 0.05
num_filas_19 = (len(matricula4_rango19))
matricula4_posicion_prueba_p5_subida_19 = math.ceil(num_filas_19*0.05)
matricula4_rango19=matricula4_rango19.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_19 = matricula4_rango19.iloc[matricula4_posicion_prueba_p5_subida_19-1]
matricula4_prueba_p5_subida_19 = fila_subida_p5_19['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_19 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_19='CUMPLE'
else:
   matricula4_resultado_p5_subida_19='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_19 = math.ceil(num_filas_19*0.95)
fila_subida_p95_19 = matricula4_rango19.iloc[matricula4_posicion_prueba_p95_subida_19-1]
matricula4_prueba_p95_subida_19 = fila_subida_p95_19['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_19 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_19='CUMPLE'
else:
   matricula4_resultado_p95_subida_19='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_19 = math.ceil(num_filas_19*0.05)
matricula4_rango19=matricula4_rango19.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_19 = matricula4_rango19.iloc[matricula4_posicion_prueba_p5_bajada_19-1]
matricula4_prueba_p5_bajada_19 = fila_bajada_p5_19['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_19 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_19='CUMPLE'
else:
   matricula4_resultado_p5_bajada_19='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_19 = math.ceil(num_filas_19*0.95)
fila_bajada_p95_19 = matricula4_rango19.iloc[matricula4_posicion_prueba_p95_bajada_19-1]
matricula4_prueba_p95_bajada_19 = fila_bajada_p95_19['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_19 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_19='CUMPLE'
else:
   matricula4_resultado_p95_bajada_19='NO CUMPLE'



matricula4_rango20 = matricula4[(matricula4['Rango'] == 20)]

#subida 0.05
num_filas_20 = (len(matricula4_rango20))
matricula4_posicion_prueba_p5_subida_20 = math.ceil(num_filas_20*0.05)
matricula4_rango20=matricula4_rango20.sort_values(by='Velocidad de subida [Kbps]')
fila_subida_p5_20 = matricula4_rango20.iloc[matricula4_posicion_prueba_p5_subida_20-1]
matricula4_prueba_p5_subida_20 = fila_subida_p5_20['Velocidad de subida [Kbps]']
if matricula4_prueba_p5_subida_20 >= matricula4_velocidad_subida:
   matricula4_resultado_p5_subida_20='CUMPLE'
else:
   matricula4_resultado_p5_subida_20='NO CUMPLE'
#subida 0.95
matricula4_posicion_prueba_p95_subida_20 = math.ceil(num_filas_20*0.95)
fila_subida_p95_20 = matricula4_rango20.iloc[matricula4_posicion_prueba_p95_subida_20-1]
matricula4_prueba_p95_subida_20 = fila_subida_p95_20['Velocidad de subida [Kbps]']
if matricula4_prueba_p95_subida_20 >= matricula4_velocidad_subida:
   matricula4_resultado_p95_subida_20='CUMPLE'
else:
   matricula4_resultado_p95_subida_20='NO CUMPLE'

#bajada 0.05
matricula4_posicion_prueba_p5_bajada_20 = math.ceil(num_filas_20*0.05)
matricula4_rango20=matricula4_rango20.sort_values(by='Velocidad de bajada [Kbps]')
fila_bajada_p5_20 = matricula4_rango20.iloc[matricula4_posicion_prueba_p5_bajada_20-1]
matricula4_prueba_p5_bajada_20 = fila_bajada_p5_20['Velocidad de bajada [Kbps]']
if matricula4_prueba_p5_bajada_20 >= matricula4_velocidad_bajada:
   matricula4_resultado_p5_bajada_20='CUMPLE'
else:
   matricula4_resultado_p5_bajada_20='NO CUMPLE'
#bajada 0.95
matricula4_posicion_prueba_p95_bajada_20 = math.ceil(num_filas_20*0.95)
fila_bajada_p95_20 = matricula4_rango20.iloc[matricula4_posicion_prueba_p95_bajada_20-1]
matricula4_prueba_p95_bajada_20 = fila_bajada_p95_20['Velocidad de bajada [Kbps]']
if matricula4_prueba_p95_bajada_20 >= matricula4_velocidad_bajada:
   matricula4_resultado_p95_bajada_20='CUMPLE'
else:
   matricula4_resultado_p95_bajada_20='NO CUMPLE'

#Rango Matricula           Velocidad Bajada(Mbps)  	Velocidad Subida(Mbps)
#Matrícula <= 50	              13,44	                    3,36
#51 >= Matrícula <= 150	      16,8                    	4,2
#151 >= Matrícula <= 400     	20,16	                    5,04
#Matrícula > 400	              23,52	                    5,88

#nombres_reales = [variable for variable in globals().keys()]
# Imprimimos la lista ordenadita, una por línea, para que sea fácil de leer
#for nombre in nombres_reales:
#    print(nombre)

# descargar y crear archiv ecxel por pestañas
#matricula1
output_path = '/content/20260617_PERFIL_1.xlsx'
# Crear un archivo Excel con múltiples hojas
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    matricula1.to_excel(writer, sheet_name='PERFIL_1', index=False, startrow=13, startcol=1)
    matricula1_rango6.to_excel(writer, sheet_name='6', index=False, startrow=13, startcol=1)
    matricula1_rango7.to_excel(writer, sheet_name='7', index=False, startrow=13, startcol=1)
    matricula1_rango8.to_excel(writer, sheet_name='8', index=False, startrow=13, startcol=1)
    matricula1_rango9.to_excel(writer, sheet_name='9', index=False, startrow=13, startcol=1)
    matricula1_rango10.to_excel(writer, sheet_name='10', index=False, startrow=13, startcol=1)
    matricula1_rango11.to_excel(writer, sheet_name='11', index=False, startrow=13, startcol=1)
    matricula1_rango12.to_excel(writer, sheet_name='12', index=False, startrow=13, startcol=1)
    matricula1_rango13.to_excel(writer, sheet_name='13', index=False, startrow=13, startcol=1)
    matricula1_rango14.to_excel(writer, sheet_name='14', index=False, startrow=13, startcol=1)
    matricula1_rango15.to_excel(writer, sheet_name='15', index=False, startrow=13, startcol=1)
    matricula1_rango16.to_excel(writer, sheet_name='16', index=False, startrow=13, startcol=1)
    matricula1_rango17.to_excel(writer, sheet_name='17', index=False, startrow=13, startcol=1)
    matricula1_rango18.to_excel(writer, sheet_name='18', index=False, startrow=13, startcol=1)
    matricula1_rango19.to_excel(writer, sheet_name='19', index=False, startrow=13, startcol=1)
    matricula1_rango20.to_excel(writer, sheet_name='20', index=False, startrow=13, startcol=1)
 # Acceder al objeto de libro de trabajo para personalizar
    workbook = writer.book


#print(f"Archivo guardado en: {output_path}")
#wb = load_workbook(output_path)

# Definir el título principal que se quiere poner en todas las hojas
main_title = "2.3.4.4.2 VELOCIDAD EFECTIVA MÍNIMA DE TRANSMISIÓN DE DATOS"
# Definir los subtítulos (para la fila 2 y siguientes)
sub_titles = ["Año-Mes(AAAA-MM)"]
sub_titles2 = ["2026-06"]
sub_titles3 = ["Sentido","Matricula","Velocidad Por Matricula","Posición de la prueba","Prueba P5%","Resultado","Posición de la prueba","Prueba P95%","Resultado"]
#sub_titles4 = ["Dowload","Matrícula <= 50",matricula1_velocidad_bajada,matricula1_posicion_prueba_p5_subida_6]
sub_titles5 = ["Sentido","Matricula","Velocidad Por Matricula","Posición de la prueba","Prueba P5%","Resultado","Posición de la prueba","Prueba P95%","Resultado"]
#sub_titles6 = ["Upload","Matrícula <= 50",matricula1_velocidad_subida,matricula1_posicion_prueba_p5_subida_6]


for sheet_name in workbook.sheetnames[1:]:
    sheet = workbook[sheet_name]
    # Colocar el título principal en la fila 1, combinado a través de un rango adecuado
    sheet.merge_cells('A1:V1')  # Ajusta el rango de celdas según el número de columnas en tus datos
    sheet['A1'] = main_title
    sheet['A1'].alignment = Alignment(horizontal="center", vertical="center")
    sheet['A1'].fill=gray_fill
    sheet['A1'].font=white_font

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    for col, sub_title1 in zip(['B'], sub_titles):
        sheet[f'{col}3'] = sub_title1
        sheet[f'{col}3'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}3'].fill=gray_fill
        sheet[f'{col}3'].font=white_font
        #ws.row_dimensions[3].height = 60

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    for col, sub_title2 in zip(['B'], sub_titles2):
        sheet[f'{col}4'] = sub_title2
        sheet[f'{col}4'].alignment = Alignment(horizontal="center", vertical="center")


    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    for col, sub_title3 in zip(columns, sub_titles3):
        sheet[f'{col}6'] = sub_title3
        sheet[f'{col}6'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}6'].fill=gray_fill
        sheet[f'{col}6'].font=white_font
    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)





    columns=['B','C','D','E','F','G','H','I','J']
    matricula1_posicion_prueba_p5_bajada = globals()[f"matricula1_posicion_prueba_p5_bajada_{sheet_name}"]
    matricula1_prueba_p5_bajada = globals()[f"matricula1_prueba_p5_bajada_{sheet_name}"]
    matricula1_resultado_p5_bajada = globals()[f"matricula1_resultado_p5_bajada_{sheet_name}"]

    matricula1_posicion_prueba_p95_bajada = globals()[f"matricula1_posicion_prueba_p95_bajada_{sheet_name}"]
    matricula1_prueba_p95_bajada = globals()[f"matricula1_prueba_p95_bajada_{sheet_name}"]
    matricula1_resultado_p95_bajada = globals()[f"matricula1_resultado_p95_bajada_{sheet_name}"]


    sub_titles4 = ["Dowload","PERFIL_1",matricula1_velocidad_bajada_mostrar,matricula1_posicion_prueba_p5_bajada,matricula1_prueba_p5_bajada,matricula1_resultado_p5_bajada,matricula1_posicion_prueba_p95_bajada,matricula1_prueba_p95_bajada,matricula1_resultado_p95_bajada]
    for col, sub_title4 in zip(columns, sub_titles4):
        sheet[f'{col}7'] = sub_title4
        sheet[f'{col}7'].alignment = Alignment(horizontal="center", vertical="center")

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    for col, sub_title5 in zip(columns, sub_titles5):
        sheet[f'{col}9'] = sub_title5
        sheet[f'{col}9'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}9'].fill=gray_fill
        sheet[f'{col}9'].font=white_font

    columns=['B','C','D','E','F','G','H','I','J']
    matricula1_posicion_prueba_p5_subida = globals()[f"matricula1_posicion_prueba_p5_subida_{sheet_name}"]
    matricula1_prueba_p5_subida = globals()[f"matricula1_prueba_p5_subida_{sheet_name}"]
    matricula1_resultado_p5_subida = globals()[f"matricula1_resultado_p5_subida_{sheet_name}"]

    matricula1_posicion_prueba_p95_subida = globals()[f"matricula1_posicion_prueba_p95_subida_{sheet_name}"]
    matricula1_prueba_p95_subida = globals()[f"matricula1_prueba_p95_subida_{sheet_name}"]
    matricula1_resultado_p95_subida = globals()[f"matricula1_resultado_p95_subida_{sheet_name}"]

    sub_titles6 = ["Upload","PERFIL_1",matricula1_velocidad_subida_mostrar,matricula1_posicion_prueba_p5_subida,matricula1_prueba_p5_subida,matricula1_resultado_p5_subida,matricula1_posicion_prueba_p95_subida,matricula1_prueba_p95_subida,matricula1_resultado_p95_subida]
    for col, sub_title6 in zip(columns, sub_titles6):
        sheet[f'{col}10'] = sub_title6
        sheet[f'{col}10'].alignment = Alignment(horizontal="center", vertical="center")

    for row in [1,3, 6, 9, 14]:  # Ajusta los números de fila según sea necesario
        sheet.row_dimensions[row].height = 60

    columns=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V']

    for col in columns:
        sheet[f'{col}14'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}14'].fill=gray_fill
        sheet[f'{col}14'].font=white_font



# Guardar el archivo actualizado
workbook.save(output_path)

# Descargar el archivo a tu máquina local
from google.colab import files
files.download(output_path)

# descargar y crear archiv ecxel por pestañas
#matricula2
output_path = '/content/20260617_PERFIL_2.xlsx'
# Crear un archivo Excel con múltiples hojas
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    matricula2.to_excel(writer, sheet_name='PERFIL_2', index=False, startrow=13, startcol=1)
    matricula2_rango6.to_excel(writer, sheet_name='6', index=False, startrow=13, startcol=1 )
    matricula2_rango7.to_excel(writer, sheet_name='7', index=False, startrow=13, startcol=1)
    matricula2_rango8.to_excel(writer, sheet_name='8', index=False, startrow=13, startcol=1)
    matricula2_rango9.to_excel(writer, sheet_name='9', index=False, startrow=13, startcol=1)
    matricula2_rango10.to_excel(writer, sheet_name='10', index=False, startrow=13, startcol=1)
    matricula2_rango11.to_excel(writer, sheet_name='11', index=False, startrow=13, startcol=1)
    matricula2_rango12.to_excel(writer, sheet_name='12', index=False, startrow=13, startcol=1)
    matricula2_rango13.to_excel(writer, sheet_name='13', index=False, startrow=13, startcol=1)
    matricula2_rango14.to_excel(writer, sheet_name='14', index=False, startrow=13, startcol=1)
    matricula2_rango15.to_excel(writer, sheet_name='15', index=False, startrow=13, startcol=1)
    matricula2_rango16.to_excel(writer, sheet_name='16', index=False, startrow=13, startcol=1)
    matricula2_rango17.to_excel(writer, sheet_name='17', index=False, startrow=13, startcol=1)
    matricula2_rango18.to_excel(writer, sheet_name='18', index=False, startrow=13, startcol=1)
    matricula2_rango19.to_excel(writer, sheet_name='19', index=False, startrow=13, startcol=1)
    matricula2_rango20.to_excel(writer, sheet_name='20', index=False, startrow=13, startcol=1)
 # Acceder al objeto de libro de trabajo para personalizar
    workbook = writer.book


#print(f"Archivo guardado en: {output_path}")
#wb = load_workbook(output_path)

# Definir el título principal que se quiere poner en todas las hojas
main_title = "2.3.4.4.2 VELOCIDAD EFECTIVA MÍNIMA DE TRANSMISIÓN DE DATOS"
# Definir los subtítulos (para la fila 2 y siguientes)
sub_titles = ["Año-Mes(AAAA-MM)"]
sub_titles2 = ["2026-06"]
sub_titles3 = ["Sentido","Matricula","Velocidad Por Matricula","Posición de la prueba","Prueba P5%","Resultado","Posición de la prueba","Prueba P95%","Resultado"]
#sub_titles4 = ["Dowload","Matrícula <= 50",matricula2_velocidad_bajada,matricula2_posicion_prueba_p5_subida_6]
sub_titles5 = ["Sentido","Matricula","Velocidad Por Matricula","Posición de la prueba","Prueba P5%","Resultado","Posición de la prueba","Prueba P95%","Resultado"]
#sub_titles6 = ["Upload","Matrícula <= 50",matricula2_velocidad_subida,matricula2_posicion_prueba_p5_subida_6]


for sheet_name in workbook.sheetnames[1:]:
    sheet = workbook[sheet_name]
    # Colocar el título principal en la fila 1, combinado a través de un rango adecuado
    sheet.merge_cells('A1:V1')  # Ajusta el rango de celdas según el número de columnas en tus datos
    sheet['A1'] = main_title
    sheet['A1'].alignment = Alignment(horizontal="center", vertical="center")
    sheet['A1'].fill=gray_fill
    sheet['A1'].font=white_font

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    for col, sub_title1 in zip(['B'], sub_titles):
        sheet[f'{col}3'] = sub_title1
        sheet[f'{col}3'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}3'].fill=gray_fill
        sheet[f'{col}3'].font=white_font

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    for col, sub_title2 in zip(['B'], sub_titles2):
        sheet[f'{col}4'] = sub_title2
        sheet[f'{col}4'].alignment = Alignment(horizontal="center", vertical="center")

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    for col, sub_title3 in zip(columns, sub_titles3):
        sheet[f'{col}6'] = sub_title3
        sheet[f'{col}6'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}6'].fill=gray_fill
        sheet[f'{col}6'].font=white_font
    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)

    columns=['B','C','D','E','F','G','H','I','J']
    matricula2_posicion_prueba_p5_bajada = globals()[f"matricula2_posicion_prueba_p5_bajada_{sheet_name}"]
    matricula2_prueba_p5_bajada = globals()[f"matricula2_prueba_p5_bajada_{sheet_name}"]
    matricula2_resultado_p5_bajada = globals()[f"matricula2_resultado_p5_bajada_{sheet_name}"]

    matricula2_posicion_prueba_p95_bajada = globals()[f"matricula2_posicion_prueba_p95_bajada_{sheet_name}"]
    matricula2_prueba_p95_bajada = globals()[f"matricula2_prueba_p95_bajada_{sheet_name}"]
    matricula2_resultado_p95_bajada = globals()[f"matricula2_resultado_p95_bajada_{sheet_name}"]


    sub_titles4 = ["Dowload","PERFIL_2",matricula2_velocidad_bajada_mostrar,matricula2_posicion_prueba_p5_bajada,matricula2_prueba_p5_bajada,matricula2_resultado_p5_bajada,matricula2_posicion_prueba_p95_bajada,matricula2_prueba_p95_bajada,matricula2_resultado_p95_bajada]
    for col, sub_title4 in zip(columns, sub_titles4):
        sheet[f'{col}7'] = sub_title4
        sheet[f'{col}7'].alignment = Alignment(horizontal="center", vertical="center")

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    for col, sub_title5 in zip(columns, sub_titles5):
        sheet[f'{col}9'] = sub_title5
        sheet[f'{col}9'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}9'].fill=gray_fill
        sheet[f'{col}9'].font=white_font




    columns=['B','C','D','E','F','G','H','I','J']
    matricula2_posicion_prueba_p5_subida = globals()[f"matricula2_posicion_prueba_p5_subida_{sheet_name}"]
    matricula2_prueba_p5_subida = globals()[f"matricula2_prueba_p5_subida_{sheet_name}"]
    matricula2_resultado_p5_subida = globals()[f"matricula2_resultado_p5_subida_{sheet_name}"]

    matricula2_posicion_prueba_p95_subida = globals()[f"matricula2_posicion_prueba_p95_subida_{sheet_name}"]
    matricula2_prueba_p95_subida = globals()[f"matricula2_prueba_p95_subida_{sheet_name}"]
    matricula2_resultado_p95_subida = globals()[f"matricula2_resultado_p95_subida_{sheet_name}"]

    sub_titles6 = ["Upload","PERFIL_2",matricula2_velocidad_subida_mostrar,matricula2_posicion_prueba_p5_subida,matricula2_prueba_p5_subida,matricula2_resultado_p5_subida,matricula2_posicion_prueba_p95_subida,matricula2_prueba_p95_subida,matricula2_resultado_p95_subida]
    for col, sub_title6 in zip(columns, sub_titles6):
        sheet[f'{col}10'] = sub_title6
        sheet[f'{col}10'].alignment = Alignment(horizontal="center", vertical="center")

    for row in [1,3, 6, 9, 14]:  # Ajusta los números de fila según sea necesario
        sheet.row_dimensions[row].height = 60

    columns=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V']

    for col in columns:
        sheet[f'{col}14'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}14'].fill=gray_fill
        sheet[f'{col}14'].font=white_font



# Guardar el archivo actualizado
workbook.save(output_path)

# Descargar el archivo a tu máquina local
from google.colab import files
files.download(output_path)

# descargar y crear archiv ecxel por pestañas
#matricula3
output_path = '/content/20260617_PERFIL_3+.xlsx'
# Crear un archivo Excel con múltiples hojas
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    matricula3.to_excel(writer, sheet_name='PERFIL_3+', index=False, startrow=13, startcol=1)
    matricula3_rango6.to_excel(writer, sheet_name='6', index=False, startrow=13, startcol=1)
    matricula3_rango7.to_excel(writer, sheet_name='7', index=False, startrow=13, startcol=1)
    matricula3_rango8.to_excel(writer, sheet_name='8', index=False, startrow=13, startcol=1)
    matricula3_rango9.to_excel(writer, sheet_name='9', index=False, startrow=13, startcol=1)
    matricula3_rango10.to_excel(writer, sheet_name='10', index=False, startrow=13, startcol=1)
    matricula3_rango11.to_excel(writer, sheet_name='11', index=False, startrow=13, startcol=1)
    matricula3_rango12.to_excel(writer, sheet_name='12', index=False, startrow=13, startcol=1)
    matricula3_rango13.to_excel(writer, sheet_name='13', index=False, startrow=13, startcol=1)
    matricula3_rango14.to_excel(writer, sheet_name='14', index=False, startrow=13, startcol=1)
    matricula3_rango15.to_excel(writer, sheet_name='15', index=False, startrow=13, startcol=1)
    matricula3_rango16.to_excel(writer, sheet_name='16', index=False, startrow=13, startcol=1)
    matricula3_rango17.to_excel(writer, sheet_name='17', index=False, startrow=13, startcol=1)
    matricula3_rango18.to_excel(writer, sheet_name='18', index=False, startrow=13, startcol=1)
    matricula3_rango19.to_excel(writer, sheet_name='19', index=False, startrow=13, startcol=1)
    matricula3_rango20.to_excel(writer, sheet_name='20', index=False, startrow=13, startcol=1)
 # Acceder al objeto de libro de trabajo para personalizar
    workbook = writer.book


#print(f"Archivo guardado en: {output_path}")
#wb = load_workbook(output_path)

# Definir el título principal que se quiere poner en todas las hojas
main_title = "2.3.4.4.2 VELOCIDAD EFECTIVA MÍNIMA DE TRANSMISIÓN DE DATOS"
# Definir los subtítulos (para la fila 2 y siguientes)
sub_titles = ["Año-Mes(AAAA-MM)"]
sub_titles2 = ["2026-06"]
sub_titles3 = ["Sentido","Matricula","Velocidad Por Matricula","Posición de la prueba","Prueba P5%","Resultado","Posición de la prueba","Prueba P95%","Resultado"]
#sub_titles4 = ["Dowload","Matrícula <= 50",matricula3_velocidad_bajada,matricula3_posicion_prueba_p5_subida_6]
sub_titles5 = ["Sentido","Matricula","Velocidad Por Matricula","Posición de la prueba","Prueba P5%","Resultado","Posición de la prueba","Prueba P95%","Resultado"]
#sub_titles6 = ["Upload","Matrícula <= 50",matricula3_velocidad_subida,matricula3_posicion_prueba_p5_subida_6]


for sheet_name in workbook.sheetnames[1:]:
    sheet = workbook[sheet_name]
    # Colocar el título principal en la fila 1, combinado a través de un rango adecuado
    sheet.merge_cells('A1:V1')  # Ajusta el rango de celdas según el número de columnas en tus datos
    sheet['A1'] = main_title
    sheet['A1'].alignment = Alignment(horizontal="center", vertical="center")
    sheet['A1'].fill=gray_fill
    sheet['A1'].font=white_font

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    for col, sub_title1 in zip(['B'], sub_titles):
        sheet[f'{col}3'] = sub_title1
        sheet[f'{col}3'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}3'].fill=gray_fill
        sheet[f'{col}3'].font=white_font
    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    for col, sub_title2 in zip(['B'], sub_titles2):
        sheet[f'{col}4'] = sub_title2
        sheet[f'{col}4'].alignment = Alignment(horizontal="center", vertical="center")

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    for col, sub_title3 in zip(columns, sub_titles3):
        sheet[f'{col}6'] = sub_title3
        sheet[f'{col}6'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}6'].fill=gray_fill
        sheet[f'{col}6'].font=white_font
    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)

    columns=['B','C','D','E','F','G','H','I','J']
    matricula3_posicion_prueba_p5_bajada = globals()[f"matricula3_posicion_prueba_p5_bajada_{sheet_name}"]
    matricula3_prueba_p5_bajada = globals()[f"matricula3_prueba_p5_bajada_{sheet_name}"]
    matricula3_resultado_p5_bajada = globals()[f"matricula3_resultado_p5_bajada_{sheet_name}"]

    matricula3_posicion_prueba_p95_bajada = globals()[f"matricula3_posicion_prueba_p95_bajada_{sheet_name}"]
    matricula3_prueba_p95_bajada = globals()[f"matricula3_prueba_p95_bajada_{sheet_name}"]
    matricula3_resultado_p95_bajada = globals()[f"matricula3_resultado_p95_bajada_{sheet_name}"]


    sub_titles4 = ["Dowload","PERFIL_3+",matricula3_velocidad_bajada_mostrar,matricula3_posicion_prueba_p5_bajada,matricula3_prueba_p5_bajada,matricula3_resultado_p5_bajada,matricula3_posicion_prueba_p95_bajada,matricula3_prueba_p95_bajada,matricula3_resultado_p95_bajada]
    for col, sub_title4 in zip(columns, sub_titles4):
        sheet[f'{col}7'] = sub_title4
        sheet[f'{col}7'].alignment = Alignment(horizontal="center", vertical="center")

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    for col, sub_title5 in zip(columns, sub_titles5):
        sheet[f'{col}9'] = sub_title5
        sheet[f'{col}9'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}9'].fill=gray_fill
        sheet[f'{col}9'].font=white_font




    columns=['B','C','D','E','F','G','H','I','J']
    matricula3_posicion_prueba_p5_subida = globals()[f"matricula3_posicion_prueba_p5_subida_{sheet_name}"]
    matricula3_prueba_p5_subida = globals()[f"matricula3_prueba_p5_subida_{sheet_name}"]
    matricula3_resultado_p5_subida = globals()[f"matricula3_resultado_p5_subida_{sheet_name}"]
    matricula3_posicion_prueba_p95_subida = globals()[f"matricula3_posicion_prueba_p95_subida_{sheet_name}"]
    matricula3_prueba_p95_subida = globals()[f"matricula3_prueba_p95_subida_{sheet_name}"]
    matricula3_resultado_p95_subida = globals()[f"matricula3_resultado_p95_subida_{sheet_name}"]

    sub_titles6 = ["Upload","PERFIL_3+",matricula3_velocidad_subida_mostrar,matricula3_posicion_prueba_p5_subida,matricula3_prueba_p5_subida,matricula3_resultado_p5_subida,matricula3_posicion_prueba_p95_subida,matricula3_prueba_p95_subida,matricula3_resultado_p95_subida]
    for col, sub_title6 in zip(columns, sub_titles6):
        sheet[f'{col}10'] = sub_title6
        sheet[f'{col}10'].alignment = Alignment(horizontal="center", vertical="center")

    for row in [1,3, 6, 9, 14]:  # Ajusta los números de fila según sea necesario
        sheet.row_dimensions[row].height = 60

    columns=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V']

    for col in columns:
        sheet[f'{col}14'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}14'].fill=gray_fill
        sheet[f'{col}14'].font=white_font


# Guardar el archivo actualizado
workbook.save(output_path)

# Descargar el archivo a tu máquina local
from google.colab import files
files.download(output_path)

# descargar y crear archiv ecxel por pestañas
#matricula4
output_path = '/content/20260617_PERFIL_4+.xlsx'
# Crear un archivo Excel con múltiples hojas
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    matricula4.to_excel(writer, sheet_name='PERFIL_4+', index=False, startrow=13, startcol=1)
    matricula4_rango6.to_excel(writer, sheet_name='6', index=False, startrow=13, startcol=1)
    matricula4_rango7.to_excel(writer, sheet_name='7', index=False, startrow=13, startcol=1)
    matricula4_rango8.to_excel(writer, sheet_name='8', index=False, startrow=13, startcol=1)
    matricula4_rango9.to_excel(writer, sheet_name='9', index=False, startrow=13, startcol=1)
    matricula4_rango10.to_excel(writer, sheet_name='10', index=False, startrow=13, startcol=1)
    matricula4_rango11.to_excel(writer, sheet_name='11', index=False, startrow=13, startcol=1)
    matricula4_rango12.to_excel(writer, sheet_name='12', index=False, startrow=13, startcol=1)
    matricula4_rango13.to_excel(writer, sheet_name='13', index=False, startrow=13, startcol=1)
    matricula4_rango14.to_excel(writer, sheet_name='14', index=False, startrow=13, startcol=1)
    matricula4_rango15.to_excel(writer, sheet_name='15', index=False, startrow=13, startcol=1)
    matricula4_rango16.to_excel(writer, sheet_name='16', index=False, startrow=13, startcol=1)
    matricula4_rango17.to_excel(writer, sheet_name='17', index=False, startrow=13, startcol=1)
    matricula4_rango18.to_excel(writer, sheet_name='18', index=False, startrow=13, startcol=1)
    matricula4_rango19.to_excel(writer, sheet_name='19', index=False, startrow=13, startcol=1)
    matricula4_rango20.to_excel(writer, sheet_name='20', index=False, startrow=13, startcol=1)
 # Acceder al objeto de libro de trabajo para personalizar
    workbook = writer.book


#print(f"Archivo guardado en: {output_path}")
#wb = load_workbook(output_path)

# Definir el título principal que se quiere poner en todas las hojas
main_title = "2.3.4.4.2 VELOCIDAD EFECTIVA MÍNIMA DE TRANSMISIÓN DE DATOS"
# Definir los subtítulos (para la fila 2 y siguientes)
sub_titles = ["Año-Mes(AAAA-MM)"]
sub_titles2 = ["2026-06"]
sub_titles3 = ["Sentido","Matricula","Velocidad Por Matricula","Posición de la prueba","Prueba P5%","Resultado","Posición de la prueba","Prueba P95%","Resultado"]
#sub_titles4 = ["Dowload","Matrícula <= 50",matricula4_velocidad_bajada,matricula4_posicion_prueba_p5_subida_6]
sub_titles5 = ["Sentido","Matricula","Velocidad Por Matricula","Posición de la prueba","Prueba P5%","Resultado","Posición de la prueba","Prueba P95%","Resultado"]
#sub_titles6 = ["Upload","Matrícula <= 50",matricula4_velocidad_subida,matricula4_posicion_prueba_p5_subida_6]


for sheet_name in workbook.sheetnames[1:]:
    sheet = workbook[sheet_name]
    # Colocar el título principal en la fila 1, combinado a través de un rango adecuado
    sheet.merge_cells('A1:V1')  # Ajusta el rango de celdas según el número de columnas en tus datos
    sheet['A1'] = main_title
    sheet['A1'].alignment = Alignment(horizontal="center", vertical="center")
    sheet['A1'].fill=gray_fill
    sheet['A1'].font=white_font

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    for col, sub_title1 in zip(['B'], sub_titles):
        sheet[f'{col}3'] = sub_title1
        sheet[f'{col}3'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}3'].fill=gray_fill
        sheet[f'{col}3'].font=white_font

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    for col, sub_title2 in zip(['B'], sub_titles2):
        sheet[f'{col}4'] = sub_title2
        sheet[f'{col}4'].alignment = Alignment(horizontal="center", vertical="center")

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    for col, sub_title3 in zip(columns, sub_titles3):
        sheet[f'{col}6'] = sub_title3
        sheet[f'{col}6'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}6'].fill=gray_fill
        sheet[f'{col}6'].font=white_font
    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)

    columns=['B','C','D','E','F','G','H','I','J']
    matricula4_posicion_prueba_p5_bajada = globals()[f"matricula4_posicion_prueba_p5_bajada_{sheet_name}"]
    matricula4_prueba_p5_bajada = globals()[f"matricula4_prueba_p5_bajada_{sheet_name}"]
    matricula4_resultado_p5_bajada = globals()[f"matricula4_resultado_p5_bajada_{sheet_name}"]

    matricula4_posicion_prueba_p95_bajada = globals()[f"matricula4_posicion_prueba_p95_bajada_{sheet_name}"]
    matricula4_prueba_p95_bajada = globals()[f"matricula4_prueba_p95_bajada_{sheet_name}"]
    matricula4_resultado_p95_bajada = globals()[f"matricula4_resultado_p95_bajada_{sheet_name}"]


    sub_titles4 = ["Dowload","PERFIL_4+",matricula4_velocidad_bajada_mostrar,matricula4_posicion_prueba_p5_bajada,matricula4_prueba_p5_bajada,matricula4_resultado_p5_bajada,matricula4_posicion_prueba_p95_bajada,matricula4_prueba_p95_bajada,matricula4_resultado_p95_bajada]
    for col, sub_title4 in zip(columns, sub_titles4):
        sheet[f'{col}7'] = sub_title4
        sheet[f'{col}7'].alignment = Alignment(horizontal="center", vertical="center")

    # Colocar los subtítulos en la fila 2 (columnas A, B, C, etc.)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    for col, sub_title5 in zip(columns, sub_titles5):
        sheet[f'{col}9'] = sub_title5
        sheet[f'{col}9'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}9'].fill=gray_fill
        sheet[f'{col}9'].font=white_font




    columns=['B','C','D','E','F','G','H','I','J']
    matricula4_posicion_prueba_p5_subida = globals()[f"matricula4_posicion_prueba_p5_subida_{sheet_name}"]
    matricula4_prueba_p5_subida = globals()[f"matricula4_prueba_p5_subida_{sheet_name}"]
    matricula4_resultado_p5_subida = globals()[f"matricula4_resultado_p5_subida_{sheet_name}"]

    matricula4_posicion_prueba_p95_subida = globals()[f"matricula4_posicion_prueba_p95_subida_{sheet_name}"]
    matricula4_prueba_p95_subida = globals()[f"matricula4_prueba_p95_subida_{sheet_name}"]
    matricula4_resultado_p95_subida = globals()[f"matricula4_resultado_p95_subida_{sheet_name}"]

    sub_titles6 = ["Upload","PERFIL_4+",matricula4_velocidad_subida_mostrar,matricula4_posicion_prueba_p5_subida,matricula4_prueba_p5_subida,matricula4_resultado_p5_subida,matricula4_posicion_prueba_p95_subida,matricula4_prueba_p95_subida,matricula4_resultado_p95_subida]
    for col, sub_title6 in zip(columns, sub_titles6):
        sheet[f'{col}10'] = sub_title6
        sheet[f'{col}10'].alignment = Alignment(horizontal="center", vertical="center")

    for row in [1,3, 6, 9, 14]:  # Ajusta los números de fila según sea necesario
        sheet.row_dimensions[row].height = 60

    columns=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V']

    for col in columns:
        sheet[f'{col}14'].alignment = Alignment(horizontal="center", vertical="center")
        sheet[f'{col}14'].fill=gray_fill
        sheet[f'{col}14'].font=white_font





# Guardar el archivo actualizado
workbook.save(output_path)

# Descargar el archivo a tu máquina local
from google.colab import files
files.download(output_path)
