# -*- coding: latin-1 -*-
import pandas as pd
import numpy as np

def ejecutar_cruce_seguro(CONTADORES, VELOCIDAD,SITIOS,PQRSD):
    print("Iniciando procesamiento seguro desde repositorio privado...")
    CONTADORES_TRATADA=CONTADORES.copy()
    VELOCIDAD_TRATADA=VELOCIDAD.copy()
    SITIOS_TRATADA=SITIOS.copy()
    PQRSD_TRATADA=PQRSD.copy()

    VELOCIDAD_TRATADA['Fecha de ejecucion'] = pd.to_datetime(VELOCIDAD_TRATADA['Fecha de ejecucion'])

    #print(VELOCIDAD_TRATADA[VELOCIDAD_TRATADA['Identificador beneficiario'] == '23135'])

    VELOCIDAD_TRATADA = VELOCIDAD_TRATADA.groupby('Identificador beneficiario')['Fecha de ejecucion'].max().reset_index()

    # Crea una nueva columna con la cantidad de días de diferencia
    VELOCIDAD_TRATADA['Dias_diferencia'] = (
     (pd.Timestamp('today') - pd.Timedelta(days=1)).normalize() - VELOCIDAD_TRATADA['Fecha de ejecucion'].dt.normalize()
    ).dt.days
    CONTADORES_TRATADA['DISPOSITIVO'] = CONTADORES_TRATADA['DISPOSITIVO'].str.split('_').str[0]


    CONTADORES_TRATADA = CONTADORES_TRATADA.groupby('DISPOSITIVO')['DIAS SIN MEDICION'].max().reset_index()

    CONTADORES_TRATADA['DISPOSITIVO'] = CONTADORES_TRATADA['DISPOSITIVO'].astype('str')
    VELOCIDAD_TRATADA['Identificador beneficiario'] = VELOCIDAD_TRATADA['Identificador beneficiario'].astype('str')

    SITIOS_TRATADA ['Id_Beneficiario'] = SITIOS_TRATADA['Id_Beneficiario'].astype('str')

    PQRSD_TRATADA['Fecha_Creacion'] = pd.to_datetime(PQRSD_TRATADA['Fecha_Creacion'])
    PQRSD_TRATADA['ID_Beneficiario'] = PQRSD_TRATADA['ID_Beneficiario'].astype('str')

    PQRSD_TRATADA = PQRSD_TRATADA[PQRSD_TRATADA['Estado']!='Cerrar']
    PQRSD_TRATADA_VELOCIDAD = PQRSD_TRATADA[PQRSD_TRATADA['SUBCATEGORIA'].isin([
        'CD-MEDICIÓN DIRECTA DE VELOCIDAD EFECTIVA DE TRANSMISIÓN DE DATOS',
        'CD-FALLA EN EJECUCIÓN DE PRUEBA DE VELOCIDAD 5 DÍAS CALENDARIO'
     ])]


    cruce = pd.merge(
     left=CONTADORES_TRATADA,
     right=SITIOS_TRATADA,
     left_on='DISPOSITIVO',
     right_on='Id_Beneficiario',
     how='left'         # Especifica que es un INNER JOIN
    )



    cruce = pd.merge(
        cruce,
        VELOCIDAD_TRATADA,
        left_on='Id_Beneficiario',
        right_on='Identificador beneficiario',
        how='left')


    print(cruce[cruce['DISPOSITIVO'] == '23135'])
    print(cruce[cruce['Identificador beneficiario'] == '23135'])
    #print(cruce[cruce['DISPOSITIVO'] == '22707'])




    # 1. Convertimos la columna a texto
    PQRSD_TRATADA_VELOCIDAD['ID_Beneficiario'] = PQRSD_TRATADA_VELOCIDAD['ID_Beneficiario'].astype(str)
    # 2. Reemplazamos la palabra "nan" (y sus variantes) por verdaderos nulos
    #PQRSD_TRATADA_VELOCIDAD['ID_Beneficiario'] = PQRSD_TRATADA_VELOCIDAD['ID_Beneficiario'].replace(['nan', '<NA>', 'None', 'nan.0'], np.nan)

    # 1. Convertimos la columna a texto
    cruce['Identificador beneficiario'] = cruce['Identificador beneficiario'].astype(str)
    # 2. Reemplazamos la palabra "nan" (y sus variantes) por verdaderos nulos
    cruce['Identificador beneficiario'] = cruce['Identificador beneficiario'].replace(['nan', '<NA>', 'None', 'nan.0'], np.nan)


    cruce = pd.merge(
        left=cruce,
        right=PQRSD_TRATADA_VELOCIDAD,
        left_on='DISPOSITIVO',
        right_on='ID_Beneficiario',
        how='outer'         # Especifica que es un INNER JOIN
    )


    #print(cruce[cruce['DISPOSITIVO'] == "22707"])

    cruce ['VERIFICACION'] = cruce['Dias_diferencia']==cruce['DIAS SIN MEDICION']
    
    #cruce.to_csv('cruce.csv', index=False)
    return cruce


