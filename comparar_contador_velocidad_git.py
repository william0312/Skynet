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

#    print(VELOCIDAD_TRATADA[VELOCIDAD_TRATADA['Identificador beneficiario'] == '12004'])

    VELOCIDAD_TRATADA = VELOCIDAD_TRATADA.groupby('Identificador beneficiario')['Fecha de ejecucion'].max().reset_index()
#    print(VELOCIDAD_TRATADA[VELOCIDAD_TRATADA['Identificador beneficiario'] == '12004'])
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
#    print (PQRSD_TRATADA)

    PQRSD_TRATADA_VELOCIDAD = PQRSD_TRATADA[PQRSD_TRATADA['SUBCATEGORIA'].isin([
        'CD-MEDICION DIRECTA DE VELOCIDAD EFECTIVA DE TRANSMISION DE DATOS',
        'CD-FALLA EN EJECUCION DE PRUEBA DE VELOCIDAD 5 DIAS CALENDARIO'
     ])]



#    print (PQRSD_TRATADA_VELOCIDAD)

    cruce = pd.merge(
     left=CONTADORES_TRATADA,
     right=SITIOS_TRATADA,
     left_on='DISPOSITIVO',
     right_on='Id_Beneficiario',
     how='right'         # Especifica que es un INNER JOIN
    )

    cruce = pd.merge(
        cruce,
        VELOCIDAD_TRATADA,
        left_on='Id_Beneficiario',
        right_on='Identificador beneficiario',
        how='left')


    print(cruce.columns)

#   print(cruce[cruce['DISPOSITIVO'] == '23135'])
#   print(cruce[cruce['Identificador beneficiario'] == '23135'])
    #print(cruce[cruce['DISPOSITIVO'] == '22707'])




    # 1. Convertimos la columna a texto
    PQRSD_TRATADA_VELOCIDAD['ID_Beneficiario'] = PQRSD_TRATADA_VELOCIDAD['ID_Beneficiario'].astype(str)
    # 1. Convertimos la columna a texto
    cruce['DISPOSITIVO'] = cruce['DISPOSITIVO'].astype(str)
    # 2. Reemplazamos la palabra "nan" (y sus variantes) por verdaderos nulos
    cruce['DISPOSITIVO'] = cruce['DISPOSITIVO'].replace(['nan', '<NA>', 'None', 'nan.0'], np.nan)


    cruce = pd.merge(
        left=cruce,
        right=PQRSD_TRATADA_VELOCIDAD,
        left_on='DISPOSITIVO',
        right_on='ID_Beneficiario',
        how='outer'         # Especifica que es un INNER JOIN
    )


    #print(cruce[cruce['DISPOSITIVO'] == "22707"])

    cruce ['VERIFICACION'] = cruce['Dias_diferencia']==cruce['DIAS SIN MEDICION']
    cruce['CANTIDAD_PQRSD'] = cruce.groupby('DISPOSITIVO')['ID'].transform('count')


    # 1. Definimos las reglas de lo que está "Bien"
    condiciones = [
    # Regla 1: Menos de 5 días Y exactamente 0 registros
    (cruce['DIAS SIN MEDICION'] < 5) & (cruce['CANTIDAD_PQRSD'] == 0),
    
    # Regla 2: Entre 5 y 8 días Y exactamente 1 registro
    (cruce['DIAS SIN MEDICION'] >= 5) & (cruce['DIAS SIN MEDICION'] < 8) & (cruce['CANTIDAD_PQRSD'] == 1),
    
    # Regla 3: Mayor a 8 días (Agrego esto por si acaso, asumiendo que debe tener 2 o más registros)
    (cruce['DIAS SIN MEDICION'] >= 8) & (cruce['CANTIDAD_PQRSD'] >= 2)
    ]

    # 2. Si se cumple alguna de las reglas de arriba, el resultado es 'Bien'
    resultados = ['Bien', 'Bien', 'Bien']

    # 3. Aplicamos la evaluación. Todo lo que no encaje en las reglas, por defecto será 'Mal'
    cruce['VERIFICACION_PQRSD'] = np.select(condiciones, resultados, default='Mal')
    
    #cruce.to_csv('cruce.csv', index=False)
    return cruce


