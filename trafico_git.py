# -*- coding: latin-1 -*-
import pandas as pd
import numpy as np

def ejecutar_cruce_seguro(CONTADORES,TRAFICO,SITIOS,PQRSD):
    print("Iniciando procesamiento seguro desde repositorio privado...")
    CONTADORES_TRATADA=CONTADORES.copy()
    TRAFICO_TRATADA=TRAFICO.copy()
    SITIOS_TRATADA=SITIOS.copy()
    PQRSD_TRATADA=PQRSD.copy()

    #tratamiento contadores

    CONTADORES_TRATADA_INDOOR = CONTADORES_TRATADA[CONTADORES_TRATADA['TIPO DE ALERTA'] == 'indoor']
    CONTADORES_TRATADA_OUTDOOR = CONTADORES_TRATADA[CONTADORES_TRATADA['TIPO DE ALERTA'] == 'outdoor']

    CONTADORES_TRATADA_TRAFICO = CONTADORES_TRATADA[CONTADORES_TRATADA['TIPO DE ALERTA'] == 'traffic']
    CONTADORES_TRATADA_TRAFICO['DISPOSITIVO_ZONA'] = CONTADORES_TRATADA_TRAFICO['DISPOSITIVO'].str.split('_').str[1]

    filas_indoor_trafico = CONTADORES_TRATADA_TRAFICO[CONTADORES_TRATADA_TRAFICO['DISPOSITIVO_ZONA'] == 'Indoor'].copy()
    filas_outdoor_trafico = CONTADORES_TRATADA_TRAFICO[CONTADORES_TRATADA_TRAFICO['DISPOSITIVO_ZONA'] == 'Outdoor'].copy()

    # 4. Eliminar la columna 'DISPOSITIVO_ZONA' para igualar la estructura
    filas_indoor_trafico = filas_indoor_trafico.drop(columns=['DISPOSITIVO_ZONA'])
    filas_outdoor_trafico = filas_outdoor_trafico.drop(columns=['DISPOSITIVO_ZONA'])

    # 5. Unir (concatenar) ambos DataFrames
    CONTADORES_TRATADA_INDOOR = pd.concat([CONTADORES_TRATADA_INDOOR, filas_indoor_trafico], ignore_index=True)
    CONTADORES_TRATADA_OUTDOOR = pd.concat([CONTADORES_TRATADA_OUTDOOR, filas_outdoor_trafico], ignore_index=True)

    PQRSD_TRATADA['Fecha_Creacion'] = pd.to_datetime(PQRSD_TRATADA['Fecha_Creacion'])
    PQRSD_TRATADA['ID_Beneficiario'] = PQRSD_TRATADA['ID_Beneficiario'].astype('int64')

    SITIOS_TRATADA['Estado_Sitio'] = SITIOS_TRATADA['Estado_Sitio'].replace(';', '_')

    #print(CONTADORES_TRATADA_OUTDOOR[CONTADORES_TRATADA_OUTDOOR['IDENTIFICADOR BENEFICIARIO'] == 23039])

    #tratamiento ping


    TRAFICO_TRATADA_INDOOR=TRAFICO_TRATADA[TRAFICO_TRATADA['TIPO DE ZONA']=='indoor']
    TRAFICO_TRATADA_OUTDOOR=TRAFICO_TRATADA[TRAFICO_TRATADA['TIPO DE ZONA']=='outdoor']

    #TRAFICO_TRATADA_INDOOR_ONLINE=TRAFICO_TRATADA_INDOOR[TRAFICO_TRATADA_INDOOR['ESTADO']=='online']
    #TRAFICO_TRATADA_INDOOR_OFFLINE=TRAFICO_TRATADA_INDOOR[TRAFICO_TRATADA_INDOOR['ESTADO']=='offline']

    #TRAFICO_TRATADA_OUTDOOR_ONLINE=TRAFICO_TRATADA_OUTDOOR[TRAFICO_TRATADA_OUTDOOR['ESTADO']=='online']
    #TRAFICO_TRATADA_OUTDOOR_OFFLINE=TRAFICO_TRATADA_OUTDOOR[TRAFICO_TRATADA_OUTDOOR['ESTADO']=='offline']

    PQRSD_TRATADA = PQRSD_TRATADA[PQRSD_TRATADA['Estado']!='Cerrar']
    PQRSD_TRATADA_INDOOR = PQRSD_TRATADA[PQRSD_TRATADA['SUBCATEGORIA'].isin([
        'CD-SIN TRAFICO DE INTERNET 24 HORAS AP INTERIOR',
        'CD-FALLA SIN TRAFICO DE INTERNET 24 HORAS AP INTERIOR'
        ])]

    PQRSD_TRATADA_OUTDOOR = PQRSD_TRATADA[PQRSD_TRATADA['SUBCATEGORIA'].isin([
        'CD-SIN TRAFICO DE INTERNET 24 HORAS AP EXTERIOR',
        'CD-FALLA SIN TRAFICO DE INTERNET 24 HORAS AP EXTERIOR'
        ])]

    #print(PQRSD_TRATADA_OUTDOOR[PQRSD_TRATADA_OUTDOOR['ID_Beneficiario'] == 37848])

    #primera revision online y contador mayor a 0
    # Realizar el INNER JOIN
    RESULTADO_INDOOR = pd.merge(
        left=TRAFICO_TRATADA_INDOOR,
        right=CONTADORES_TRATADA_INDOOR,
        left_on='ID BENEFICIARIO',
        right_on='IDENTIFICADOR BENEFICIARIO',
        how='outer'         # Especifica que es un INNER JOIN
    )
    RESULTADO_INDOOR = pd.merge(
        left=RESULTADO_INDOOR,
        right=SITIOS_TRATADA,
        left_on='ID BENEFICIARIO',
        right_on='Id_Beneficiario',
        how='outer'         # Especifica que es un INNER JOIN
    )
    RESULTADO_INDOOR = pd.merge(
        left=RESULTADO_INDOOR,
        right=PQRSD_TRATADA_INDOOR,
        left_on='ID BENEFICIARIO',
        right_on='ID_Beneficiario',
        how='outer'         # Especifica que es un INNER JOIN
    )

    RESULTADO_OUTDOOR = pd.merge(
        left=TRAFICO_TRATADA_OUTDOOR,
        right=CONTADORES_TRATADA_OUTDOOR,
        left_on='ID BENEFICIARIO',
        right_on='IDENTIFICADOR BENEFICIARIO',
        how='outer'         # Especifica que es un INNER JOIN
    )
    RESULTADO_OUTDOOR = pd.merge(
        left=RESULTADO_OUTDOOR,
        right=SITIOS_TRATADA,
        left_on='ID BENEFICIARIO',
        right_on='Id_Beneficiario',
        how='outer'         # Especifica que es un INNER JOIN
    )
    RESULTADO_OUTDOOR = pd.merge(
        left=RESULTADO_OUTDOOR,
        right=PQRSD_TRATADA_OUTDOOR,
        left_on='ID BENEFICIARIO',
        right_on='ID_Beneficiario',
        how='outer'         # Especifica que es un INNER JOIN
    )


    es_ping_mayor_7 = RESULTADO_INDOOR['ULTIMO PING ONLINE'].astype(str).str.contains('mas de 7', case=False, na=False)
    ping_indoor_7_dias = RESULTADO_INDOOR[es_ping_mayor_7].copy()
    RESULTADO_INDOOR = RESULTADO_INDOOR[~es_ping_mayor_7].copy()
    ping_indoor_7_dias['VERIFICACION'] = ping_indoor_7_dias['CONTADOR DEL DIA'] > 7


    RESULTADO_INDOOR['ULTIMO PING ONLINE'] = pd.to_datetime(RESULTADO_INDOOR['ULTIMO PING ONLINE'])

    RESULTADO_INDOOR['Dias_diferencia'] = (
     (pd.Timestamp('today') - pd.Timedelta(days=1)).normalize() - RESULTADO_INDOOR['ULTIMO PING ONLINE'].dt.normalize()
    ).dt.days

    RESULTADO_INDOOR ['VERIFICACION'] = RESULTADO_INDOOR['Dias_diferencia']==RESULTADO_INDOOR['CONTADOR DEL DIA']

    RESULTADO_INDOOR = pd.concat([RESULTADO_INDOOR, ping_indoor_7_dias], ignore_index=True)

    RESULTADO_INDOOR.to_csv('cruce_trafico_indoor.csv', index=False)

    es_ping_mayor_7 = RESULTADO_OUTDOOR['ULTIMO PING ONLINE'].astype(str).str.contains('mas de 7', case=False, na=False)
    ping_outdoor_7_dias = RESULTADO_OUTDOOR[es_ping_mayor_7].copy()
    RESULTADO_OUTDOOR = RESULTADO_OUTDOOR[~es_ping_mayor_7].copy()
    
    ping_outdoor_7_dias['VERIFICACION'] = ping_outdoor_7_dias['CONTADOR DEL DIA'] > 7


    RESULTADO_OUTDOOR['ULTIMO PING ONLINE'] = pd.to_datetime(RESULTADO_OUTDOOR['ULTIMO PING ONLINE'])

    RESULTADO_OUTDOOR['Dias_diferencia'] = (
     (pd.Timestamp('today') - pd.Timedelta(days=1)).normalize() - RESULTADO_OUTDOOR['ULTIMO PING ONLINE'].dt.normalize()
    ).dt.days


    RESULTADO_OUTDOOR ['VERIFICACION'] = RESULTADO_OUTDOOR['Dias_diferencia']==RESULTADO_OUTDOOR['CONTADOR DEL DIA']

    RESULTADO_OUTDOOR = pd.concat([RESULTADO_OUTDOOR, ping_outdoor_7_dias], ignore_index=True)

    cruce['CANTIDAD_PQRSD'] = cruce.groupby('DISPOSITIVO')['ID'].transform('count')

    return (RESULTADO_OUTDOOR,RESULTADO_INDOOR)




