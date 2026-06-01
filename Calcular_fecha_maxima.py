#procedimientos para calcular tiempos fuera de horario y festivos


import pandas as pd
import numpy as np
from datetime import datetime, timedelta
#!pip install holidays
#import holidays
# -*- coding: latin-1 -*-

'''
path1 = '/content/PQRSD_Evento_Excepcional_Interventoria.xlsx'
PQRSD_Evento_Excepcional_Interventoria = pd.read_excel(path1)
PQRSD_Evento_Excepcional_Interventoria.shape

path2 = '/content/PQRSD_Formato_Servicios.xlsx'
PQRSD_Formato_Servicios = pd.read_excel(path2)
PQRSD_Formato_Servicios.shape

path3 = '/content/GEN_Sitios_Activos_SI.xlsx'
GEN_Sitios_Activos_SI = pd.read_excel(path3)
GEN_Sitios_Activos_SI.shape

path4 = '/content/PR_Evento_Excepcional.xlsx'
PR_Evento_Excepcional = pd.read_excel(path4)
PR_Evento_Excepcional.shape
'''

##carga archivos locales
'''
# 1. PQRSD Interventoría
path1 = 'PQRSD_Evento_Excepcional_Interventoria.xlsx'
PQRSD_Evento_Excepcional_Interventoria = pd.read_excel(path1)
print(f"PQRSD Interventoría cargado. Filas y columnas: {PQRSD_Evento_Excepcional_Interventoria.shape}")

# 2. PQRSD Formato Servicios
path2 = 'PQRSD_Formato_Servicios.xlsx'
PQRSD_Formato_Servicios = pd.read_excel(path2)
print(f"PQRSD Formato Servicios cargado. Filas y columnas: {PQRSD_Formato_Servicios.shape}")

# 3. GEN Sitios Activos SI
path3 = 'GEN_Sitios_Activos_SI.xlsx'
GEN_Sitios_Activos_SI = pd.read_excel(path3)
print(f"Sitios Activos cargado. Filas y columnas: {GEN_Sitios_Activos_SI.shape}")

# 4. PR Evento Excepcional
path4 = 'PR_Evento_Excepcional.xlsx'
PR_Evento_Excepcional = pd.read_excel(path4)
print(f"PR Evento Excepcional cargado. Filas y columnas: {PR_Evento_Excepcional.shape}")

print("¡Todos los archivos se cargaron correctamente!")

'''

def parse_hora(hora_str, fecha):
    """Convierte string de hora en datetime, acepta HH:MM y HH:MM:SS."""
    if pd.isna(hora_str):
        return None
    try:
        return datetime.combine(fecha, datetime.strptime(hora_str, "%H:%M").time())
    except ValueError:
        return datetime.combine(fecha, datetime.strptime(hora_str, "%H:%M:%S").time())

def calcular_minutos_dentro(fila, col_inicio, col_fin, festivos=None):


    inicio_periodo = fila[col_inicio]
    fin_periodo = fila[col_fin]


     # Si la prioridad no es 1 → devolver diferencia en días
    if "PRIORIDAD" in fila and fila["PRIORIDAD"] != 1:
        if pd.notna(inicio_periodo) and pd.notna(fin_periodo):
            return (fin_periodo - inicio_periodo).days
        else:
            return 0

#    inicio_periodo = fila["fecha_maxima_generada"]
#    fin_periodo = fila["fecha_comparacion"]
    total_dentro = 0

    actual = inicio_periodo
    while actual.date() <= fin_periodo.date():
        # Verificar si es festivo → se ignora
        if festivos and actual.date() in festivos:
            actual += timedelta(days=1)
            continue

        # Día de la semana en español
        dia_semana = actual.strftime("%A").lower()
        dia_semana = (dia_semana.replace("monday", "lunes")
                                 .replace("tuesday", "martes")
                                 .replace("wednesday", "miercoles")
                                 .replace("thursday", "jueves")
                                 .replace("friday", "viernes")
                                 .replace("saturday", "sabado")
                                 .replace("sunday", "domingo"))

        inicio_str = fila[f"inicio_{dia_semana}"]
        fin_str = fila[f"fin_{dia_semana}"]

        # Intervalo real del día en curso
        dia_inicio = datetime.combine(actual.date(), datetime.min.time())
        dia_fin = datetime.combine(actual.date(), datetime.max.time())
        rango_inicio = max(inicio_periodo, dia_inicio)
        rango_fin = min(fin_periodo, dia_fin)

        if rango_inicio < rango_fin and pd.notna(inicio_str) and pd.notna(fin_str):
            # Horario de apertura/cierre de ese día
            h_inicio = parse_hora(inicio_str, actual.date())
            h_fin = parse_hora(fin_str, actual.date())

            # Intersección entre el periodo y el horario de atención
            dentro_inicio = max(rango_inicio, h_inicio)
            dentro_fin = min(rango_fin, h_fin)

            if dentro_inicio < dentro_fin:
                minutos_dentro = (dentro_fin - dentro_inicio).total_seconds() / 60
                total_dentro += minutos_dentro

        actual += timedelta(days=1)

    return int(total_dentro)


###segundo bloque de metodos

def parse_hora2(hora_str, fecha):
    """Convierte un string HH:MM o HH:MM:SS en datetime con la fecha dada."""
    hora_str = str(hora_str).strip()
    try:
        # Acepta "8:00" o "08:00:00"
        if len(hora_str.split(":")) == 2:
            h = datetime.strptime(hora_str, "%H:%M").time()
        else:
            h = datetime.strptime(hora_str, "%H:%M:%S").time()
        return datetime.combine(fecha.date(), h)
    except Exception:
        return None

def calcular_fecha_visita(row, festivos=None):
    actual = pd.to_datetime(row["fecha_final"], errors="coerce")
    #actual = row["fecha_creacion_agenda"]           # punto de partida
    minutos_restantes = row["minutos_disponibles"]  # minutos que faltan

  # 🔹 Si la prioridad no es 1 → simplemente sumar días según minutos_disponibles
    if "PRIORIDAD" in row and row["PRIORIDAD"] != 1:
        return actual + timedelta(minutes=minutos_restantes)

    while minutos_restantes > 0:
        # Saltar festivos
        if festivos and actual.date() in festivos:
            #print(f"[DEBUG] Festivo {actual.date()} → se salta")
            actual = (actual + timedelta(days=1)).replace(hour=0, minute=0, second=0)
            continue

        # Día en español
        dia_semana = actual.strftime("%A").lower()
        dia_semana = (dia_semana.replace("monday", "lunes")
                                   .replace("tuesday", "martes")
                                   .replace("wednesday", "miercoles")
                                   .replace("thursday", "jueves")
                                   .replace("friday", "viernes")
                                   .replace("saturday", "sabado")
                                   .replace("sunday", "domingo"))

        inicio_str = row.get(f"inicio_{dia_semana}", None)
        fin_str = row.get(f"fin_{dia_semana}", None)

        print(f"[DEBUG] Día {dia_semana}: inicio={inicio_str}, fin={fin_str}, actual={actual}")

        if pd.isna(inicio_str) or pd.isna(fin_str):
            # No hay jornada → siguiente día
            actual = (actual + timedelta(days=1)).replace(hour=0, minute=0, second=0)
            continue

        h_inicio = parse_hora2(inicio_str, actual)
        h_fin = parse_hora2(fin_str, actual)

        if not h_inicio or not h_fin:
            # Si no se pudo parsear → pasar al día siguiente
            actual = (actual + timedelta(days=1)).replace(hour=0, minute=0, second=0)
            continue

        # Aseguramos que actual esté dentro de la jornada
        if actual < h_inicio:
            actual = h_inicio
        if actual > h_fin:
            # Ya pasó la jornada, ir al siguiente día
            actual = (actual + timedelta(days=1)).replace(hour=0, minute=0, second=0)
            continue

        # Minutos disponibles en la jornada desde "actual"
        minutos_jornada = (h_fin - actual).total_seconds() / 60

        if minutos_jornada >= minutos_restantes:
            # Se alcanza a terminar dentro de este día
            return actual + timedelta(minutes=minutos_restantes)
        else:
            # Consumimos la jornada y seguimos al día siguiente
            minutos_restantes -= minutos_jornada
            actual = (actual + timedelta(days=1)).replace(hour=0, minute=0, second=0)

    return actual

def ejecutar_cruce_seguro(CONTADORES,TRAFICO,SITIOS,PQRSD):

columnas_GEN_Sitios_Activos_SI = [
    'Id_Beneficiario','DDA_Horas','Estado','Consolidado_Lunes','Consolidado_Martes','Consolidado_Miercoles','Consolidado_Jueves','Consolidado_Viernes','Consolidado_Sabado','Consolidado_Domingo'
]
#Subconjuntos con columnas relevantes
GEN_Sitios_Activos_SI = GEN_Sitios_Activos_SI[columnas_GEN_Sitios_Activos_SI]


GEN_Sitios_Activos_SI[["inicio_lunes", "fin_lunes"]] = GEN_Sitios_Activos_SI["Consolidado_Lunes"].str.split(" a ", expand=True)
GEN_Sitios_Activos_SI[["inicio_martes", "fin_martes"]] = GEN_Sitios_Activos_SI["Consolidado_Martes"].str.split(" a ", expand=True)
GEN_Sitios_Activos_SI[["inicio_miercoles", "fin_miercoles"]] = GEN_Sitios_Activos_SI["Consolidado_Miercoles"].str.split(" a ", expand=True)
GEN_Sitios_Activos_SI[["inicio_jueves", "fin_jueves"]] = GEN_Sitios_Activos_SI["Consolidado_Jueves"].str.split(" a ", expand=True)
GEN_Sitios_Activos_SI[["inicio_viernes", "fin_viernes"]] = GEN_Sitios_Activos_SI["Consolidado_Viernes"].str.split(" a ", expand=True)
GEN_Sitios_Activos_SI[["inicio_sabado", "fin_sabado"]] = GEN_Sitios_Activos_SI["Consolidado_Sabado"].str.split(" a ", expand=True)
GEN_Sitios_Activos_SI[["inicio_domingo", "fin_domingo"]] = GEN_Sitios_Activos_SI["Consolidado_Domingo"].str.split(" a ", expand=True)

#Organizar PQRSD_Formato_Servicios

columnas_PQRSD_Formato_Servicios = [
    'ID','TICKETCCC','ID_Beneficiario','DEPARTAMENTO','CIUDAD','GRUPO','DDA','CATEGORIA','SUBCATEGORIA','PRIORIDAD','descripcion_creacion','Fecha_Creacion','Estado','fecha_creacion_agenda','fecha_maxima_atencion','Nueva_fecha_maxima_atencion'
]
# Subconjuntos con columnas relevantes
PQRSD_Formato_Servicios = PQRSD_Formato_Servicios[columnas_PQRSD_Formato_Servicios]
#PQRSD_Formato_Servicios = PQRSD_Formato_Servicios[
#    (PQRSD_Formato_Servicios['Estado'] == 'Agendar')
# ]


#Organizar PQRSD_Evento_Excepcional_Interventoria

columnas_PQRSD_Evento_Excepcional_Interventoria = [
    'ID','UUID_PQRSD','id_beneficiario','Inicio_Parada_Reloj','Fin_Parada_Reloj'
]
PQRSD_Evento_Excepcional_Interventoria = PQRSD_Evento_Excepcional_Interventoria[columnas_PQRSD_Evento_Excepcional_Interventoria]



columnas_PR_Evento_Excepcional = [
    'ID','Fecha_Creacion_Plantilla','id_beneficiario','dateparada','finishparada'
]
PR_Evento_Excepcional = PR_Evento_Excepcional[columnas_PR_Evento_Excepcional]



PQRSD_Evento_Excepcional_Interventoria_tratada=PQRSD_Evento_Excepcional_Interventoria.copy()
PQRSD_Formato_Servicios_tratada=PQRSD_Formato_Servicios.copy()
GEN_Sitios_Activos_SI_tratada=GEN_Sitios_Activos_SI.copy()
PR_Evento_Excepcional_tratada=PR_Evento_Excepcional.copy()



####inicia



#INICIO
#
#
#
#unir dataframe

pqrsd_unido= pd.merge(PQRSD_Formato_Servicios_tratada,GEN_Sitios_Activos_SI_tratada, left_on=['ID_Beneficiario'],right_on =['Id_Beneficiario'], how='left')
print("Paso 1:", pqrsd_unido.shape[0])

pqrsd_unido_1 = pd.merge(pqrsd_unido,PQRSD_Evento_Excepcional_Interventoria_tratada, left_on=['ID'],right_on =['UUID_PQRSD'], how='left')
print("Paso 2:", pqrsd_unido_1.shape[0])
#pqrsd_unido_1.to_excel('df_cruce_inicial.xlsx', index=False)


pqrsd_unido_2 = pd.merge(pqrsd_unido_1,PR_Evento_Excepcional_tratada, left_on=['ID_y'],right_on =['ID'], how='left')
print("Paso 3:", pqrsd_unido_2.shape[0])

pqrsd_unido_2 = pqrsd_unido_2.dropna(subset=['fecha_creacion_agenda'])
pqrsd_unido_2 = pqrsd_unido_2[pqrsd_unido_2['finishparada'] >= pqrsd_unido_2['fecha_creacion_agenda']]
print("Paso 3.1:", pqrsd_unido_2.shape[0])
#pqrsd_unido_2.to_excel('df_cruce_inicial.xlsx', index=False)

#se modifica el dda para los pqrsd que tengan prioridad diferente de 1


# condiciones
pqrsd_unido_2["PRIORIDAD"] = pd.to_numeric(pqrsd_unido_2["PRIORIDAD"], errors='coerce')
condiciones_dda = [
    (pqrsd_unido_2["PRIORIDAD"] ==1),
    (pqrsd_unido_2["PRIORIDAD"] !=1) & (pqrsd_unido_2["DDA"] =="BAJO") ,
    (pqrsd_unido_2["PRIORIDAD"] !=1) & (pqrsd_unido_2["DDA"] =="MEDIO") ,
    (pqrsd_unido_2["PRIORIDAD"] !=1) & (pqrsd_unido_2["DDA"] =="ALTO") ,
    (pqrsd_unido_2["PRIORIDAD"] !=1) & (pqrsd_unido_2["DDA"] =="MUY ALTO")
]

# valores según condición
valores = [pqrsd_unido_2["DDA_Horas"], 168,240,288,360]

# columna nueva
pqrsd_unido_2["DDA_Horas"] = np.select(condiciones_dda, valores, default=0)





# condiciones
condiciones = [
    (pqrsd_unido_2["fecha_creacion_agenda"].notna()) & (pqrsd_unido_2["dateparada"] <= pqrsd_unido_2["fecha_creacion_agenda"]),
    (pqrsd_unido_2["fecha_creacion_agenda"].notna()) & (pqrsd_unido_2["dateparada"] > pqrsd_unido_2["fecha_creacion_agenda"])
]

# valores según condición
valores = [1, 2]

# columna nueva
pqrsd_unido_2["bandera_inicio_parada"] = np.select(condiciones, valores, default=0)

# Ordenar por ID_x
pqrsd_unido_2 = pqrsd_unido_2.sort_values(by='ID_x')

# Crear columna de índice (1,2,3,...)
#pqrsd_unido_2['indice'] = range(1, len(pqrsd_unido_2) + 1)

#solo trabajo con el valor 1, y agrupo para dejar el finish parada mayor
# Filtrar solo bandera_inicio_parada = 1
valor1 = pqrsd_unido_2[pqrsd_unido_2['bandera_inicio_parada'] == 1]
# Obtener índice del máximo finishparada por cada ID_x
agrupado1 = valor1.groupby('ID_x')['finishparada'].idxmax()
# Usar esos índices para traer todas las columnas
resultadovalor1 = valor1.loc[agrupado1]
# columna nueva
resultadovalor1["fecha_maxima_generada"] = resultadovalor1 ["finishparada"]

valor2 = pqrsd_unido_2[pqrsd_unido_2['bandera_inicio_parada'] == 2]
#union entre valor 1 y 2
df_unido = pd.concat([valor2, resultadovalor1], ignore_index=True)
df_unido = df_unido.sort_values(by='ID_x')
df_unido['fecha_maxima_generada'] = pd.to_datetime(df_unido['fecha_maxima_generada'], errors='coerce')
df_unido['fecha_maxima_generada'] = df_unido.groupby('ID_x')['fecha_maxima_generada'].transform(lambda x: x.fillna(x.max()))

df_unido["dateparada"] = pd.to_datetime(df_unido["dateparada"])
df_unido["fecha_maxima_generada"] = pd.to_datetime(df_unido["fecha_maxima_generada"])
df_unido["fecha_maxima_generada"] = df_unido["fecha_maxima_generada"].fillna(df_unido["fecha_creacion_agenda"])


#df_unido.to_excel('df_inicial.xlsx', index=False)
#
#
#

#re validacion para sacar todas las paradas que no tuvieron salto pero si alargan el tiempo ejm
#parada 1  inicio 01-01-2025   finalizo 02-01-2025
#parada 2 inicio 02-01-2025 finalizo 04-01-2025
#realizo nuevamente la validacion por si al cambiar la fecha_maxima_generada entras paradas tipo 1


cambio = True
iteracion = 0
max_iter = 20   # límite de seguridad (máx. 20 iteraciones)

while cambio and iteracion < max_iter:
    iteracion += 1
    print(f"--- Iteración {iteracion} ---")

    # Asegurar formatos de fechas
    df_unido["dateparada"] = pd.to_datetime(df_unido["dateparada"])
    df_unido["finishparada"] = pd.to_datetime(df_unido["finishparada"])
    df_unido["fecha_maxima_generada"] = pd.to_datetime(df_unido["fecha_maxima_generada"], errors="coerce")

    # Guardar estado anterior
    df_prev = df_unido.copy()

    # condiciones
    condiciones = [
        (df_unido["fecha_maxima_generada"].notna()) & (df_unido["dateparada"] <= df_unido["fecha_maxima_generada"]),
        (df_unido["fecha_maxima_generada"].notna()) & (df_unido["dateparada"] > df_unido["fecha_maxima_generada"])
    ]
    valores = [1, 2]

    df_unido["bandera_inicio_parada"] = np.select(condiciones, valores, default=0)

    # --- Proceso con valor 1
    valor1 = df_unido[df_unido['bandera_inicio_parada'] == 1]
    agrupado1 = valor1.groupby('ID_x')['finishparada'].idxmax()
    resultadovalor1 = valor1.loc[agrupado1].copy()
    resultadovalor1["fecha_maxima_generada"] = resultadovalor1["finishparada"]

    # --- Proceso con valor 2
    valor2 = df_unido[df_unido['bandera_inicio_parada'] == 2]

    # --- Unión
    df_unido = pd.concat([valor2, resultadovalor1], ignore_index=True)
    df_unido = df_unido.sort_values(by="ID_x")

    # --- Asignar la fecha máxima final
    df_unido["fecha_maxima_generada"] = (
        df_unido.groupby("ID_x")["fecha_maxima_generada"].transform("max")
    )

    # --- Verificar si hubo cambios
    cambio = not df_unido[["ID_x", "fecha_maxima_generada"]].equals(
        df_prev[["ID_x", "fecha_maxima_generada"]]
    )
print(f"Finalizó en {iteracion} iteraciones")

#diferencia en minutos
df_unido["minutos_acomulados"] = (
    (df_unido["dateparada"] - df_unido["fecha_maxima_generada"])
    .dt.total_seconds() / 60
)
#convierto valores negativos a 0
df_unido["minutos_acomulados"] = df_unido["minutos_acomulados"].where(df_unido["minutos_acomulados"] >= 0, 0)

df_unido["fecha_comparacion"] = (
    (df_unido["fecha_maxima_generada"] + pd.to_timedelta(df_unido["minutos_acomulados"], unit="m"))
)

#
#
#
##festivos
#
##

# Festivos de Colombia
# Pasando una lista de años
festivos_col = holidays.Colombia(years=[2024, 2025, 2026])
df_unido["minutos_fuera"] = df_unido.apply(
    calcular_minutos_dentro,
    axis=1,
    col_inicio="fecha_maxima_generada",
    col_fin="fecha_comparacion",
    festivos=festivos_col
)


#
##
#
#para las bandera 2 dejo el valor minimo de minutos_fuera porque solo le realizo el calculo a la sigueitne banderaparada2

# Filtrar solo los que tienen bandera = 2
mask = df_unido["bandera_inicio_parada"] == 2
# Calcular el mínimo por ID_x
minimos = df_unido.loc[mask].groupby("ID_x")["minutos_fuera"].transform("min")
# Asignar el mínimo a todos los registros bandera=2
df_unido.loc[mask, "minutos_fuera"] = minimos




#se revisa si se debe seguir revisando ese pqrsd
condiciones2 = [
    (df_unido["minutos_fuera"] < (df_unido["DDA_Horas"]*60)),
    (df_unido["minutos_fuera"] >= (df_unido["DDA_Horas"]*60))
]
# valores según condición
valores2 = [1, 2]
# columna nueva
df_unido["continua"] = np.select(condiciones2, valores2, default=0)

# DataFrame con los que tienen continua = 2
df_continua_2_1 = df_unido[df_unido["continua"] == 2].copy()
# DataFrame con los que tienen continua = 1
df_continua_1_1 = df_unido[df_unido["continua"] == 1].copy()

df_continua_2_1["fecha_final"] = df_continua_2_1["fecha_maxima_generada"]

#df_unido.to_excel('df_unido.xlsx', index=False)

df_unido = df_continua_1_1.copy()



#
#
#
#
#ahora si valido la siguiente parada que tuvo un salto, genero indisponibilidad

#segundo archivo
# condiciones
condiciones2 = [
    (df_unido["fecha_comparacion"].notna()) & (df_unido["dateparada"] <= df_unido["fecha_comparacion"]),
    (df_unido["fecha_comparacion"].notna()) & (df_unido["dateparada"] > df_unido["fecha_comparacion"])
]

# valores según condición
valores2 = [1, 2]

# columna nueva
df_unido["bandera_inicio_parada2"] = np.select(condiciones2, valores2, default=0)


# Ordenar por ID_x
df_unido = df_unido.sort_values(by='ID_x')
#solo trabajo con el valor 1, y agrupo para dejar el finish parada mayor
# Filtrar solo bandera_inicio_parada = 1
valor2_1 = df_unido[df_unido['bandera_inicio_parada2'] == 1]
# Obtener índice del máximo finishparada por cada ID_x
agrupado2_1 = valor2_1.groupby('ID_x')['finishparada'].idxmax()
# Usar esos índices para traer todas las columnas
resultadovalor2_1 = valor2_1.loc[agrupado2_1]
# columna nueva
resultadovalor2_1["fecha_maxima_generada2"] = resultadovalor2_1 ["finishparada"]


valor2_2 = df_unido[df_unido['bandera_inicio_parada2'] == 2]

#union entre valor 1 y 2
df_unido_2 = pd.concat([valor2_2, resultadovalor2_1], ignore_index=True)
df_unido_2 = df_unido_2.sort_values(by='ID_x')
df_unido_2['fecha_maxima_generada2'] = pd.to_datetime(df_unido_2['fecha_maxima_generada2'], errors='coerce')
df_unido_2['fecha_maxima_generada2'] = df_unido_2.groupby('ID_x')['fecha_maxima_generada2'].transform(lambda x: x.fillna(x.max()))

#
#
#dejo el minutos_fuera menor para iniciar el siguiente siclo de validacion
# Reemplazar por el mínimo dentro de cada grupo de ID_x
df_unido_2["minutos_fuera"] = df_unido_2.groupby("ID_x")["minutos_fuera"].transform("min")



#se revisa si se debe seguir revisando ese pqrsd
condiciones2 = [
    (df_unido_2["minutos_fuera"] < (df_unido_2["DDA_Horas"]*60)),
    (df_unido_2["minutos_fuera"] >= (df_unido_2["DDA_Horas"]*60))
]
# valores según condición
valores2 = [1, 2]
# columna nueva
df_unido_2["continua"] = np.select(condiciones2, valores2, default=0)

# DataFrame con los que tienen continua = 2
df_continua_2_2 = df_unido_2[df_unido_2["continua"] == 2].copy()
# DataFrame con los que tienen continua = 1
df_continua_1_2 = df_unido_2[df_unido_2["continua"] == 1].copy()

df_continua_2_2["fecha_final"] = df_continua_2_2["fecha_comparacion"]

#df_unido_2.to_excel('df_unido_2.xlsx', index=False)




#copio df_unido_2 a df_unido_3 para trabajar con un nuevo dataframe
df_unido_3 = df_continua_1_2.copy()

#
#
#
# condiciones en el segundo flujo
# 1. Asegurar formato de fecha ANTES de comparar
df_unido_3["dateparada"] = pd.to_datetime(df_unido_3["dateparada"], errors='coerce')
df_unido_3["fecha_maxima_generada2"] = pd.to_datetime(df_unido_3["fecha_maxima_generada2"], errors='coerce')

condiciones3 = [
        (df_unido_3["fecha_maxima_generada2"].notna()) & (df_unido_3["dateparada"] <= df_unido_3["fecha_maxima_generada2"]),
        (df_unido_3["fecha_maxima_generada2"].notna()) & (df_unido_3["dateparada"] > df_unido_3["fecha_maxima_generada2"])
]
valores3 = [1, 2]

df_unido_3["bandera_inicio_parada3"] = np.select(condiciones3, valores3, default=0)

# Ordenar por ID_x
df_unido_3 = df_unido_3.sort_values(by='ID_x')

# Crear columna de índice (1,2,3,...)
#pqrsd_unido_2['indice'] = range(1, len(pqrsd_unido_2) + 1)

#solo trabajo con el valor 1, y agrupo para dejar el finish parada mayor
# Filtrar solo bandera_inicio_parada = 1
valor1 = df_unido_3[df_unido_3['bandera_inicio_parada3'] == 1]
# Obtener índice del máximo finishparada por cada ID_x
agrupado1 = valor1.groupby('ID_x')['finishparada'].idxmax()
# Usar esos índices para traer todas las columnas
resultadovalor1 = valor1.loc[agrupado1]
# columna nueva
resultadovalor1["fecha_maxima_generada3"] = resultadovalor1 ["finishparada"]

valor2 = df_unido_3[df_unido_3['bandera_inicio_parada3'] == 2]
#union entre valor 1 y 2
df_unido_3 = pd.concat([valor2, resultadovalor1], ignore_index=True)
df_unido_3 = df_unido_3.sort_values(by='ID_x')
df_unido_3['fecha_maxima_generada3'] = pd.to_datetime(df_unido_3['fecha_maxima_generada3'], errors='coerce')
df_unido_3['fecha_maxima_generada3'] = df_unido_3.groupby('ID_x')['fecha_maxima_generada3'].transform(lambda x: x.fillna(x.max()))

df_unido_3["dateparada"] = pd.to_datetime(df_unido_3["dateparada"])
df_unido_3["fecha_maxima_generada3"] = pd.to_datetime(df_unido_3["fecha_maxima_generada3"])
df_unido_3["fecha_maxima_generada3"] = df_unido_3["fecha_maxima_generada3"].fillna(df_unido_3["fecha_maxima_generada3"])


#df_unido_3.to_excel('df_unido_3.xlsx', index=False)

#
##
#
#se inicia nuevamente para los sitios que tenian mas de un salto


#re validacion para sacar todas las paradas que no tuvieron salto pero si alargan el tiempo ejm
#parada 1  inicio 01-01-2025   finalizo 02-01-2025
#parada 2 inicio 02-01-2025 finalizo 04-01-2025
#realizo nuevamente la validacion por si al cambiar la fecha_maxima_generada entras paradas tipo 1


cambio = True
iteracion = 0
max_iter = 20   # límite de seguridad (máx. 20 iteraciones)

while cambio and iteracion < max_iter:
    iteracion += 1
    print(f"--- Iteración {iteracion} ---")

    # Asegurar formatos de fechas
    df_unido_3["dateparada"] = pd.to_datetime(df_unido_3["dateparada"])
    df_unido_3["finishparada"] = pd.to_datetime(df_unido_3["finishparada"])
    df_unido_3["fecha_maxima_generada2"] = pd.to_datetime(df_unido_3["fecha_maxima_generada2"], errors="coerce")

    # Guardar estado anterior
    df_prev2 = df_unido_3.copy()

    # condiciones
    condiciones3 = [
        (df_unido_3["fecha_maxima_generada2"].notna()) & (df_unido_3["dateparada"] <= df_unido_3["fecha_maxima_generada2"]),
        (df_unido_3["fecha_maxima_generada2"].notna()) & (df_unido_3["dateparada"] > df_unido_3["fecha_maxima_generada2"])
    ]
    valores3 = [1, 2]

    df_unido_3["bandera_inicio_parada3"] = np.select(condiciones3, valores3, default=0)

    # --- Proceso con valor 1
    valor3_1 = df_unido_3[df_unido_3['bandera_inicio_parada3'] == 1]
    agrupado3 = valor3_1.groupby('ID_x')['finishparada'].idxmax()
    resultadovalor3 = valor3_1.loc[agrupado3].copy()
    resultadovalor3["fecha_maxima_generada3"] = resultadovalor3["finishparada"]



    # --- Proceso con valor 2
    valor3_2 = df_unido_3[df_unido_3['bandera_inicio_parada3'] == 2]

    # --- Unión
    df_unido_3_2 = pd.concat([valor3_2, resultadovalor3], ignore_index=True)
    df_unido_3_2 = df_unido_3_2.sort_values(by="ID_x")

    #df_unido_3_2.to_excel('df_unido_3_2.xlsx', index=False)
    #resultadovalor3.to_excel('resultadovalor3.xlsx', index=False)


    # --- Asignar la fecha máxima final
    df_unido_3_2["fecha_maxima_generada3"] = (
        df_unido_3_2.groupby("ID_x")["fecha_maxima_generada3"].transform("max")
    )

    # --- Verificar si hubo cambios
    cambio = not df_unido_3_2[["ID_x", "fecha_maxima_generada3"]].equals(
        df_prev2[["ID_x", "fecha_maxima_generada3"]]
    )
print(f"Finalizó en {iteracion} iteraciones")

#diferencia en minutos
df_unido_3_2["minutos_acomulados_2"] = (
    (df_unido_3_2["dateparada"] - df_unido_3_2["fecha_maxima_generada3"])
    .dt.total_seconds() / 60
)
#convierto valores negativos a 0
df_unido_3_2["minutos_acomulados_2"] = df_unido_3_2["minutos_acomulados_2"].where(df_unido_3_2["minutos_acomulados_2"] >= 0, 0)

df_unido_3_2["fecha_comparacion_2"] = (
    (df_unido_3_2["fecha_maxima_generada3"] + pd.to_timedelta(df_unido_3_2["minutos_acomulados_2"], unit="m"))
)

#
#
#
#se calculan los minutos dentro

df_unido_3_2["minutos_fuera2"] = df_unido_3_2.apply(
    calcular_minutos_dentro,
    axis=1,
    col_inicio="fecha_maxima_generada3",
    col_fin="fecha_comparacion_2",
    festivos=festivos_col
)

df_unido_3_2["minutos_fuera2"]=df_unido_3_2["minutos_fuera2"]+df_unido_3_2["minutos_fuera"]

#reviso si el dda le alcanza para la segunda validacion

#df_unido_3_2["fecha_comparacion_2"] = np.where(
#    (df_unido_3_2["DDA_Horas"]* 60 ) >= (df_unido_3_2["minutos_fuera"] + df_unido_3_2["minutos_fuera2"]),
#    df_unido_3_2["fecha_comparacion_2"],
#    df_unido_3_2["fecha_maxima_generada3"]
#)

df_unido_3_2.to_excel('df_unido_3_2.xlsx', index=False)


#ahora muestro todo incluyendo los que ya se habian descartado por vencimiento ans
#se realizaron 2 descartes

df_unido_3_2 = df_unido_3_2.rename(columns={"fecha_maxima_generada3": "fecha_final"})
df_unido_3_2["fecha_final"] = pd.to_datetime(df_unido_3_2["fecha_final"].astype(str).str[:16],
                                          format="%Y-%m-%d %H:%M")


df_final = pd.concat([df_unido_3_2, df_continua_2_2, df_continua_2_1], ignore_index=True)

df_final["minutos_fuera2"] = df_final["minutos_fuera2"].fillna(0)

df_final["minutos_disponibles"] = (df_final["DDA_Horas"] * 60) - df_final["minutos_fuera2"]
#df_final["minutos_disponibles"] = np.where(
#    df_final["PRIORIDAD"] == 1,
#    (df_final["DDA_Horas"] * 60) - df_final["minutos_fuera2"],
#    20 - df_final["minutos_fuera2"]
#)

df_final["minutos_disponibles"] = df_final["minutos_disponibles"].where(df_final["minutos_disponibles"] >= 0, 0)

#
#
#
#
#

#calculo la fecha maxima de atencion

df_final = df_final.copy()
       # SLA en horas

# Convertir todo a string primero, luego a datetime
#df_final["fecha_final"] = pd.to_datetime(df_final["fecha_final"].astype(str), errors="coerce", dayfirst=True)
#df_final["fecha_final"] = df_final["fecha_final"].fillna(pd.Timestamp.today())
df_final["DDA_Minutos"] = (df_final["DDA_Horas"]*60)



#df_test = df_final.iloc[:20].copy()
#df_test['fecha_visita'] = df_test.apply(lambda row: calcular_fecha_visita(row, festivos_col), axis=1)


df_final['fecha_visita'] = df_final.apply(lambda row: calcular_fecha_visita(row, festivos_col), axis=1)


#
#
#
columnas_df_final = [
    'bandera_inicio_parada3','ID_x','TICKETCCC','ID_Beneficiario','DEPARTAMENTO','CIUDAD','GRUPO','DDA','CATEGORIA','SUBCATEGORIA','PRIORIDAD','descripcion_creacion','Fecha_Creacion','Estado_x','fecha_creacion_agenda','fecha_maxima_atencion','Nueva_fecha_maxima_atencion','DDA_Horas','ID_y','Inicio_Parada_Reloj','Fin_Parada_Reloj','fecha_final','minutos_disponibles','bandera_inicio_parada3','fecha_visita'
    ]
# Subconjuntos con columnas relevantes
df_final = df_final[columnas_df_final]
return (df_final)
#df_final.to_excel('df_final.xlsx', index=False)








