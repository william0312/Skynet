# -*- coding: latin-1 -*-
import pandas as pd
import numpy as np


def ejecutar_cruce_seguro():
    url = "https://jacs.ruijienetworks.com/api/users/login?version=2"

    # Datos a enviar en formato JSON
    payload = {
        "account": "centrosdigitales@sky.net.co",
        "password": "Skynet2023*"
    }

    # Hacemos la petición POST
    response = requests.post(url, json=payload)

    if response.status_code == 200:
        sesion = response.json()  # Convertimos a dict

        # Accedemos al access_token
        access_token = sesion.get("data", {}).get("access_token")

        print("Access Token:", access_token[:10], '.....')

    else:
        print("Error en la petición:", response.status_code, response.text)


    # Zona horaria de Bogotá
    bogota_tz = ZoneInfo("America/Bogota")

    # Fecha actual en Bogotá
    now_bogota = datetime.now(bogota_tz)

    # Convertir a timestamp (epoch unix en segundos)
    epoch_unix = int(now_bogota.timestamp())

    print("Fecha Bogotá:", now_bogota)
    print("Epoch Unix:", epoch_unix)

    # Nombre de archivo con fecha actual
    now_bogota = datetime.now(bogota_tz)
    fecha_archivo = now_bogota.strftime("%Y-%m-%d_%H-%M")

    # URL base del endpoint
    url = "https://jacs.ruijienetworks.com/service/api/maint/devices?"

    # Parámetros que van en la URL (?param1=valor1&param2=valor2...)
    params = {
        "version": "2",
        "sort_field": "createTime",
        "sort_type": "desc",
        "common_type": "AP",
        "group_id": "8464",
        "project_type": "project",
        "page": 1,
        "per_page":"30000",
        "access_token":access_token,
        "_t":epoch_unix
    }

    # Headers (usa los que ya tienes funcionando)
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "User-Agent": "python-requests/2.x"
    }


    # Hacemos la petición GET
    response = requests.get(url, params=params, headers=headers)

    # Verificamos la respuesta
    if response.status_code == 200:
        data = response.json()
    else:
        print("Error:", response.status_code, response.text)

    df = pd.DataFrame(data)

    # Exportar a un archivo .json
    df.to_json(f'Dispositivos controladora {fecha_archivo}.json', orient='records', lines=True, force_ascii=False)


    # Extraer lista de dispositivos
    device_list2 = data.get("deviceList", [])

    # Convertir en DataFrame
    df = pd.DataFrame(device_list2)

    # Convertir a datetime con timestamps en milisegundos
    df["lastOnline_dt"] = pd.to_datetime(df["lastOnline"], unit="ms") - pd.Timedelta(hours=5)
    df["createTime_dt"] = pd.to_datetime(df["createTime"], unit="ms") - pd.Timedelta(hours=5)

    # Exportar
    df.to_csv(f"Dispositivos controladora {fecha_archivo}.csv", index=False)

    print("Archivo exportado correctamente.")
    return(df)
    
