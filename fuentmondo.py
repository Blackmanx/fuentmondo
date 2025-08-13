import requests
import json
import copy
import openpyxl

#
# Carga los datos del payload desde un archivo JSON local.
#
def cargar_payload(ruta_archivo):
    """
    Lee y devuelve el contenido de un archivo JSON.
    """
    try:
        with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
            return json.load(archivo)
    except FileNotFoundError:
        print(f"Error: El archivo '{ruta_archivo}' no se encontró.")
        return None
    except json.JSONDecodeError:
        print(f"Error: El archivo '{ruta_archivo}' no tiene un formato JSON válido.")
        return None

#
# Realiza una petición POST a la API de Futmondo con el payload proporcionado.
#
def llamar_api(url, payload):
    """
    Realiza una petición POST y gestiona la respuesta.
    """
    if not payload:
        return None

    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()  # Lanza una excepción para errores de HTTP
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error al realizar la petición a la API en {url}: {e}")
        return None

#
# Función para guardar la respuesta de la API en un archivo JSON.
#
def guardar_respuesta(datos, nombre_archivo):
    """
    Guarda los datos JSON en un archivo local para su posterior análisis.
    """
    try:
        with open(nombre_archivo, 'w', encoding='utf-8') as f:
            json.dump(datos, f, indent=4, ensure_ascii=False)
        print(f"Respuesta de la API guardada en '{nombre_archivo}'.")
    except Exception as e:
        print(f"Error al guardar el archivo '{nombre_archivo}': {e}")

#
# Función para actualizar la hoja de Excel con los datos de la clasificación.
#
def actualizar_excel(datos_general, datos_teams, nombre_archivo_excel, sheet_name, fila_inicio):
    """
    Sobreescribe una hoja específica del archivo Excel con los datos de la API.
    """
    try:
        # Cargar el archivo de Excel
        workbook = openpyxl.load_workbook(nombre_archivo_excel)
        sheet = workbook[sheet_name]
        
        print(f"Hoja '{sheet_name}' cargada con éxito.")

        # Mapear los nombres de los equipos a sus puntos generales
        puntos_generales = {equipo['teamname']: equipo['points'] for equipo in datos_teams['answer']['teams']}
        
        # Obtener el ranking de la API, que ya viene ordenado
        ranking_general = datos_general['answer']['ranking']

        # Limpiar las filas existentes antes de escribir los nuevos datos
        # Esto es importante para sobreescribir el orden y los valores
        for row in sheet[f'{fila_inicio}:{sheet.max_row}']:
            for cell in row:
                cell.value = None

        # Escribir los nuevos datos
        for i, equipo in enumerate(ranking_general):
            fila_actual = fila_inicio + i
            nombre_equipo = equipo['name']
            
            # Escribir la posición
            sheet.cell(row=fila_actual, column=1).value = f"{i+1}º"
            
            # Escribir el nombre del equipo
            sheet.cell(row=fila_actual, column=2).value = nombre_equipo
            
            # Escribir los puntos totales (del ranking general)
            sheet.cell(row=fila_actual, column=3).value = equipo['points']

            # Escribir los puntos generales (del endpoint de equipos)
            # Se utiliza el nombre del equipo para buscar los puntos generales.
            # No se necesita un mapeo manual si los nombres coinciden en ambos JSON.
            # Si en algún momento no coinciden, se debería implementar un mapeo.
            
            # Escribir los puntos generales
            sheet.cell(row=fila_actual, column=4).value = puntos_generales.get(nombre_equipo, 0)
        
        # Guardar el archivo
        workbook.save(nombre_archivo_excel)
        print(f"El archivo '{nombre_archivo_excel}' ha sido actualizado con éxito. ¡Se han sobreescrito los datos!")

    except FileNotFoundError:
        print(f"Error: El archivo de Excel '{nombre_archivo_excel}' no se encontró.")
    except KeyError as e:
        print(f"Error: No se encontró la hoja '{sheet_name}' o una clave en el JSON: {e}")
    except Exception as e:
        print(f"Ocurrió un error al procesar el archivo de Excel: {e}")

#
# Función que procesa los datos y actualiza una hoja específica del Excel.
#
def procesar_y_actualizar_division(payload_file, excel_file, sheet_name, start_row):
    """
    Coordina la carga del payload, las llamadas a la API y la actualización del Excel
    para una división específica.
    """
    payload = cargar_payload(payload_file)
    if not payload:
        return
    
    # URLs de la API
    API_URL_GENERAL = "https://api.futmondo.com/1/ranking/general"
    API_URL_TEAMS = "https://api.futmondo.com/2/championship/teams"
    
    # --- Llamada 1: Obtener la clasificación general ---
    print(f"\n--- Llamando a la API de clasificación general para '{sheet_name}' ---")
    payload_general = copy.deepcopy(payload)
    if 'roundId' in payload_general['query']:
        payload_general['query'].pop('roundId', None)
    
    datos_general = llamar_api(API_URL_GENERAL, payload_general)
    if datos_general:
        guardar_respuesta(datos_general, f"respuesta_general_{sheet_name.replace(' ', '_')}.json")
    
    # --- Llamada 2: Obtener los equipos con el nuevo endpoint ---
    print(f"\n--- Llamando a la API de equipos para '{sheet_name}' ---")
    payload_teams = copy.deepcopy(payload)
    payload_teams['query'] = {"championshipId": payload_teams["query"]["championshipId"]}
    datos_teams = llamar_api(API_URL_TEAMS, payload_teams)
    if datos_teams:
        guardar_respuesta(datos_teams, f"respuesta_teams_{sheet_name.replace(' ', '_')}.json")
        
    # --- Paso final: Actualizar el archivo de Excel si todas las llamadas fueron exitosas ---
    if datos_general and datos_teams:
        print(f"\n--- Procesando y actualizando el archivo de Excel para '{sheet_name}' ---")
        actualizar_excel(datos_general, datos_teams, excel_file, sheet_name, start_row)
    else:
        print(f"\nFallo al obtener los datos necesarios para '{sheet_name}'. No se pudo actualizar el Excel.")


#
# Función principal del script.
#
def main():
    """
    Función principal que coordina las llamadas a la API y la gestión de archivos para ambas divisiones.
    """
    # Configuración de la 1a División
    PAYLOAD_PRIMERA = "payload_primera.json"
    EXCEL_FILE_PRIMERA = "SuperLiga Fuentmondo 25-26.xlsx"
    SHEET_PRIMERA = "Clasificación 1a DIV"
    ROW_START_PRIMERA = 2

    # Configuración de la 2a División
    PAYLOAD_SEGUNDA = "payload.json"
    EXCEL_FILE_SEGUNDA = "SuperLiga Fuentmondo 25-26.xlsx"
    SHEET_SEGUNDA = "Clasificación 2a DIV"
    ROW_START_SEGUNDA = 2

    # Procesar la 1a División
    procesar_y_actualizar_division(PAYLOAD_PRIMERA, EXCEL_FILE_PRIMERA, SHEET_PRIMERA, ROW_START_PRIMERA)

    # Procesar la 2a División
    procesar_y_actualizar_division(PAYLOAD_SEGUNDA, EXCEL_FILE_SEGUNDA, SHEET_SEGUNDA, ROW_START_SEGUNDA)

if __name__ == '__main__':
    main()
