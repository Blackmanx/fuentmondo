import requests
import json
import copy
import openpyxl
import os
import io
import base64
import msal
import webbrowser
import tkinter as tk
import pyperclip
import time
from collections import defaultdict
from dotenv import load_dotenv

# Carga las variables de entorno desde el archivo .env
load_dotenv()

# --- CONFIGURACIÓN DE MICROSOFT GRAPH (ONEDRIVE) ---
CLIENT_ID = os.getenv("CLIENT_ID")
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
AUTHORITY = 'https://login.microsoftonline.com/common/'
SCOPES = ['Files.ReadWrite.All']
ONEDRIVE_SHARE_LINK = "https://1drv.ms/x/s!AidvQapyuNp6jBKR5uMUCaBYdLl0?e=3kXyKW"

# Muestra una ventana para elegir dónde guardar el archivo.
def choose_save_option():
    """Crea una ventana con botones para que el usuario elija la ubicación de guardado."""
    root = tk.Tk()
    root.title("Opción de Guardado")

    choice = [None]

    # Centrar la ventana
    window_width = 400
    window_height = 150
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    root.attributes('-topmost', True)

    def select_option(option):
        choice[0] = option
        root.destroy()

    tk.Label(root, text="¿Dónde quieres guardar el Excel modificado?", pady=15, font=("Helvetica", 12)).pack()

    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    btn_onedrive = tk.Button(button_frame, text="Guardar en OneDrive", command=lambda: select_option('onedrive'), height=2, width=20, bg="#0078D4", fg="white")
    btn_onedrive.pack(side=tk.LEFT, padx=10)

    btn_local = tk.Button(button_frame, text="Guardar Localmente", command=lambda: select_option('local'), height=2, width=20)
    btn_local.pack(side=tk.RIGHT, padx=10)

    root.mainloop()
    return choice[0]

# Muestra una ventana para copiar el código de autenticación.
def show_auth_code_window(message, verification_uri):
    """Crea una ventana con el código de autenticación para que el usuario lo copie."""
    root = tk.Tk()
    root.title("Código de Autenticación")

    # Centrar la ventana
    window_width = 450
    window_height = 200
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    root.attributes('-topmost', True)

    try:
        user_code = message.split("enter the code ")[1].split(" to authenticate")[0]
    except IndexError:
        user_code = "No se pudo extraer el código"

    def copy_and_open():
        pyperclip.copy(user_code)
        print("Código copiado al portapapeles.")
        webbrowser.open(verification_uri)
        root.destroy()

    tk.Label(root, text="Copia este código y pégalo en la ventana del navegador que se abrirá:", wraplength=420, pady=10).pack()

    code_font = ("Courier", 16, "bold")
    code_entry = tk.Entry(root, justify='center', font=code_font, relief='flat', bd=0, highlightthickness=1)
    code_entry.insert(0, user_code)
    code_entry.config(state='readonly', readonlybackground='white', fg='black')
    code_entry.pack(pady=10, ipady=5)

    tk.Button(root, text="Copiar Código y Abrir Navegador", command=copy_and_open, height=2, bg="#0078D4", fg="white").pack(pady=15, padx=20, fill='x')

    root.mainloop()

# Obtiene un token de acceso para la API de Microsoft Graph de forma interactiva.
def get_access_token():
    """Se autentica de forma interactiva y obtiene un token de acceso."""
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    result = None

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)

        if "error" in flow:
            print(f"\nERROR AL INICIAR LA AUTENTICACIÓN:\nError: {flow.get('error')}\nDescripción: {flow.get('error_description')}")
            return None

        show_auth_code_window(flow["message"], flow["verification_uri"])

        result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        return result['access_token']
    else:
        print("Error al obtener el token de acceso:", result.get("error_description"))
        return None

# Codifica el enlace de compartición de OneDrive a un formato compatible con la API de Graph.
def encode_sharing_link(sharing_link):
    base64_value = base64.b64encode(sharing_link.encode('utf-8')).decode('utf-8')
    return 'u!' + base64_value.rstrip('=').replace('/', '_').replace('+', '-')

# Obtiene el ID del Drive y el ID del archivo a partir de un enlace de compartición.
def get_drive_item_from_share_link(access_token, share_url):
    encoded_url = encode_sharing_link(share_url)
    api_url = f"{GRAPH_API_ENDPOINT}/shares/{encoded_url}/driveItem"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(api_url, headers=headers)
    response.raise_for_status()
    data = response.json()
    return data['parentReference']['driveId'], data['id']

# Descarga el contenido de un archivo Excel desde OneDrive.
def download_excel_from_onedrive(access_token, drive_id, item_id):
    api_url = f"{GRAPH_API_ENDPOINT}/drives/{drive_id}/items/{item_id}/content"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(api_url, headers=headers)
    response.raise_for_status()
    print("Excel descargado de OneDrive con éxito.")
    return response.content

# Sube (o sobrescribe) el contenido de un archivo Excel a OneDrive, con reintentos si está bloqueado.
def upload_excel_to_onedrive(access_token, drive_id, item_id, file_content):
    api_url = f"{GRAPH_API_ENDPOINT}/drives/{drive_id}/items/{item_id}/content"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }

    max_retries = 3
    retry_delay = 5 # segundos

    for attempt in range(max_retries):
        try:
            response = requests.put(api_url, headers=headers, data=file_content)
            response.raise_for_status()
            print("Excel subido a OneDrive con éxito.")
            return # Si tiene éxito, salimos de la función
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 423 and attempt < max_retries - 1:
                print(f"El archivo está bloqueado. Reintentando en {retry_delay} segundos... (Intento {attempt + 1}/{max_retries})")
                time.sleep(retry_delay)
            else:
                raise # Si no es un error 423 o es el último intento, relanzamos la excepción

    print("No se pudo subir el archivo después de varios intentos.")

# Carga un archivo JSON (payload) desde una ruta específica.
def cargar_payload(ruta_archivo):
    try:
        with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
            return json.load(archivo)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error al cargar '{ruta_archivo}': {e}")
        return None

# Realiza una llamada POST a una API con un payload JSON.
def llamar_api(url, payload):
    if not payload: return None
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error en la llamada a la API '{url}': {e}")
        return None

# Guarda los datos de respuesta de la API en un archivo JSON.
def guardar_respuesta(datos, nombre_archivo):
    try:
        with open(nombre_archivo, 'w', encoding='utf-8') as f:
            json.dump(datos, f, indent=4, ensure_ascii=False)
        print(f"Respuesta de la API guardada en '{nombre_archivo}'.")
    except Exception as e:
        print(f"Error al guardar el archivo '{nombre_archivo}': {e}")

# Actualiza una hoja de cálculo de Excel en memoria con los nuevos datos.
def actualizar_hoja_excel(workbook, datos_general, datos_teams, sheet_name, fila_inicio, columna_inicio):
    """Actualiza una hoja específica dentro de un objeto workbook de openpyxl."""
    try:
        sheet = workbook[sheet_name]
        puntos_generales_dict = {e['teamname']: e['points'] for e in datos_teams['answer']['teams']}
        ranking_general_list = datos_general['answer']['ranking']

        # Combina los datos de puntos totales y generales en una sola estructura.
        equipos_para_ordenar = []
        for equipo in ranking_general_list:
            nombre_equipo = equipo['name']
            equipos_para_ordenar.append({
                'name': nombre_equipo,
                'points': equipo['points'], # Puntos totales
                'general_points': puntos_generales_dict.get(nombre_equipo, 0) # Puntos generales
            })

        # Ordena la lista. Criterio primario: 'points' (descendente). Criterio secundario (desempate): 'general_points' (descendente).
        ranking_ordenado = sorted(equipos_para_ordenar, key=lambda x: (x['points'], x['general_points']), reverse=True)

        # Limpia el área de datos existente
        for row in sheet.iter_rows(min_row=fila_inicio, max_row=sheet.max_row, min_col=1, max_col=columna_inicio + 3):
            for cell in row:
                cell.value = None

        # Escribe los nuevos datos ya ordenados
        for i, equipo in enumerate(ranking_ordenado):
            fila_actual = fila_inicio + i

            # --- INICIO DEL CAMBIO ---
            # Lógica corregida para escribir el puesto en la columna correcta según la división.
            if sheet_name == "Clasificación 1a DIV":
                sheet.cell(row=fila_actual, column=1).value = f"{i+1}º"
            elif sheet_name == "Clasificación 2a DIV":
                sheet.cell(row=fila_actual, column=2).value = f"{i+1}º"
            # --- FIN DEL CAMBIO ---

            sheet.cell(row=fila_actual, column=columna_inicio).value = equipo['name']
            sheet.cell(row=fila_actual, column=columna_inicio + 1).value = equipo['points']
            sheet.cell(row=fila_actual, column=columna_inicio + 2).value = equipo['general_points']

        print(f"Hoja '{sheet_name}' actualizada en memoria.")
    except Exception as e:
        print(f"Error procesando la hoja '{sheet_name}' en memoria: {e}")

# Procesa y genera un informe JSON completo para una ronda específica.
def procesar_ronda_completa(payload_file, output_file, round_number, division):
    print(f"\n--- Procesando resultados de la ronda para la {division} división ---")
    API_URL_ROUND = "https://api.futmondo.com/1/ranking/round"
    API_URL_LINEUP = "https://api.futmondo.com/1/userteam/roundlineup"
    payload_base = cargar_payload(payload_file)
    if not payload_base: return

    payload_round = copy.deepcopy(payload_base)
    payload_round['query']['roundNumber'] = round_number
    if 'roundId' in payload_round['query']: payload_round['query'].pop('roundId')
    datos_ronda = llamar_api(API_URL_ROUND, payload_round)

    if not datos_ronda or 'answer' not in datos_ronda or 'matches' not in datos_ronda['answer']:
        print("Error: Respuesta de API de ronda inválida.")
        return

    teams_in_round = datos_ronda['answer'].get('ranking', [])
    matches = datos_ronda['answer']['matches']
    team_map_id = {i + 1: team['_id'] for i, team in enumerate(teams_in_round)}
    team_map_name = {i + 1: team['name'] for i, team in enumerate(teams_in_round)}

    resultados_finales, puntos_equipos_por_ronda, jugadores_ronda = [], [], []

    for match in matches:
        ids = [team_map_id.get(p) for p in match['p']]
        nombres = [team_map_name.get(p) for p in match['p']]

        # Lógica para manejar los dos formatos de respuesta de la API para las puntuaciones.
        if 'data' in match and 'partial' in match['data']:
            puntos = match['data']['partial']
        elif 'm' in match:
            puntos = match['m']
        else:
            puntos = [0, 0] # Si no hay datos, se asume 0.

        for i in range(2):
            puntos_equipos_por_ronda.append({"equipo": nombres[i], "puntos": puntos[i]})

        lineups, capitanes = [], []
        for i in range(2):
            payload_lineup = copy.deepcopy(payload_base)
            payload_lineup['query'].update({'round': round_number, 'userteamId': ids[i]})
            datos_lineup = llamar_api(API_URL_LINEUP, payload_lineup)
            lineup_players = datos_lineup.get('answer', {}).get('players', [])
            lineups.append(lineup_players)

            capitan = next((p['name'] for p in lineup_players if p.get('cpt')), "")
            capitanes.append(capitan)

            for player in lineup_players:
                jugadores_ronda.append({"nombre": player['name'], "puntos": player['points'], "equipo": nombres[i], "es_capitan": player.get('cpt')})

        jugadores_repetidos = [p['name'] for p in lineups[0] if p['name'] in {p2['name'] for p2 in lineups[1]}]

        resultados_finales.append({
            "Combate": f"{nombres[0]} vs {nombres[1]}",
            f"{nombres[0]}": {"Puntuacion": puntos[0], "Capitan": capitanes[0]},
            f"{nombres[1]}": {"Puntuacion": puntos[1], "Capitan": capitanes[1]},
            "Jugadores repetidos": jugadores_repetidos
        })

    peores_equipos = sorted(puntos_equipos_por_ronda, key=lambda x: x['puntos'])[:3]
    lista_peores_equipos = [{"posicion": i + 1, **equipo} for i, equipo in enumerate(peores_equipos)]

    def encontrar_peores(jugadores, key_filter=None):
        min_puntos = float('inf')
        peores_map = defaultdict(list)
        iterable = filter(key_filter, jugadores) if key_filter else jugadores
        for jugador in iterable:
            if jugador['puntos'] < min_puntos:
                min_puntos = jugador['puntos']
                peores_map.clear()
            if jugador['puntos'] == min_puntos:
                if jugador['equipo'] not in peores_map[jugador['nombre']]:
                    peores_map[jugador['nombre']].append(jugador['equipo'])
        return [{"nombre": n, "puntos": min_puntos, "equipos": e} for n, e in peores_map.items()]

    peores_capitanes_final = encontrar_peores(jugadores_ronda, lambda j: j['es_capitan'])
    peores_jugadores_final = encontrar_peores(jugadores_ronda)

    resumen_final = {
        "Resultados por combate": resultados_finales,
        "Peor Capitan": peores_capitanes_final,
        "Peor Jugador": peores_jugadores_final,
        "Los 3 peores equipos de la ronda": lista_peores_equipos
    }
    guardar_respuesta(resumen_final, output_file)

# Función principal que orquesta la ejecución del script.
def main():
    """Función principal que orquesta la ejecución del script."""
    LOCAL_EXCEL_FILENAME = "SuperLiga Fuentmondo 25-26.xlsx"

    modo = choose_save_option()
    if not modo:
        print("No se seleccionó ninguna opción. Finalizando el script.")
        return

    # --- 1. Obtener todos los datos de la API primero ---
    print("\n--- OBTENIENDO DATOS DE FUTMONDO ---")

    # Datos 1a División
    payload_1a = cargar_payload("payload_primera.json")
    if not payload_1a: return
    payload_general_1a = copy.deepcopy(payload_1a)
    if 'roundId' in payload_general_1a['query']: payload_general_1a['query'].pop('roundId', None)
    datos_general_1a = llamar_api("https://api.futmondo.com/1/ranking/general", payload_general_1a)
    payload_teams_1a = copy.deepcopy(payload_1a)
    payload_teams_1a['query'] = {"championshipId": payload_1a["query"]["championshipId"]}
    datos_teams_1a = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams_1a)

    # Datos 2a División
    payload_2a = cargar_payload("payload.json")
    if not payload_2a: return
    payload_general_2a = copy.deepcopy(payload_2a)
    if 'roundId' in payload_general_2a['query']: payload_general_2a['query'].pop('roundId', None)
    datos_general_2a = llamar_api("https://api.futmondo.com/1/ranking/general", payload_general_2a)
    payload_teams_2a = copy.deepcopy(payload_2a)
    payload_teams_2a['query'] = {"championshipId": payload_2a["query"]["championshipId"]}
    datos_teams_2a = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams_2a)

    if not all([datos_general_1a, datos_teams_1a, datos_general_2a, datos_teams_2a]):
        print("Faltan datos de la API. No se puede actualizar el Excel. Finalizando.")
        return

    # --- 2. Cargar el Workbook de Excel (local o desde OneDrive) ---
    print(f"\n--- ACTUALIZANDO CLASIFICACIONES (MODO: {modo}) ---")
    workbook = None
    access_token = None
    drive_id = None
    item_id = None

    if modo == 'local':
        try:
            workbook = openpyxl.load_workbook(LOCAL_EXCEL_FILENAME)
            print(f"Archivo '{LOCAL_EXCEL_FILENAME}' cargado localmente.")
        except FileNotFoundError:
            print(f"Error: No se encontró el archivo '{LOCAL_EXCEL_FILENAME}' en el directorio.")
            return
    else: # modo 'onedrive'
        access_token = get_access_token()
        if not access_token:
            print("No se pudo obtener el token de acceso. Finalizando el script.")
            return
        try:
            drive_id, item_id = get_drive_item_from_share_link(access_token, ONEDRIVE_SHARE_LINK)
            excel_content = download_excel_from_onedrive(access_token, drive_id, item_id)
            workbook = openpyxl.load_workbook(io.BytesIO(excel_content))
        except Exception as e:
            print(f"Error al descargar o cargar el archivo de OneDrive: {e}")
            return

    # --- 3. Actualizar ambas hojas en el objeto Workbook ---
    actualizar_hoja_excel(workbook, datos_general_1a, datos_teams_1a, "Clasificación 1a DIV", 5, 2)
    actualizar_hoja_excel(workbook, datos_general_2a, datos_teams_2a, "Clasificación 2a DIV", 2, 3)

    # --- 4. Guardar el Workbook modificado (localmente o en OneDrive) ---
    if modo == 'local':
        try:
            workbook.save(LOCAL_EXCEL_FILENAME)
            print(f"Archivo '{LOCAL_EXCEL_FILENAME}' guardado localmente con éxito.")
        except Exception as e:
            print(f"Error al guardar el archivo local '{LOCAL_EXCEL_FILENAME}': {e}")
    else: # modo 'onedrive'
        buffer = io.BytesIO()
        workbook.save(buffer)
        upload_excel_to_onedrive(access_token, drive_id, item_id, buffer.getvalue())

    # --- 5. Procesar análisis de rondas (esto no cambia) ---
    print("\n--- INICIANDO ANÁLISIS DE RONDAS ---")
    ROUND_NUMBER = "6868e3437775a433449ea12b"
    procesar_ronda_completa("payload_primera.json", "resultados_ronda_1a_div.json", ROUND_NUMBER, "primera")
    procesar_ronda_completa("payload.json", "resultados_ronda_2a_div.json", ROUND_NUMBER, "segunda")

    print("\n--- Proceso completado. ---")


if __name__ == '__main__':
    main()
