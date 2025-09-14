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

# --- Funciones de Interfaz Gráfica (Tkinter) ---

# Crea una ventana con botones para que el usuario elija la ubicación de guardado.
def choose_save_option():
    root = tk.Tk()
    root.title("Opción de Guardado")
    choice = [None]
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

# Crea una ventana con el código de autenticación para que el usuario lo copie.
def show_auth_code_window(message, verification_uri):
    root = tk.Tk()
    root.title("Código de Autenticación")
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

# --- Funciones de Autenticación y OneDrive ---

# Se autentica de forma interactiva y obtiene un token de acceso para Microsoft Graph.
def get_access_token():
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
    retry_delay = 5
    for attempt in range(max_retries):
        try:
            response = requests.put(api_url, headers=headers, data=file_content)
            response.raise_for_status()
            print("Excel subido a OneDrive con éxito.")
            return
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 423 and attempt < max_retries - 1:
                print(f"El archivo está bloqueado. Reintentando en {retry_delay} segundos... (Intento {attempt + 1}/{max_retries})")
                time.sleep(retry_delay)
            else:
                raise
    print("No se pudo subir el archivo después de varios intentos.")

# --- Funciones de API de Futmondo y Lógica de Datos ---

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

# Obtiene y muestra la alineación y el capitán de un equipo específico para una ronda.
def obtener_y_mostrar_alineacion(payload_base, team_id, round_id, team_name):
    print(f"\n--- OBTENIENDO ALINEACIÓN PARA: {team_name} ---")
    API_URL_LINEUP = "https://api.futmondo.com/1/userteam/roundlineup"
    payload_lineup = copy.deepcopy(payload_base)
    payload_lineup['query'].update({'round': round_id, 'userteamId': team_id})
    datos_lineup = llamar_api(API_URL_LINEUP, payload_lineup)
    if datos_lineup and 'answer' in datos_lineup and 'players' in datos_lineup['answer']:
        lineup_players = datos_lineup['answer']['players']
        capitan = "No asignado"
        print(f"Jugadores de {team_name}:")
        for player in lineup_players:
            es_capitan = ""
            if player.get('cpt'):
                capitan = player['name']
                es_capitan = " (Capitán)"
            print(f"- {player['name']}{es_capitan}")
        print(f"\nCapitán seleccionado: {capitan}")
    else:
        print(f"No se pudo obtener la alineación para {team_name}.")
        if datos_lineup:
            print("Respuesta de la API:", datos_lineup)

# Obtiene una lista de los capitanes de todos los equipos para una ronda específica.
def get_captains_for_round(payload_base, datos_ronda, name_map={}):
    API_URL_LINEUP = "https://api.futmondo.com/1/userteam/roundlineup"
    team_captains = []
    if 'answer' not in datos_ronda or 'ranking' not in datos_ronda['answer']:
        return []
    ranking = datos_ronda['answer']['ranking']
    for team_info in ranking:
        team_id = team_info['_id']
        team_name_api = team_info['name']
        payload_lineup = {
            "header": copy.deepcopy(payload_base["header"]),
            "query": {
                "championshipId": payload_base["query"]["championshipId"],
                "round": datos_ronda['query']['roundNumber'],
                "userteamId": team_id
            }
        }
        datos_lineup = llamar_api(API_URL_LINEUP, payload_lineup)
        if datos_lineup and 'answer' in datos_lineup and 'players' in datos_lineup['answer']:
            lineup_players = datos_lineup['answer']['players']
            capitan = next((p['name'] for p in lineup_players if p.get('cpt')), "N/A")
            canonical_name = name_map.get(team_name_api, team_name_api)
            team_captains.append({"team_name": canonical_name, "capitan": capitan})
    return team_captains

# --- [NUEVA FUNCIÓN] ---
# Procesa la respuesta de la API de rondas, manejando casos especiales como la jornada 1.5.
def procesar_rondas_api(rounds_list):
    if not rounds_list:
        return {}

    rounds_map = {}
    for r in rounds_list:
        round_num = r.get('number')
        round_id = r.get('id')

        if not round_num or not round_id:
            continue

        # Regla especial: la jornada 1.5 se trata como la jornada 6
        if round_num == 1.5:
            rounds_map[6] = round_id
            print("Jornada especial 1.5 mapeada como Jornada 6.")
        # Solo procesar números enteros para el resto
        elif round_num % 1 == 0:
            rounds_map[int(round_num)] = round_id

    return rounds_map

# --- Funciones de Actualización de Excel ---

# Actualiza una hoja de clasificación dentro de un objeto workbook de openpyxl.
def actualizar_hoja_excel(workbook, datos_general, datos_teams, sheet_name, fila_inicio, columna_inicio, name_map={}):
    try:
        sheet = workbook[sheet_name]
        puntos_generales_dict = {name_map.get(e['teamname'], e['teamname']): e['points'] for e in datos_teams['answer']['teams']}
        ranking_general_list = datos_general['answer']['ranking']
        equipos_para_ordenar = []
        for equipo in ranking_general_list:
            nombre_equipo_api = equipo['name']
            nombre_equipo_canonico = name_map.get(nombre_equipo_api, nombre_equipo_api)
            equipos_para_ordenar.append({
                'name': nombre_equipo_canonico,
                'points': equipo['points'],
                'general_points': puntos_generales_dict.get(nombre_equipo_canonico, 0)
            })
        ranking_ordenado = sorted(equipos_para_ordenar, key=lambda x: (x['points'], x['general_points']), reverse=True)
        for row in sheet.iter_rows(min_row=fila_inicio, max_row=sheet.max_row, min_col=columna_inicio, max_col=columna_inicio + 2):
            for cell in row:
                cell.value = None
        for i, equipo in enumerate(ranking_ordenado):
            fila_actual = fila_inicio + i
            sheet.cell(row=fila_actual, column=columna_inicio).value = equipo['name']
            sheet.cell(row=fila_actual, column=columna_inicio + 1).value = equipo['points']
            sheet.cell(row=fila_actual, column=columna_inicio + 2).value = equipo['general_points']
        print(f"Hoja '{sheet_name}' actualizada en memoria.")
    except Exception as e:
        print(f"Error procesando la hoja '{sheet_name}' en memoria: {e}")

# Sobrescribe las cabeceras de la hoja 'Capitanes' con los nombres canónicos.
def actualizar_cabeceras_capitanes(workbook, teams_dict_1a, teams_dict_2a):
    try:
        sheet = workbook["Capitanes"]
        print("Actualizando cabeceras de 1ª División en la hoja 'Capitanes'...")
        sorted_teams_1a = sorted(teams_dict_1a.items(), key=lambda item: int(item[0]))
        current_col = 3
        for _, team_name in sorted_teams_1a:
            sheet.cell(row=3, column=current_col).value = team_name
            current_col += 2
        print("Actualizando cabeceras de 2ª División en la hoja 'Capitanes'...")
        sorted_teams_2a = sorted(teams_dict_2a.items(), key=lambda item: int(item[0]))
        current_col = 44
        for _, team_name in sorted_teams_2a:
            sheet.cell(row=3, column=current_col).value = team_name
            current_col += 2
        print("Cabeceras de la hoja 'Capitanes' actualizadas.")
    except Exception as e:
        print(f"Error al actualizar las cabeceras de la hoja 'Capitanes': {e}")

# Actualiza la hoja de capitanes para una jornada específica.
def actualizar_hoja_capitanes(workbook, round_number, team_captains_list):
    try:
        sheet = workbook["Capitanes"]
        team_to_captain_col = {}
        for col_idx in range(3, 150):
            team_name_cell = sheet.cell(row=3, column=col_idx)
            if team_name_cell.value:
                team_name = str(team_name_cell.value).strip()
                team_to_captain_col[team_name] = col_idx
        row_found = False
        target_row_label = f"Jornada {round_number}"
        for row_idx, row in enumerate(sheet.iter_rows(min_row=5, max_row=sheet.max_row, min_col=2, max_col=2), 5):
            cell_value = row[0].value
            if isinstance(cell_value, str) and cell_value.strip().lower() == target_row_label.lower():
                for captain_info in team_captains_list:
                    team_name = captain_info['team_name'].strip()
                    captain = captain_info['capitan']
                    if team_name in team_to_captain_col:
                        col_to_update = team_to_captain_col[team_name]
                        sheet.cell(row=row_idx, column=col_to_update).value = captain
                    else:
                        print(f"  -> Aviso: El equipo '{team_name}' no se encontró en la cabecera de la hoja 'Capitanes'.")
                print(f"Capitanes de la '{target_row_label}' actualizados en memoria.")
                row_found = True
                break
        if not row_found:
            print(f"Advertencia: No se encontró la fila para '{target_row_label}' en la hoja 'Capitanes'.")
    except Exception as e:
        print(f"Error actualizando la hoja 'Capitanes' para la jornada {round_number}: {e}")

# Itera sobre todas las rondas de un campeonato y actualiza la hoja de capitanes.
def actualizar_capitanes_historico(workbook, rounds_map, payload_base, division_name, name_map={}):
    print(f"\n--- INICIANDO ACTUALIZACIÓN HISTÓRICA DE CAPITANES PARA {division_name.upper()} ---")
    sorted_round_numbers = sorted(rounds_map.keys())
    for round_number in sorted_round_numbers:
        round_id = rounds_map[round_number]
        print(f"Procesando capitanes de la Jornada {round_number}...")
        payload_round = copy.deepcopy(payload_base)
        payload_round['query'].update({'roundNumber': round_id, 'championshipId': payload_base['query']['championshipId']})
        datos_ronda = llamar_api("https://api.futmondo.com/1/ranking/round", payload_round)
        if not datos_ronda or 'answer' not in datos_ronda or datos_ronda['answer'] == 'api.error.general':
            print(f"  -> Error: No se pudieron obtener datos para la Jornada {round_number}. Saltando.")
            continue
        datos_ronda['query']['roundNumber'] = round_id
        team_captains = get_captains_for_round(payload_base, datos_ronda, name_map)
        if not team_captains:
            print(f"  -> Advertencia: No se encontraron capitanes para la Jornada {round_number}.")
            continue
        actualizar_hoja_capitanes(workbook, round_number, team_captains)

# Procesa y genera un informe JSON completo para una ronda específica.
def procesar_ronda_completa(datos_ronda, output_file, payload_base, name_map={}):
    print(f"\n--- Procesando resultados de la ronda para la {datos_ronda.get('division', 'división desconocida')} ---")
    API_URL_LINEUP = "https://api.futmondo.com/1/userteam/roundlineup"
    if not datos_ronda or 'answer' not in datos_ronda or 'matches' not in datos_ronda['answer']:
        print("Error: Respuesta de API de ronda inválida.")
        return
    teams_in_round_list = datos_ronda['answer'].get('ranking', [])
    matches = datos_ronda['answer']['matches']
    round_id_actual = datos_ronda['query']['roundId']
    team_map_id = {i + 1: team['_id'] for i, team in enumerate(teams_in_round_list)}
    team_map_name = {i + 1: name_map.get(team['name'], team['name']) for i, team in enumerate(teams_in_round_list)}
    resultados_finales, puntos_equipos_por_ronda, jugadores_ronda = [], [], []
    for match in matches:
        ids = [team_map_id.get(p) for p in match['p']]
        nombres = [team_map_name.get(p) for p in match['p']]
        if 'data' in match and 'partial' in match['data']:
            puntos = match['data']['partial']
        elif 'm' in match:
            puntos = match['m']
        else:
            puntos = [0, 0]
        for i in range(2):
            puntos_equipos_por_ronda.append({"equipo": nombres[i], "puntos": puntos[i]})
        lineups, capitanes = [], []
        for i in range(2):
            payload_lineup = {
                "header": copy.deepcopy(payload_base["header"]),
                "query": {
                    "championshipId": payload_base["query"]["championshipId"],
                    "round": round_id_actual,
                    "userteamId": ids[i]
                }
            }
            datos_lineup = llamar_api(API_URL_LINEUP, payload_lineup)
            lineup_players = datos_lineup.get('answer', {}).get('players', [])
            lineups.append(lineup_players)
            capitan = next((p['name'] for p in lineup_players if p.get('cpt')), "N/A")
            capitanes.append(capitan)
            for player in lineup_players:
                jugadores_ronda.append({
                    "nombre": player['name'],
                    "puntos": player['points'],
                    "equipo": nombres[i],
                    "es_capitan": player.get('cpt', False) # Asegurarse de que el valor sea booleano
                })
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
            puntos_jugador = jugador.get('puntos', 0)
            if puntos_jugador < min_puntos:
                min_puntos = puntos_jugador
                peores_map.clear()
            if puntos_jugador == min_puntos:
                if jugador['equipo'] not in peores_map[jugador['nombre']]:
                    peores_map[jugador['nombre']].append(jugador['equipo'])
        return [{"nombre": n, "puntos": min_puntos, "equipos": e} for n, e in peores_map.items()]
    peores_capitanes_final = encontrar_peores(jugadores_ronda, lambda j: j.get('es_capitan'))
    peores_jugadores_final = encontrar_peores(jugadores_ronda)
    resumen_final = {
        "Resultados por combate": resultados_finales,
        "Peor Capitan": peores_capitanes_final,
        "Peor Jugador": peores_jugadores_final,
        "Los 3 peores equipos de la ronda": lista_peores_equipos
    }
    guardar_respuesta(resumen_final, output_file)

# --- Función Principal ---
def main():
    LOCAL_EXCEL_FILENAME = "SuperLiga Fuentmondo 25-26.xlsx"
    TEAMS_1A = {
        "1":"Galácticos de la noche FC", "2":"AL-CARRER F.C.", "3":"QUE BARBARIDAD FC",
        "4":"Fuentino Pérez", "5":"CALAMARES CON TORRIJAS🦑🍞", "6":"CD Congelados",
        "7":"THE LIONS", "8":"EL CHOLISMO FC", "9":"Real Fermín C.F.",
        "10":"Real 🥚🥚 Bailarines 🪩F.C", "11":"MORRITOS F.C.", "12":"Poli Ejido CF",
        "13":"Juaki la bomba", "14":"LA MARRANERA", "15":"Larios Limon FC",
        "16":"PANAKOTA F.F.", "17":"Real Pezqueñines FC", "18":"LOS POKÉMON 🐭🟡🐭",
        "19":"El Huracán CF", "20":"Lim Hijo de Puta"
    }
    TEAMS_2A = {
      "1":"SANTA LUCIA FC", "2":"Osasuna N.S.R", "3":"Tetitas Colesterol . F.C",
      "4":"Pollos sin cabeza 🐥🧄", "5":"Charo la  Picanta FC", "6":"Kostas Mariotas",
      "7":"Real Pescados el Puerto Fc", "8":"Team pepino", "9":"🇧🇷Samba Rovinha 🇧🇷",
      "10":"Banano Vallekano 🍌⚡", "11":"SICARIOS CF", "12":"Minabo De Kiev",
      "13":"Todo por la camiseta 🇪🇸", "14":"parker f.c.", "15":"Molinardo fc",
      "16":"Lazaroneta", "17":"ElBarto F.C", "18":"BANANEROS FC",
      "19":"Morenetes de la Giralda 🍩", "20":"Jamon York F.C.", "21":"Elche pero Peor",
      "22":"Motobetis a primera!", "23":"MTB Drink Team", "24":"Patejas"
    }
    modo = choose_save_option()
    if not modo:
        print("No se seleccionó ninguna opción. Finalizando el script.")
        return

    print("\n--- OBTENIENDO DATOS DE FUTMONDO ---")
    payload_1a = cargar_payload("payload_primera.json")
    if not payload_1a: return
    payload_2a = cargar_payload("payload.json")
    if not payload_2a: return

    # Obtenemos primero la lista de todas las jornadas disponibles
    rounds_data_1a = llamar_api("https://api.futmondo.com/1/userteam/rounds", copy.deepcopy(payload_1a))
    rounds_data_2a = llamar_api("https://api.futmondo.com/1/userteam/rounds", copy.deepcopy(payload_2a))

    rounds_map_1a = procesar_rondas_api(rounds_data_1a.get('answer', []))
    rounds_map_2a = procesar_rondas_api(rounds_data_2a.get('answer', []))

    if not rounds_map_1a or not rounds_map_2a:
        print("Error: No se pudo obtener y procesar la lista de rondas de la API. Finalizando.")
        return

    # Obtenemos los datos de la última jornada para cada división
    latest_round_id_1a = rounds_map_1a[max(rounds_map_1a.keys())]
    payload_round_1a = copy.deepcopy(payload_1a)
    payload_round_1a['query'].update({'roundNumber': latest_round_id_1a})
    datos_ronda_1a = llamar_api("https://api.futmondo.com/1/ranking/round", payload_round_1a)

    latest_round_id_2a = rounds_map_2a[max(rounds_map_2a.keys())]
    payload_round_2a = copy.deepcopy(payload_2a)
    payload_round_2a['query'].update({'roundNumber': latest_round_id_2a})
    datos_ronda_2a = llamar_api("https://api.futmondo.com/1/ranking/round", payload_round_2a)

    # --- [CORRECCIÓN] ---
    # Creamos el mapa de nombres usando la lista de equipos de 'datos_ronda'
    # que tiene el orden correcto y fijo, en lugar de 'datos_general'.
    map_1a, map_2a = {}, {}
    if datos_ronda_1a and 'answer' in datos_ronda_1a and 'ranking' in datos_ronda_1a['answer']:
        round_ranking_1a = datos_ronda_1a['answer']['ranking']
        if len(round_ranking_1a) >= len(TEAMS_1A):
            map_1a = {round_ranking_1a[i]['name']: TEAMS_1A[str(i + 1)] for i in range(len(TEAMS_1A))}
            print("Mapeo de nombres para 1a División creado con éxito.")
        else:
            print(f"ADVERTENCIA: No se pudo crear el mapeo para 1a División. La API de ronda devolvió {len(round_ranking_1a)} equipos y se proporcionaron {len(TEAMS_1A)}.")

    if datos_ronda_2a and 'answer' in datos_ronda_2a and 'ranking' in datos_ronda_2a['answer']:
        round_ranking_2a = datos_ronda_2a['answer']['ranking']
        if len(round_ranking_2a) >= len(TEAMS_2A):
            map_2a = {round_ranking_2a[i]['name']: TEAMS_2A[str(i + 1)] for i in range(len(TEAMS_2A))}
            print("Mapeo de nombres para 2a División creado con éxito.")
        else:
            print(f"ADVERTENCIA: No se pudo crear el mapeo para 2a División. La API de ronda devolvió {len(round_ranking_2a)} equipos y se proporcionaron {len(TEAMS_2A)}.")

    # Obtenemos los datos restantes que sí dependen de la clasificación general
    datos_general_1a = llamar_api("https://api.futmondo.com/1/ranking/general", copy.deepcopy(payload_1a))
    datos_general_2a = llamar_api("https://api.futmondo.com/1/ranking/general", copy.deepcopy(payload_2a))

    payload_teams_1a = copy.deepcopy(payload_1a)
    payload_teams_1a['query'] = {"championshipId": payload_1a["query"]["championshipId"]}
    datos_teams_1a = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams_1a)

    payload_teams_2a = copy.deepcopy(payload_2a)
    payload_teams_2a['query'] = {"championshipId": payload_2a["query"]["championshipId"]}
    datos_teams_2a = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams_2a)

    if not all([datos_general_1a, datos_teams_1a, datos_general_2a, datos_teams_2a, datos_ronda_1a, datos_ronda_2a]):
        print("Faltan datos clave de la API. No se puede actualizar el Excel. Finalizando.")
        return

    print(f"\n--- CARGANDO EXCEL (MODO: {modo}) ---")
    workbook, access_token, drive_id, item_id = None, None, None, None
    try:
        if modo == 'local':
            workbook = openpyxl.load_workbook(LOCAL_EXCEL_FILENAME)
            print(f"Archivo '{LOCAL_EXCEL_FILENAME}' cargado localmente.")
        else:
            access_token = get_access_token()
            if not access_token: raise Exception("No se pudo obtener el token de acceso.")
            drive_id, item_id = get_drive_item_from_share_link(access_token, ONEDRIVE_SHARE_LINK)
            excel_content = download_excel_from_onedrive(access_token, drive_id, item_id)
            workbook = openpyxl.load_workbook(io.BytesIO(excel_content))
    except Exception as e:
        print(f"Error fatal al cargar el archivo Excel: {e}")
        return

    actualizar_cabeceras_capitanes(workbook, TEAMS_1A, TEAMS_2A)
    actualizar_hoja_excel(workbook, datos_general_1a, datos_teams_1a, "Clasificación 1a DIV", 5, 2, map_1a)
    actualizar_hoja_excel(workbook, datos_general_2a, datos_teams_2a, "Clasificación 2a DIV", 2, 3, map_2a)
    actualizar_capitanes_historico(workbook, rounds_map_1a, payload_1a, "1a División", map_1a)
    actualizar_capitanes_historico(workbook, rounds_map_2a, payload_2a, "2a División", map_2a)

    try:
        if modo == 'local':
            workbook.save(LOCAL_EXCEL_FILENAME)
            print(f"\nArchivo '{LOCAL_EXCEL_FILENAME}' guardado localmente con éxito.")
        else:
            buffer = io.BytesIO()
            workbook.save(buffer)
            upload_excel_to_onedrive(access_token, drive_id, item_id, buffer.getvalue())
    except Exception as e:
        print(f"\nError al guardar el archivo Excel: {e}")

    print("\n--- INICIANDO ANÁLISIS DE LA ÚLTIMA RONDA PARA INFORMES JSON ---")
    datos_ronda_1a['division'] = 'primera'
    datos_ronda_1a['query']['roundId'] = latest_round_id_1a
    datos_ronda_2a['division'] = 'segunda'
    datos_ronda_2a['query']['roundId'] = latest_round_id_2a

    procesar_ronda_completa(datos_ronda_1a, "resultados_ronda_1a_div.json", payload_1a, map_1a)
    procesar_ronda_completa(datos_ronda_2a, "resultados_ronda_2a_div.json", payload_2a, map_2a)

    team_id_motobetis = "66b9c5d20edfaa6140f45f75"
    if rounds_map_2a:
        latest_round_id_2a_debug = rounds_map_2a[max(rounds_map_2a.keys())]
        obtener_y_mostrar_alineacion(payload_2a, team_id_motobetis, latest_round_id_2a_debug, "Motobetis a primera!")
    else:
        print("No se pudo ejecutar la depuración para Motobetis: el mapa de rondas está vacío.")

    print("\n--- Proceso completado. ---")

if __name__ == '__main__':
    main()
