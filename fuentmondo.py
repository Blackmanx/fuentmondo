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
import pyperclip  # <--- LIBRERÍA AÑADIDA
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
        # --- INICIO DEL CAMBIO ---
        pyperclip.copy(user_code) # Usamos pyperclip, que es más fiable
        # --- FIN DEL CAMBIO ---
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

# --- [Aquí van el resto de funciones del script, que no han cambiado] ---
def encode_sharing_link(sharing_link):
    base64_value = base64.b64encode(sharing_link.encode('utf-8')).decode('utf-8')
    return 'u!' + base64_value.rstrip('=').replace('/', '_').replace('+', '-')

def get_drive_item_from_share_link(access_token, share_url):
    encoded_url = encode_sharing_link(share_url)
    api_url = f"{GRAPH_API_ENDPOINT}/shares/{encoded_url}/driveItem"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(api_url, headers=headers)
    response.raise_for_status()
    data = response.json()
    return data['parentReference']['driveId'], data['id']

def download_excel_from_onedrive(access_token, drive_id, item_id):
    api_url = f"{GRAPH_API_ENDPOINT}/drives/{drive_id}/items/{item_id}/content"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(api_url, headers=headers)
    response.raise_for_status()
    print("Excel descargado de OneDrive con éxito.")
    return response.content

def upload_excel_to_onedrive(access_token, drive_id, item_id, file_content):
    api_url = f"{GRAPH_API_ENDPOINT}/drives/{drive_id}/items/{item_id}/content"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    response = requests.put(api_url, headers=headers, data=file_content)
    response.raise_for_status()
    print("Excel subido a OneDrive con éxito.")

def cargar_payload(ruta_archivo):
    try:
        with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
            return json.load(archivo)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error al cargar '{ruta_archivo}': {e}")
        return None

def llamar_api(url, payload):
    if not payload: return None
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error en la llamada a la API '{url}': {e}")
        return None

def guardar_respuesta(datos, nombre_archivo):
    try:
        with open(nombre_archivo, 'w', encoding='utf-8') as f:
            json.dump(datos, f, indent=4, ensure_ascii=False)
        print(f"Respuesta de la API guardada en '{nombre_archivo}'.")
    except Exception as e:
        print(f"Error al guardar el archivo '{nombre_archivo}': {e}")

def actualizar_excel_en_memoria(file_content, datos_general, datos_teams, sheet_name, fila_inicio, columna_inicio):
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(file_content))
        sheet = workbook[sheet_name]
        puntos_generales = {e['teamname']: e['points'] for e in datos_teams['answer']['teams']}
        ranking_general = datos_general['answer']['ranking']

        for row in sheet.iter_rows(min_row=fila_inicio, max_row=sheet.max_row, min_col=1, max_col=columna_inicio + 3):
            for cell in row:
                cell.value = None

        for i, equipo in enumerate(ranking_general):
            fila_actual = fila_inicio + i
            if columna_inicio == 2: # Para 2a DIV
                sheet.cell(row=fila_actual, column=1).value = f"{i+1}º"
            sheet.cell(row=fila_actual, column=columna_inicio).value = equipo['name']
            sheet.cell(row=fila_actual, column=columna_inicio + 1).value = equipo['points']
            sheet.cell(row=fila_actual, column=columna_inicio + 2).value = puntos_generales.get(equipo['name'], 0)

        buffer = io.BytesIO()
        workbook.save(buffer)
        print(f"Hoja '{sheet_name}' actualizada en memoria.")
        return buffer.getvalue()
    except Exception as e:
        print(f"Error procesando el Excel en memoria: {e}")
        return None

def procesar_y_actualizar_division(access_token, payload_file, sheet_name, start_row, start_col):
    try:
        drive_id, item_id = get_drive_item_from_share_link(access_token, ONEDRIVE_SHARE_LINK)
        excel_content = download_excel_from_onedrive(access_token, drive_id, item_id)

        payload = cargar_payload(payload_file)
        if not payload: return

        payload_general = copy.deepcopy(payload)
        if 'roundId' in payload_general['query']: payload_general['query'].pop('roundId', None)
        datos_general = llamar_api("https://api.futmondo.com/1/ranking/general", payload_general)

        payload_teams = copy.deepcopy(payload)
        payload_teams['query'] = {"championshipId": payload["query"]["championshipId"]}
        datos_teams = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams)

        if datos_general and datos_teams:
            updated_content = actualizar_excel_en_memoria(excel_content, datos_general, datos_teams, sheet_name, start_row, start_col)
            if updated_content:
                upload_excel_to_onedrive(access_token, drive_id, item_id, updated_content)
        else:
            print(f"No se obtuvieron los datos de Futmondo para '{sheet_name}'.")

    except Exception as e:
        print(f"Ha ocurrido un error durante el proceso de la división '{sheet_name}': {e}")

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
        puntos = match['data']['partial']

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

def main():
    access_token = get_access_token()
    if not access_token:
        print("No se pudo obtener el token de acceso. Finalizando el script.")
        return

    print("\n--- INICIANDO ACTUALIZACIÓN DE CLASIFICACIONES EN ONEDRIVE ---")
    procesar_y_actualizar_division(access_token, "payload_primera.json", "Clasificación 1a DIV", 5, 2)
    procesar_y_actualizar_division(access_token, "payload.json", "Clasificación 2a DIV", 2, 3)

    print("\n--- INICIANDO ANÁLISIS DE RONDAS ---")
    ROUND_NUMBER = "6868e3437775a433449ea12b"
    procesar_ronda_completa("payload_primera.json", "resultados_ronda_1a_div.json", ROUND_NUMBER, "primera")
    procesar_ronda_completa("payload.json", "resultados_ronda_2a_div.json", ROUND_NUMBER, "segunda")

if __name__ == '__main__':
    main()
