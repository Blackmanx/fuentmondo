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

# Crea una ventana para que el usuario elija el modo de ejecución.
def choose_save_option():
    root = tk.Tk()
    root.title("Modo de Ejecución")
    choice = [None]
    window_width = 550
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

    tk.Label(root, text="Elige qué quieres hacer:", pady=15, font=("Helvetica", 12)).pack()
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    btn_onedrive = tk.Button(button_frame, text="Actualizar Excel en OneDrive", command=lambda: select_option('onedrive'), height=2, width=25, bg="#0078D4", fg="white")
    btn_onedrive.pack(side=tk.LEFT, padx=5)

    btn_local = tk.Button(button_frame, text="Actualizar Excel Localmente", command=lambda: select_option('local'), height=2, width=25)
    btn_local.pack(side=tk.LEFT, padx=5)

    # --- [NUEVO BOTÓN] ---
    btn_multas = tk.Button(button_frame, text="Generar Solo Multas (HTML)", command=lambda: select_option('multas_only'), height=2, width=25, bg="#28a745", fg="white")
    btn_multas.pack(side=tk.LEFT, padx=5)

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
        os.makedirs(os.path.dirname(nombre_archivo), exist_ok=True)
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
        if round_num == 1.5:
            rounds_map[6] = round_id
            print("Jornada especial 1.5 mapeada como Jornada 6.")
        elif round_num % 1 == 0:
            rounds_map[int(round_num)] = round_id
    return rounds_map

# Calcula las multas de una jornada con un desglose detallado.
def calcular_multas_jornada(
    teams_in_round, matches, team_map_name, dict_alineaciones, dict_capitanes,
    lista_peores_equipos, peores_jugadores_final, peores_capitanes_final
):
    multas_finales = {}
    for team_name in teams_in_round:
        multas_finales[team_name] = {
            "multa_total": 0.0,
            "desglose": {
                "jugadores_repetidos": {"cantidad": 0, "multa": 0.0},
                "capitan_repetido_con_rival": {"aplicado": False, "multa": 0.0},
                "tenias_capitan_rival": {"aplicado": False, "multa": 0.0},
                "peor_equipo_jornada": {"posicion": 0, "multa": 0.0},
                "alinear_peor_jugador": {"aplicado": False, "multa": 0.0},
                "elegir_peor_capitan": {"aplicado": False, "multa": 0.0}
            }
        }
    for match in matches:
        team_indices = match['p']
        team_a_name = team_map_name.get(team_indices[0])
        team_b_name = team_map_name.get(team_indices[1])
        if not team_a_name or not team_b_name:
            continue
        lineup_a = dict_alineaciones.get(team_a_name, [])
        lineup_b = dict_alineaciones.get(team_b_name, [])
        nombres_a = {p['name'] for p in lineup_a}
        nombres_b = {p['name'] for p in lineup_b}
        capitan_a = dict_capitanes.get(team_a_name, "N/A")
        capitan_b = dict_capitanes.get(team_b_name, "N/A")
        repetidos = nombres_a.intersection(nombres_b)
        if repetidos:
            multa_repetidos = len(repetidos) * 0.5
            multas_finales[team_a_name]["desglose"]["jugadores_repetidos"] = {"cantidad": len(repetidos), "multa": multa_repetidos}
            multas_finales[team_b_name]["desglose"]["jugadores_repetidos"] = {"cantidad": len(repetidos), "multa": multa_repetidos}
        if capitan_a == capitan_b and capitan_a != "N/A":
            multas_finales[team_a_name]["desglose"]["capitan_repetido_con_rival"] = {"aplicado": True, "multa": 1.0}
            multas_finales[team_b_name]["desglose"]["capitan_repetido_con_rival"] = {"aplicado": True, "multa": 1.0}
        if capitan_a in nombres_b and capitan_a != capitan_b:
            multas_finales[team_b_name]["desglose"]["tenias_capitan_rival"] = {"aplicado": True, "multa": 1.0}
        if capitan_b in nombres_a and capitan_a != capitan_b:
            multas_finales[team_a_name]["desglose"]["tenias_capitan_rival"] = {"aplicado": True, "multa": 1.0}
    multas_peores = {1: 2.0, 2: 1.5, 3: 1.0}
    for item in lista_peores_equipos:
        pos = item['posicion']
        team_name = item['equipo']
        if team_name in multas_finales and pos in multas_peores:
            multas_finales[team_name]["desglose"]["peor_equipo_jornada"] = {"posicion": pos, "multa": multas_peores[pos]}
    nombres_peores_jugadores = {p['nombre'] for p in peores_jugadores_final}
    for team_name, alineacion in dict_alineaciones.items():
        nombres_jugadores = {p['name'] for p in alineacion}
        if not nombres_peores_jugadores.isdisjoint(nombres_jugadores):
            multas_finales[team_name]["desglose"]["alinear_peor_jugador"] = {"aplicado": True, "multa": 1.0}
    nombres_peores_capitanes = {p['nombre'] for p in peores_capitanes_final}
    for team_name, capitan in dict_capitanes.items():
        if capitan in nombres_peores_capitanes:
            multas_finales[team_name]["desglose"]["elegir_peor_capitan"] = {"aplicado": True, "multa": 1.0}
    for team_name, data in multas_finales.items():
        total = sum(d.get('multa', 0.0) for d in data['desglose'].values())
        multas_finales[team_name]['multa_total'] = round(total, 2)
    return multas_finales

# Genera el contenido HTML para una única tabla de multas de jornada.
def _generar_tabla_multas_jornada_html(multas_data):
    sorted_teams = sorted(multas_data.items(), key=lambda item: item[1]['multa_total'], reverse=True)
    table_rows = ""
    for team_name, data in sorted_teams:
        multa_total = data.get('multa_total', 0.0)
        if multa_total == 0: continue
        desglose = data.get('desglose', {})
        desglose_html = "<ul>"
        jr = desglose.get("jugadores_repetidos", {})
        if jr.get("multa", 0) > 0:
            desglose_html += f"<li>Jugadores repetidos ({jr.get('cantidad', 0)}): {jr.get('multa', 0):.2f}€</li>"
        cr = desglose.get("capitan_repetido_con_rival", {})
        if cr.get("multa", 0) > 0:
            desglose_html += f"<li>Capitán repetido con rival: {cr.get('multa', 0):.2f}€</li>"
        tcr = desglose.get("tenias_capitan_rival", {})
        if tcr.get("multa", 0) > 0:
            desglose_html += f"<li>Alinear al capitán del rival: {tcr.get('multa', 0):.2f}€</li>"
        pe = desglose.get("peor_equipo_jornada", {})
        if pe.get("multa", 0) > 0:
            pos_map = {1: "Peor", 2: "2º Peor", 3: "3er Peor"}
            pos_str = pos_map.get(pe.get("posicion"), f"{pe.get('posicion')}º Peor")
            desglose_html += f"<li>{pos_str} equipo de la jornada: {pe.get('multa', 0):.2f}€</li>"
        apj = desglose.get("alinear_peor_jugador", {})
        if apj.get("multa", 0) > 0:
            desglose_html += f"<li>Alinear al peor jugador: {apj.get('multa', 0):.2f}€</li>"
        epc = desglose.get("elegir_peor_capitan", {})
        if epc.get("multa", 0) > 0:
            desglose_html += f"<li>Elegir al peor capitán: {epc.get('multa', 0):.2f}€</li>"
        desglose_html += "</ul>"
        table_rows += f"""
        <tr>
            <td>{team_name}</td>
            <td class="total-multa">{multa_total:.2f}€</td>
            <td class="desglose">{desglose_html}</td>
        </tr>"""
    if not table_rows:
        table_rows = '<tr><td colspan="3" style="text-align:center;">No se registraron multas en esta jornada.</td></tr>'
    return f"""
    <table>
        <thead>
            <tr>
                <th>Equipo</th>
                <th>Multa Total</th>
                <th>Desglose</th>
            </tr>
        </thead>
        <tbody>{table_rows}</tbody>
    </table>"""

# Genera el contenido HTML para la tabla de multas totales.
def _generar_tabla_multas_totales_html(multas_acumuladas):
    sorted_teams = sorted(multas_acumuladas.items(), key=lambda item: item[1], reverse=True)
    table_rows = ""
    for team_name, total_multa in sorted_teams:
        table_rows += f"""
        <tr>
            <td>{team_name}</td>
            <td class="total-multa">{total_multa:.2f}€</td>
        </tr>"""
    return f"""
    <table>
        <thead>
            <tr>
                <th>Equipo</th>
                <th>Total Acumulado</th>
            </tr>
        </thead>
        <tbody>{table_rows}</tbody>
    </table>"""

# Genera la página HTML completa con todos los datos y la navegación.
def generar_pagina_html_completa(datos_informe, output_path):
    contenido_html = ""
    nav_links_html = ""

    for div_key, div_data in datos_informe.items():
        div_titulo = "1ª División" if div_key == "primera" else "2ª División"
        id_totales = f"{div_key}-totales"
        contenido_html += f"""
        <div id="{id_totales}" class="content-section">
            <h2>Multas Totales - {div_titulo}</h2>
            {_generar_tabla_multas_totales_html(div_data['totales'])}
        </div>"""
        nav_links_html += f'<a href="#" class="nav-link" data-target="{id_totales}">Totales {div_titulo}</a>'

        nav_links_html += '<div class="dropdown">'
        nav_links_html += f'<button class="dropbtn">Jornadas {div_titulo} &#9662;</button>'
        nav_links_html += '<div class="dropdown-content">'
        for jornada_data in sorted(div_data['jornadas'], key=lambda x: x['numero']):
            jornada_num = jornada_data['numero']
            id_jornada = f"{div_key}-jornada-{jornada_num}"
            contenido_html += f"""
            <div id="{id_jornada}" class="content-section" style="display:none;">
                <h2>Multas Jornada {jornada_num} - {div_titulo}</h2>
                {_generar_tabla_multas_jornada_html(jornada_data['multas'])}
            </div>"""
            nav_links_html += f'<a href="#" data-target="{id_jornada}">Jornada {jornada_num}</a>'
        nav_links_html += '</div></div>'


    html_completo = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Informe de Multas - SuperLiga Fuentmondo</title>
        <style>
            :root {{
                --primary-bg: #2c3e50; --secondary-bg: #34495e; --accent-color: #3498db;
                --text-light: #ecf0f1; --text-dark: #333; --border-color: #ddd;
                --body-bg: #f4f7f6; --white: #fff; --shadow: rgba(0,0,0,0.1);
            }}
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; margin: 0; background-color: var(--body-bg); color: var(--text-dark); }}
            .container {{ max-width: 1100px; margin: 20px auto; padding: 20px; background: var(--white); box-shadow: 0 0 15px var(--shadow); border-radius: 8px; }}
            .header {{ display: flex; align-items: center; justify-content: space-between; padding: 10px 20px; background-color: var(--primary-bg); color: var(--text-light); position: fixed; top: 0; width: 100%; box-sizing: border-box; z-index: 1001; }}
            .header h1 {{ margin: 0; font-size: 1.5em; }}
            .hamburger {{ display: none; font-size: 24px; background: none; border: none; color: var(--text-light); cursor: pointer; }}
            .navbar {{ display: flex; align-items: center; }}
            .navbar a, .dropbtn {{ display: inline-block; color: var(--text-light); text-align: center; padding: 14px 16px; text-decoration: none; font-size: 16px; border: none; background: none; cursor: pointer; outline: none; }}
            .navbar a:hover, .dropdown:hover .dropbtn {{ background-color: var(--secondary-bg); }}
            .navbar a.active {{ background-color: var(--accent-color); font-weight: bold; }}
            .dropdown {{ position: relative; display: inline-block; }}
            .dropdown-content {{ display: none; position: absolute; background-color: #f9f9f9; min-width: 160px; box-shadow: 0 8px 16px var(--shadow); z-index: 1; border-radius: 4px; overflow: hidden; }}
            .dropdown-content a {{ color: var(--text-dark); padding: 12px 16px; display: block; text-align: left; }}
            .dropdown-content a:hover {{ background-color: #f1f1f1; }}
            .dropdown-content.show {{ display: block; }}
            .overlay {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 999; }}
            main {{ padding-top: 80px; }}
            h2 {{ text-align: center; color: var(--primary-bg); border-bottom: 2px solid var(--accent-color); padding-bottom: 10px; margin-top: 0; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ padding: 12px 15px; border: 1px solid var(--border-color); text-align: left; }}
            th {{ background-color: var(--accent-color); color: var(--white); }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            .total-multa {{ font-weight: bold; text-align: center; color: #c0392b; }}
            .desglose ul {{ margin: 0; padding-left: 20px; }}
            .desglose li {{ margin-bottom: 5px; }}
            @media screen and (max-width: 850px) {{
                .header h1 {{ font-size: 1.2em; }}
                .hamburger {{ display: block; }}
                .navbar {{ position: fixed; top: 0; left: 0; height: 100%; width: 280px; background-color: var(--primary-bg); flex-direction: column; align-items: flex-start; padding-top: 60px; transform: translateX(-100%); transition: transform 0.3s ease-in-out; z-index: 1000; }}
                .navbar.open {{ transform: translateX(0); }}
                .navbar a, .dropbtn {{ width: 100%; text-align: left; padding: 15px 20px; box-sizing: border-box; }}
                .dropdown {{ width: 100%; }}
                .dropdown-content {{ position: static; box-shadow: none; background-color: var(--secondary-bg); border-radius: 0; }}
                .dropdown-content a {{ padding-left: 40px; color: var(--text-light); }}
                main {{ padding-top: 70px; }}
                .container {{ padding: 10px; margin: 10px; }}
                th, td {{ padding: 8px; font-size: 13px; }}
            }}
        </style>
    </head>
    <body>
        <div class="overlay"></div>
        <header class="header">
            <h1>Informe de Multas</h1>
            <button class="hamburger" aria-label="Abrir menú">☰</button>
            <nav class="navbar">{nav_links_html}</nav>
        </header>
        <main>
            <div class="container" id="main-container">{contenido_html}</div>
        </main>
        <script>
            document.addEventListener('DOMContentLoaded', function() {{
                const hamburger = document.querySelector('.hamburger');
                const navbar = document.querySelector('.navbar');
                const overlay = document.querySelector('.overlay');
                const navLinks = document.querySelectorAll('.navbar a, .navbar .dropbtn');

                function closeMenu() {{
                    navbar.classList.remove('open');
                    overlay.style.display = 'none';
                }}

                hamburger.addEventListener('click', function() {{
                    navbar.classList.toggle('open');
                    overlay.style.display = navbar.classList.contains('open') ? 'block' : 'none';
                }});

                overlay.addEventListener('click', closeMenu);

                function showContent(targetId) {{
                    document.querySelectorAll('.content-section').forEach(section => {{
                        section.style.display = 'none';
                    }});
                    const targetElement = document.getElementById(targetId);
                    if (targetElement) {{
                        targetElement.style.display = 'block';
                    }}

                    document.querySelectorAll('.nav-link').forEach(link => link.classList.remove('active'));
                    const activeLink = document.querySelector(`.nav-link[data-target='${{targetId}}']`);
                    if (activeLink) {{
                        activeLink.classList.add('active');
                    }}
                    if (window.innerWidth <= 850) {{
                        closeMenu();
                    }}
                }}

                navLinks.forEach(link => {{
                    link.addEventListener('click', function(e) {{
                        if (this.classList.contains('dropbtn')) {{
                            e.preventDefault();
                            this.nextElementSibling.classList.toggle('show');
                        }} else {{
                            e.preventDefault();
                            showContent(this.dataset.target);
                        }}
                    }});
                }});

                window.addEventListener('click', function(e) {{
                    if (!e.target.matches('.dropbtn')) {{
                        document.querySelectorAll('.dropdown-content.show').forEach(dd => {{
                            dd.classList.remove('show');
                        }});
                    }}
                }});

                showContent('{list(datos_informe.keys())[0]}-totales');
                const firstNavLink = document.querySelector('.nav-link');
                if(firstNavLink) firstNavLink.classList.add('active');
            }});
        </script>
    </body>
    </html>"""

    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_completo)
        print(f"Informe HTML completo guardado en '{output_path}'.")
    except Exception as e:
        print(f"Error al guardar el archivo HTML final: {e}")

# --- Funciones de Excel (sin cambios) ---
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

# --- Lógica de Procesamiento de Datos ---

# Procesa y genera un informe JSON completo para una ronda específica.
def procesar_ronda_completa(datos_ronda, output_file, payload_base, name_map={}):
    if not datos_ronda or 'answer' not in datos_ronda or 'matches' not in datos_ronda['answer']:
        print("Error: Respuesta de API de ronda inválida.")
        return None, None
    teams_in_round_list = datos_ronda['answer'].get('ranking', [])
    matches = datos_ronda['answer']['matches']
    round_id_actual = datos_ronda['query']['roundId']
    team_map_id = {i + 1: team['_id'] for i, team in enumerate(teams_in_round_list)}
    team_map_name = {i + 1: name_map.get(team['name'], team['name']) for i, team in enumerate(teams_in_round_list)}
    resultados_finales, puntos_equipos_por_ronda, jugadores_ronda = [], [], []
    dict_alineaciones, dict_capitanes = {}, {}
    for match in matches:
        ids = [team_map_id.get(p) for p in match['p']]
        nombres = [team_map_name.get(p) for p in match['p']]
        puntos = match.get('data', {}).get('partial', match.get('m', [0, 0]))
        for i in range(2):
            if nombres[i]:
                puntos_equipos_por_ronda.append({"equipo": nombres[i], "puntos": puntos[i]})
        lineups, capitanes = [], []
        for i in range(2):
            if not ids[i]: continue
            payload_lineup = {
                "header": copy.deepcopy(payload_base["header"]),
                "query": {"championshipId": payload_base["query"]["championshipId"], "round": round_id_actual, "userteamId": ids[i]}
            }
            datos_lineup = llamar_api("https://api.futmondo.com/1/userteam/roundlineup", payload_lineup)
            lineup_players = datos_lineup.get('answer', {}).get('players', [])
            lineups.append(lineup_players)
            capitan = next((p['name'] for p in lineup_players if p.get('cpt')), "N/A")
            capitanes.append(capitan)
            if nombres[i]:
                dict_alineaciones[nombres[i]] = lineup_players
                dict_capitanes[nombres[i]] = capitan
            for player in lineup_players:
                jugadores_ronda.append({
                    "nombre": player['name'], "puntos": player['points'],
                    "equipo": nombres[i], "es_capitan": player.get('cpt', False)
                })
        jugadores_repetidos = [p['name'] for p in lineups[0] if p['name'] in {p2['name'] for p2 in lineups[1]}]
        if nombres[0] and nombres[1]:
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
                if jugador['equipo'] and jugador['equipo'] not in peores_map[jugador['nombre']]:
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
    multas_jornada = calcular_multas_jornada(
        teams_in_round=list(team_map_name.values()), matches=matches, team_map_name=team_map_name,
        dict_alineaciones=dict_alineaciones, dict_capitanes=dict_capitanes,
        lista_peores_equipos=lista_peores_equipos, peores_jugadores_final=peores_jugadores_final,
        peores_capitanes_final=peores_capitanes_final
    )
    return multas_jornada

# Itera sobre todas las jornadas para procesar y DEVOLVER resultados y multas.
def procesar_historico_jornadas(rounds_map, payload_base, name_map, division_str):
    print(f"\n--- RECOPILANDO DATOS DE MULTAS PARA {division_str.upper()} ---")
    multas_acumuladas = defaultdict(float)
    datos_jornadas = []
    sorted_rounds = sorted(rounds_map.keys())
    for round_number in sorted_rounds:
        round_id = rounds_map[round_number]
        print(f"Procesando Jornada {round_number}...")
        payload_round = copy.deepcopy(payload_base)
        payload_round['query'].update({'roundNumber': round_id})
        datos_ronda = llamar_api("https://api.futmondo.com/1/ranking/round", payload_round)
        if datos_ronda and 'answer' in datos_ronda:
            datos_ronda['query']['roundId'] = round_id
            output_file = f"resultados/jornada_{round_number}_{division_str}.json"
            multas_de_la_jornada = procesar_ronda_completa(datos_ronda, output_file, payload_base, name_map)
            if multas_de_la_jornada:
                datos_jornadas.append({'numero': round_number, 'multas': multas_de_la_jornada})
                for team, data in multas_de_la_jornada.items():
                    multas_acumuladas[team] += data.get('multa_total', 0.0)
        else:
            print(f"No se pudieron obtener datos para la Jornada {round_number}. Saltando.")
    return datos_jornadas, dict(multas_acumuladas)

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

    rounds_data_1a = llamar_api("https://api.futmondo.com/1/userteam/rounds", copy.deepcopy(payload_1a))
    rounds_data_2a = llamar_api("https://api.futmondo.com/1/userteam/rounds", copy.deepcopy(payload_2a))
    rounds_map_1a = procesar_rondas_api(rounds_data_1a.get('answer', []))
    rounds_map_2a = procesar_rondas_api(rounds_data_2a.get('answer', []))
    if not rounds_map_1a or not rounds_map_2a:
        print("Error: No se pudo obtener y procesar la lista de rondas de la API. Finalizando.")
        return

    latest_round_id_1a = rounds_map_1a[max(rounds_map_1a.keys())]
    payload_round_1a = copy.deepcopy(payload_1a)
    payload_round_1a['query'].update({'roundNumber': latest_round_id_1a})
    datos_ronda_1a = llamar_api("https://api.futmondo.com/1/ranking/round", payload_round_1a)
    latest_round_id_2a = rounds_map_2a[max(rounds_map_2a.keys())]
    payload_round_2a = copy.deepcopy(payload_2a)
    payload_round_2a['query'].update({'roundNumber': latest_round_id_2a})
    datos_ronda_2a = llamar_api("https://api.futmondo.com/1/ranking/round", payload_round_2a)

    map_1a, map_2a = {}, {}
    if datos_ronda_1a and 'answer' in datos_ronda_1a and 'ranking' in datos_ronda_1a['answer']:
        round_ranking_1a = datos_ronda_1a['answer']['ranking']
        if len(round_ranking_1a) >= len(TEAMS_1A):
            map_1a = {round_ranking_1a[i]['name']: TEAMS_1A[str(i + 1)] for i in range(len(TEAMS_1A))}
    if datos_ronda_2a and 'answer' in datos_ronda_2a and 'ranking' in datos_ronda_2a['answer']:
        round_ranking_2a = datos_ronda_2a['answer']['ranking']
        if len(round_ranking_2a) >= len(TEAMS_2A):
            map_2a = {round_ranking_2a[i]['name']: TEAMS_2A[str(i + 1)] for i in range(len(TEAMS_2A))}

    # --- [LÓGICA CONDICIONAL] PROCESO DE EXCEL ---
    if modo in ['local', 'onedrive']:
        print("\n--- PROCESANDO ARCHIVO EXCEL ---")
        datos_general_1a = llamar_api("https://api.futmondo.com/1/ranking/general", copy.deepcopy(payload_1a))
        datos_general_2a = llamar_api("https://api.futmondo.com/1/ranking/general", copy.deepcopy(payload_2a))
        payload_teams_1a = copy.deepcopy(payload_1a); payload_teams_1a['query'] = {"championshipId": payload_1a["query"]["championshipId"]}
        datos_teams_1a = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams_1a)
        payload_teams_2a = copy.deepcopy(payload_2a); payload_teams_2a['query'] = {"championshipId": payload_2a["query"]["championshipId"]}
        datos_teams_2a = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams_2a)

        if all([datos_general_1a, datos_teams_1a, datos_general_2a, datos_teams_2a]):
            workbook = None
            try:
                if modo == 'local':
                    workbook = openpyxl.load_workbook(LOCAL_EXCEL_FILENAME)
                else:
                    access_token = get_access_token()
                    if not access_token: raise Exception("No se pudo obtener el token de acceso.")
                    drive_id, item_id = get_drive_item_from_share_link(access_token, ONEDRIVE_SHARE_LINK)
                    excel_content = download_excel_from_onedrive(access_token, drive_id, item_id)
                    workbook = openpyxl.load_workbook(io.BytesIO(excel_content))

                if workbook:
                    actualizar_cabeceras_capitanes(workbook, TEAMS_1A, TEAMS_2A)
                    actualizar_hoja_excel(workbook, datos_general_1a, datos_teams_1a, "Clasificación 1a DIV", 5, 2, map_1a)
                    actualizar_hoja_excel(workbook, datos_general_2a, datos_teams_2a, "Clasificación 2a DIV", 2, 3, map_2a)
                    actualizar_capitanes_historico(workbook, rounds_map_1a, payload_1a, "1a División", map_1a)
                    actualizar_capitanes_historico(workbook, rounds_map_2a, payload_2a, "2a División", map_2a)

                    if modo == 'local':
                        workbook.save(LOCAL_EXCEL_FILENAME)
                        print(f"\nArchivo '{LOCAL_EXCEL_FILENAME}' guardado localmente.")
                    else:
                        buffer = io.BytesIO()
                        workbook.save(buffer)
                        upload_excel_to_onedrive(access_token, drive_id, item_id, buffer.getvalue())
            except Exception as e:
                print(f"Error durante el procesamiento del Excel: {e}")
        else:
            print("Faltan datos clave de la API para el Excel. Saltando actualización del Excel.")
    else:
        print("\n--- MODO 'SOLO MULTAS' SELECCIONADO: SALTANDO PROCESO DE EXCEL ---")


    # --- GENERACIÓN DEL INFORME HTML (SE EJECUTA SIEMPRE) ---
    datos_jornadas_1a, totales_1a = procesar_historico_jornadas(rounds_map_1a, payload_1a, map_1a, "primera")
    datos_jornadas_2a, totales_2a = procesar_historico_jornadas(rounds_map_2a, payload_2a, map_2a, "segunda")

    datos_informe_completo = {
        "primera": {"jornadas": datos_jornadas_1a, "totales": totales_1a},
        "segunda": {"jornadas": datos_jornadas_2a, "totales": totales_2a}
    }

    generar_pagina_html_completa(datos_informe_completo, "informe_multas.html")

    print("\n--- Proceso completado. ---")

if __name__ == '__main__':
    main()
