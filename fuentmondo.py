from datetime import datetime
import subprocess
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

load_dotenv()

# --- CONFIGURACIÓN GLOBAL ---
CLIENT_ID = os.getenv("CLIENT_ID")
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
AUTHORITY = 'https://login.microsoftonline.com/common/'
SCOPES = ['Files.ReadWrite.All']
ONEDRIVE_SHARE_LINK = "https://1drv.ms/x/s!AidvQapyuNp6jBKR5uMUCaBYdLl0?e=3kXyKW"

# --- FUNCIONES DE INTERFAZ GRÁFICA (TKINTER) ---

def choose_save_option():
    """Crea y muestra una ventana para que el usuario elija el modo de ejecución."""
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

    btn_multas = tk.Button(button_frame, text="Generar Solo Informe (HTML)", command=lambda: select_option('multas_only'), height=2, width=25, bg="#28a745", fg="white")
    btn_multas.pack(side=tk.LEFT, padx=5)

    root.mainloop()
    return choice[0]

def show_auth_code_window(message, verification_uri):
    """Crea una ventana con el código de autenticación para que el usuario lo copie."""
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

# --- FUNCIONES DE AUTENTICACIÓN Y ONEDRIVE ---

def get_access_token():
    """Se autentica de forma interactiva y obtiene un token de acceso para Microsoft Graph."""
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

def encode_sharing_link(sharing_link):
    """Codifica un enlace de compartición de OneDrive a un formato compatible con la API de Graph."""
    base64_value = base64.b64encode(sharing_link.encode('utf-8')).decode('utf-8')
    return 'u!' + base64_value.rstrip('=').replace('/', '_').replace('+', '-')

def get_drive_item_from_share_link(access_token, share_url):
    """Obtiene el ID del Drive y el ID del archivo a partir de un enlace de compartición."""
    encoded_url = encode_sharing_link(share_url)
    api_url = f"{GRAPH_API_ENDPOINT}/shares/{encoded_url}/driveItem"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(api_url, headers=headers)
    response.raise_for_status()
    data = response.json()
    return data['parentReference']['driveId'], data['id']

def download_excel_from_onedrive(access_token, drive_id, item_id):
    """Descarga el contenido de un archivo Excel desde OneDrive."""
    api_url = f"{GRAPH_API_ENDPOINT}/drives/{drive_id}/items/{item_id}/content"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(api_url, headers=headers)
    response.raise_for_status()
    print("Excel descargado de OneDrive con éxito.")
    return response.content

def upload_excel_to_onedrive(access_token, drive_id, item_id, file_content):
    """Sube (o sobrescribe) el contenido de un archivo Excel a OneDrive, con reintentos si está bloqueado."""
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

# --- FUNCIONES DE API Y LÓGICA DE DATOS ---

def cargar_payload(ruta_archivo):
    """Carga un archivo JSON (payload) desde una ruta específica."""
    try:
        with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
            return json.load(archivo)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error al cargar '{ruta_archivo}': {e}")
        return None

def llamar_api(url, payload):
    """Realiza una llamada POST a una API con un payload JSON."""
    if not payload: return None
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error en la llamada a la API '{url}': {e}")
        return None

def guardar_respuesta(datos, nombre_archivo):
    """Guarda los datos de respuesta de la API en un archivo JSON."""
    try:
        os.makedirs(os.path.dirname(nombre_archivo), exist_ok=True)
        with open(nombre_archivo, 'w', encoding='utf-8') as f:
            json.dump(datos, f, indent=4, ensure_ascii=False)
        print(f"Respuesta de la API guardada en '{nombre_archivo}'.")
    except Exception as e:
        print(f"Error al guardar el archivo '{nombre_archivo}': {e}")

def get_captains_for_round(payload_base, datos_ronda, name_map={}):
    """Obtiene una lista de los capitanes de todos los equipos para una ronda específica."""
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

def procesar_rondas_api(rounds_list):
    """Procesa la respuesta de la API de rondas, convirtiendo los números de ronda a enteros."""
    if not rounds_list:
        return {}
    rounds_map = {}
    for r in rounds_list:
        round_num = r.get('number')
        round_id = r.get('id')
        if not round_num or not round_id:
            continue
        if round_num % 1 == 0:
            rounds_map[int(round_num)] = round_id
    return rounds_map

def calcular_multas_jornada(teams_in_round, matches, team_map_name, dict_alineaciones, dict_capitanes, lista_peores_equipos, peores_jugadores_final, peores_capitanes_final):
    """Calcula las multas de una jornada con un desglose detallado."""
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

        repetidos_iniciales = nombres_a.intersection(nombres_b)
        repetidos_para_multa = repetidos_iniciales.copy()

        if capitan_a != "N/A":
            repetidos_para_multa.discard(capitan_a)
        if capitan_b != "N/A":
            repetidos_para_multa.discard(capitan_b)

        if repetidos_para_multa:
            cantidad_repetidos_multables = len(repetidos_para_multa)
            multa_repetidos = cantidad_repetidos_multables * 0.5
            multas_finales[team_a_name]["desglose"]["jugadores_repetidos"] = {"cantidad": cantidad_repetidos_multables, "multa": multa_repetidos}
            multas_finales[team_b_name]["desglose"]["jugadores_repetidos"] = {"cantidad": cantidad_repetidos_multables, "multa": multa_repetidos}

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

# --- FUNCIONES DE GENERACIÓN DE HTML (CON TAILWIND CSS) ---

def _generar_tabla_multas_jornada_html(multas_data):
    """Genera el HTML para la tabla de multas de una jornada."""
    sorted_teams = sorted(multas_data.items(), key=lambda item: item[1]['multa_total'], reverse=True)
    table_rows = ""
    for i, (team_name, data) in enumerate(sorted_teams):
        multa_total = data.get('multa_total', 0.0)
        if multa_total == 0: continue
        desglose = data.get('desglose', {})
        desglose_html = "<ul class='list-disc list-inside space-y-1'>"

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

        row_bg = 'bg-slate-50' if i % 2 != 0 else 'bg-white'
        table_rows += f"""
        <tr class="{row_bg}">
            <td class="p-3 border border-slate-300">{team_name}</td>
            <td class="p-3 border border-slate-300 text-center font-bold text-red-600">{multa_total:.2f}€</td>
            <td class="p-3 border border-slate-300">{desglose_html}</td>
        </tr>"""
    if not table_rows:
        table_rows = '<tr><td colspan="3" class="text-center p-4 border border-slate-300">No se registraron multas en esta jornada.</td></tr>'

    return f"""
    <div class="overflow-x-auto">
        <table class="w-full text-left border-collapse">
            <thead class="bg-slate-200">
                <tr>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300">Equipo</th>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300 text-center">Multa Total</th>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300">Desglose</th>
                </tr>
            </thead>
            <tbody>{table_rows}</tbody>
        </table>
    </div>"""

def _generar_tabla_multas_totales_html(multas_acumuladas):
    """Genera el HTML para la tabla de multas totales acumuladas."""
    sorted_teams = sorted(multas_acumuladas.items(), key=lambda item: item[1], reverse=True)
    table_rows = ""
    for i, (team_name, total_multa) in enumerate(sorted_teams):
        row_bg = 'bg-slate-50' if i % 2 != 0 else 'bg-white'
        table_rows += f"""
        <tr class="{row_bg}">
            <td class="p-3 border border-slate-300">{team_name}</td>
            <td class="p-3 border border-slate-300 text-center font-bold text-red-600">{total_multa:.2f}€</td>
        </tr>"""
    return f"""
    <div class="overflow-x-auto">
        <table class="w-full text-left border-collapse">
            <thead class="bg-slate-200">
                <tr>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300">Equipo</th>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300 text-center">Total Acumulado</th>
                </tr>
            </thead>
            <tbody>{table_rows}</tbody>
        </table>
    </div>"""

def _generar_tabla_clasificacion_html(ranking_ordenado):
    """Genera el HTML para la tabla de clasificación."""
    table_rows = ""
    for i, equipo in enumerate(ranking_ordenado):
        row_bg = 'bg-slate-50' if i % 2 != 0 else 'bg-white'
        table_rows += f"""
        <tr class="{row_bg}">
            <td class="p-3 border border-slate-300 text-center">{i + 1}</td>
            <td class="p-3 border border-slate-300">{equipo['name']}</td>
            <td class="p-3 border border-slate-300 text-center">{equipo['points']}</td>
            <td class="p-3 border border-slate-300 text-center">{equipo['general_points']}</td>
        </tr>"""
    return f"""
    <div class="overflow-x-auto">
        <table class="w-full text-left border-collapse">
            <thead class="bg-slate-200">
                <tr>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300 text-center">Pos.</th>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300">Equipo</th>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300 text-center">Puntos (J)</th>
                    <th class="p-3 font-bold uppercase text-slate-600 border border-slate-300 text-center">Puntos (G)</th>
                </tr>
            </thead>
            <tbody>{table_rows}</tbody>
        </table>
    </div>"""

def _generar_tabla_capitanes_html(datos_capitanes, team_names):
    """Genera el HTML para la tabla del historial de capitanes."""
    if not datos_capitanes: return "<p>No hay datos de capitanes disponibles.</p>"
    sorted_jornadas = sorted(datos_capitanes.keys())
    sorted_teams = sorted(list(team_names))
    header_cols = "".join(f"<th class='p-3 font-bold uppercase text-slate-600 border border-slate-300 sticky top-0 bg-slate-200'>{team}</th>" for team in sorted_teams)
    header = f"<tr><th class='p-3 font-bold uppercase text-slate-600 border border-slate-300 sticky top-0 bg-slate-200'>Jornada</th>{header_cols}</tr>"

    body_rows = ""
    for i, jornada_num in enumerate(sorted_jornadas):
        capitanes_jornada = {item['team_name']: item for item in datos_capitanes[jornada_num]}
        row_bg = 'bg-slate-50' if i % 2 != 0 else 'bg-white'
        row_cols = f"<td class='p-3 border border-slate-300 font-semibold'>Jornada {jornada_num}</td>"
        for team_name in sorted_teams:
            cap_info = capitanes_jornada.get(team_name)
            if cap_info:
                capitan_name = cap_info.get('capitan', 'N/A')
                is_repeated = cap_info.get('is_repeated_3_times', False)
                class_attr = ' class="bg-yellow-300 font-semibold"' if is_repeated else ''
                row_cols += f'<td{class_attr} class="p-3 border border-slate-300">{capitan_name}</td>'
            else:
                row_cols += '<td class="p-3 border border-slate-300">-</td>'
        body_rows += f"<tr class='{row_bg}'>{row_cols}</tr>"

    return f"""
    <div class="overflow-x-auto">
        <table class="w-full text-left border-collapse">
            <thead>{header}</thead>
            <tbody>{body_rows}</tbody>
        </table>
    </div>"""

def generar_pagina_html_completa(datos_informe, output_path):
    """Genera la página HTML completa con todos los datos y la navegación usando Tailwind CSS."""
    contenido_html = ""
    nav_links_html = ""

    for div_key, div_data in datos_informe.items():
        div_titulo = "1ª División" if div_key == "primera" else "2ª División"

        id_clasificacion = f"{div_key}-clasificacion"
        id_capitanes = f"{div_key}-capitanes"
        id_totales = f"{div_key}-totales"

        nav_links_html += f'<a href="#" class="nav-link block px-4 py-2 text-white hover:bg-slate-700 md:inline-block" data-target="{id_clasificacion}">Clasificación {div_titulo}</a>'
        nav_links_html += f'<a href="#" class="nav-link block px-4 py-2 text-white hover:bg-slate-700 md:inline-block" data-target="{id_capitanes}">Capitanes {div_titulo}</a>'
        nav_links_html += f'<a href="#" class="nav-link block px-4 py-2 text-white hover:bg-slate-700 md:inline-block" data-target="{id_totales}">Multas Totales {div_titulo}</a>'

        contenido_html += f'<div id="{id_clasificacion}" class="content-section p-6 bg-white rounded-lg shadow-md"> <h2 class="text-2xl font-bold text-center text-slate-700 border-b-2 border-sky-500 pb-3 mb-6">Clasificación - {div_titulo}</h2> {_generar_tabla_clasificacion_html(div_data["clasificacion"])} </div>'
        contenido_html += f'<div id="{id_capitanes}" class="content-section p-6 bg-white rounded-lg shadow-md"> <h2 class="text-2xl font-bold text-center text-slate-700 border-b-2 border-sky-500 pb-3 mb-6">Historial de Capitanes - {div_titulo}</h2> {_generar_tabla_capitanes_html(div_data["capitanes"], div_data["totales"].keys())} </div>'
        contenido_html += f'<div id="{id_totales}" class="content-section p-6 bg-white rounded-lg shadow-md"> <h2 class="text-2xl font-bold text-center text-slate-700 border-b-2 border-sky-500 pb-3 mb-6">Multas Totales - {div_titulo}</h2> {_generar_tabla_multas_totales_html(div_data["totales"])} </div>'

        nav_links_html += '<div class="relative dropdown-container">'
        nav_links_html += f'<button class="dropdown-btn block w-full text-left px-4 py-2 text-white hover:bg-slate-700 md:inline-block md:w-auto">Multas Jornada ({div_titulo}) &#9662;</button>'
        nav_links_html += '<div class="dropdown-content hidden md:absolute bg-white text-black rounded-md shadow-lg mt-2 py-1 z-20">'

        sorted_jornadas = sorted(div_data['jornadas'], key=lambda x: x['numero'])
        for jornada_data in sorted_jornadas:
            jornada_num = jornada_data['numero']
            id_jornada = f"{div_key}-jornada-{jornada_num}"
            contenido_html += f'<div id="{id_jornada}" class="content-section hidden p-6 bg-white rounded-lg shadow-md"> <h2 class="text-2xl font-bold text-center text-slate-700 border-b-2 border-sky-500 pb-3 mb-6">Multas Jornada {jornada_num} - {div_titulo}</h2> {_generar_tabla_multas_jornada_html(jornada_data["multas"])} </div>'
            nav_links_html += f'<a href="#" class="block px-4 py-2 hover:bg-slate-100" data-target="{id_jornada}">Jornada {jornada_num}</a>'

        nav_links_html += '</div></div>'

    html_completo = f"""
    <!DOCTYPE html>
    <html lang="es" class="scroll-smooth">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Informe - SuperLiga Fuentmondo</title>
        <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    </head>
    <body class="bg-slate-100 text-slate-800 font-sans">

        <header class="bg-slate-800 text-white flex justify-between items-center p-4 shadow-lg fixed top-0 left-0 right-0 z-50">
            <h1 class="text-xl font-bold">Informe SuperLiga</h1>
            <button id="hamburger-btn" class="md:hidden text-2xl">☰</button>
            <nav id="navbar" class="fixed top-0 left-0 h-full w-64 bg-slate-800 transform -translate-x-full transition-transform duration-300 ease-in-out md:relative md:translate-x-0 md:flex md:w-auto md:h-auto md:bg-transparent">
                <div class="p-4 md:flex md:items-center md:gap-2">
                    {nav_links_html}
                </div>
            </nav>
        </header>

        <div id="overlay" class="fixed inset-0 bg-black bg-opacity-50 z-30 hidden md:hidden"></div>

        <main class="pt-20">
            <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">
                {contenido_html}
            </div>
        </main>

        <script>
            document.addEventListener('DOMContentLoaded', () => {{
                const hamburgerBtn = document.getElementById('hamburger-btn');
                const navbar = document.getElementById('navbar');
                const overlay = document.getElementById('overlay');
                const navLinks = document.querySelectorAll('.nav-link, .dropdown-content a');
                const dropdownBtns = document.querySelectorAll('.dropdown-btn');

                function toggleMenu() {{
                    const isOffScreen = navbar.classList.contains('-translate-x-full');
                    navbar.classList.toggle('-translate-x-full', !isOffScreen);
                    navbar.classList.toggle('translate-x-0', isOffScreen);
                    overlay.classList.toggle('hidden');
                }}

                hamburgerBtn.addEventListener('click', toggleMenu);
                overlay.addEventListener('click', toggleMenu);

                function showContent(targetId) {{
                    document.querySelectorAll('.content-section').forEach(section => section.classList.add('hidden'));
                    const targetElement = document.getElementById(targetId);
                    if (targetElement) {{
                        targetElement.classList.remove('hidden');
                    }}

                    document.querySelectorAll('.nav-link').forEach(link => link.classList.remove('bg-sky-600'));
                    const activeLink = document.querySelector(`.nav-link[data-target='${{targetId}}']`);
                    if (activeLink) {{
                        activeLink.classList.add('bg-sky-600');
                    }}
                }}

                navLinks.forEach(link => {{
                    link.addEventListener('click', e => {{
                        e.preventDefault();
                        const targetId = e.currentTarget.dataset.target;
                        showContent(targetId);
                        if (window.innerWidth < 768) {{
                            toggleMenu();
                        }}
                        document.querySelectorAll('.dropdown-content').forEach(d => d.classList.add('hidden'));
                    }});
                }});

                dropdownBtns.forEach(btn => {{
                    btn.addEventListener('click', e => {{
                        e.stopPropagation();
                        const dropdownContent = e.currentTarget.nextElementSibling;
                        document.querySelectorAll('.dropdown-content').forEach(d => {{
                            if (d !== dropdownContent) d.classList.add('hidden');
                        }});
                        dropdownContent.classList.toggle('hidden');
                    }});
                }});

                window.addEventListener('click', () => {{
                    document.querySelectorAll('.dropdown-content').forEach(d => d.classList.add('hidden'));
                }});

                const firstSectionId = document.querySelector('.content-section')?.id;
                if (firstSectionId) {{
                    showContent(firstSectionId);
                }}
            }});
        </script>
    </body>
    </html>"""

    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_completo)
        print(f"Informe HTML (Tailwind CSS) completo guardado en '{output_path}'.")
    except Exception as e:
        print(f"Error al guardar el archivo HTML final: {e}")

# --- FUNCIONES DE EXCEL ---

def _procesar_y_ordenar_clasificacion(datos_general, datos_teams, name_map={}):
    """Procesa y ordena los datos de clasificación de la API."""
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
    return sorted(equipos_para_ordenar, key=lambda x: (x['points'], x['general_points']), reverse=True)

def actualizar_hoja_excel(workbook, ranking_ordenado, sheet_name, fila_inicio, columna_inicio):
    """Actualiza una hoja de Excel con los datos de clasificación."""
    try:
        sheet = workbook[sheet_name]
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
    """Actualiza las cabeceras con los nombres de los equipos en la hoja 'Capitanes'."""
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
    """Actualiza la fila de una jornada con los capitanes de cada equipo en el Excel."""
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
                        print(f"   -> Aviso: El equipo '{team_name}' no se encontró en la cabecera de la hoja 'Capitanes'.")
                print(f"Capitanes de la '{target_row_label}' actualizados en memoria.")
                row_found = True
                break
        if not row_found:
            print(f"Advertencia: No se encontró la fila para '{target_row_label}' en la hoja 'Capitanes'.")
    except Exception as e:
        print(f"Error actualizando la hoja 'Capitanes' para la jornada {round_number}: {e}")

def actualizar_capitanes_historico(workbook, rounds_map, payload_base, division_name, name_map={}):
    """Itera sobre todas las jornadas para actualizar el histórico de capitanes en el Excel."""
    print(f"\n--- INICIANDO ACTUALIZACIÓN HISTÓRICA DE CAPITANES PARA {division_name.upper()} (EXCEL) ---")
    sorted_round_numbers = sorted(rounds_map.keys())
    for round_number in sorted_round_numbers:
        round_id = rounds_map[round_number]
        print(f"Procesando capitanes de la Jornada {round_number} para Excel...")
        payload_round = copy.deepcopy(payload_base)
        payload_round['query'].update({'roundNumber': round_id, 'championshipId': payload_base['query']['championshipId']})
        datos_ronda = llamar_api("https://api.futmondo.com/1/ranking/round", payload_round)
        if not datos_ronda or 'answer' not in datos_ronda or datos_ronda['answer'] == 'api.error.general':
            print(f"   -> Error: No se pudieron obtener datos para la Jornada {round_number}. Saltando.")
            continue
        datos_ronda['query']['roundNumber'] = round_id
        team_captains = get_captains_for_round(payload_base, datos_ronda, name_map)
        if not team_captains:
            print(f"   -> Advertencia: No se encontraron capitanes para la Jornada {round_number}.")
            continue
        actualizar_hoja_capitanes(workbook, round_number, team_captains)

# --- LÓGICA DE PROCESAMIENTO DE DATOS ---

def procesar_ronda_completa(datos_ronda, output_file, payload_base, name_map={}):
    """Procesa todos los datos de una ronda y calcula las multas correspondientes."""
    if not datos_ronda or 'answer' not in datos_ronda or 'matches' not in datos_ronda['answer']:
        print("Error: Respuesta de API de ronda inválida.")
        return None
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

def procesar_historico_jornadas(rounds_map, payload_base, name_map, division_str):
    """Itera sobre todas las jornadas para procesar y devolver los resultados y multas."""
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

def recopilar_historico_capitanes(rounds_map, payload_base, name_map, division_str):
    """Recopila el historial de capitanes y detecta la tercera repetición."""
    print(f"\n--- RECOPILANDO HISTORIAL DE CAPITANES PARA {division_str.upper()} ---")
    capitanes_por_jornada = {}
    contador_capitanes = defaultdict(lambda: defaultdict(int))
    sorted_rounds = sorted(rounds_map.keys())

    for round_number in sorted_rounds:
        round_id = rounds_map[round_number]
        print(f"Obteniendo capitanes de la Jornada {round_number}...")
        payload_round = copy.deepcopy(payload_base)
        payload_round['query'].update({'roundNumber': round_id})
        datos_ronda = llamar_api("https://api.futmondo.com/1/ranking/round", payload_round)

        if not datos_ronda or 'answer' not in datos_ronda:
            print(f"No se pudieron obtener datos para la Jornada {round_number}. Saltando.")
            continue

        datos_ronda['query']['roundNumber'] = round_id
        team_captains = get_captains_for_round(payload_base, datos_ronda, name_map)

        jornada_data = []
        for cap_info in team_captains:
            team_name = cap_info['team_name']
            capitan = cap_info['capitan']
            if capitan != "N/A":
                contador_capitanes[team_name][capitan] += 1

            cap_info['is_repeated_3_times'] = (contador_capitanes[team_name][capitan] == 3)
            jornada_data.append(cap_info)

        capitanes_por_jornada[round_number] = jornada_data

    return capitanes_por_jornada

# --- FUNCIONES DE GIT ---

def subir_informe_a_github(ruta_archivo_html, GITHUB_TOKEN, GITHUB_USERNAME, GITHUB_REPO):
    """Sube el informe HTML a un repositorio de GitHub automáticamente."""
    print("\n--- INTENTANDO SUBIR INFORME A GITHUB ---")
    try:
        remote_url = f"https://{GITHUB_TOKEN}@github.com/{GITHUB_USERNAME}/{GITHUB_REPO}.git"

        subprocess.run(["git", "add", ruta_archivo_html], check=True, capture_output=True, text=True)

        # Comprobar si hay cambios para hacer commit
        status_result = subprocess.run(["git", "status", "--porcelain"], check=True, capture_output=True, text=True)
        if ruta_archivo_html not in status_result.stdout:
            print("✅ No hay cambios detectados en el informe. No se necesita subir nada.")
            return

        mensaje_commit = f"Informe actualizado automáticamente - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        subprocess.run(["git", "commit", "-m", mensaje_commit], check=True, capture_output=True, text=True)

        subprocess.run(["git", "push", remote_url], check=True, capture_output=True, text=True)

        print(f"✅ Informe '{ruta_archivo_html}' subido a GitHub con éxito.")

    except FileNotFoundError:
        print("❌ Error: Git no está instalado o no se encuentra en el PATH del sistema.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error durante la ejecución de un comando de Git: {e.stderr}")
    except Exception as e:
        print(f"❌ Ocurrió un error inesperado al intentar subir a GitHub: {e}")

# --- FUNCIÓN PRINCIPAL ---

def main():
    """Función principal que orquesta la ejecución del script."""
    GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
    GITHUB_USERNAME = os.getenv("GITHUB_USERNAME")
    GITHUB_REPO = os.getenv("GITHUB_REPO")

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
      "4":"Pollos sin cabeza 🐥🧄", "5":"Charo la   Picanta FC", "6":"Kostas Mariotas",
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

    datos_general_1a = llamar_api("https://api.futmondo.com/1/ranking/general", copy.deepcopy(payload_1a))
    datos_general_2a = llamar_api("https://api.futmondo.com/1/ranking/general", copy.deepcopy(payload_2a))
    payload_teams_1a = copy.deepcopy(payload_1a); payload_teams_1a['query'] = {"championshipId": payload_1a["query"]["championshipId"]}
    datos_teams_1a = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams_1a)
    payload_teams_2a = copy.deepcopy(payload_2a); payload_teams_2a['query'] = {"championshipId": payload_2a["query"]["championshipId"]}
    datos_teams_2a = llamar_api("https://api.futmondo.com/2/championship/teams", payload_teams_2a)

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

    clasificacion_1a = _procesar_y_ordenar_clasificacion(datos_general_1a, datos_teams_1a, map_1a)
    clasificacion_2a = _procesar_y_ordenar_clasificacion(datos_general_2a, datos_teams_2a, map_2a)

    if modo in ['local', 'onedrive']:
        print("\n--- PROCESANDO ARCHIVO EXCEL ---")
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
                    actualizar_hoja_excel(workbook, clasificacion_1a, "Clasificación 1a DIV", 5, 2)
                    actualizar_hoja_excel(workbook, clasificacion_2a, "Clasificación 2a DIV", 2, 3)
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
        print("\n--- MODO 'SOLO INFORME' SELECCIONADO: SALTANDO PROCESO DE EXCEL ---")

    datos_jornadas_1a, totales_1a = procesar_historico_jornadas(rounds_map_1a, payload_1a, map_1a, "primera")
    datos_jornadas_2a, totales_2a = procesar_historico_jornadas(rounds_map_2a, payload_2a, map_2a, "segunda")

    capitanes_1a = recopilar_historico_capitanes(rounds_map_1a, payload_1a, map_1a, "primera")
    capitanes_2a = recopilar_historico_capitanes(rounds_map_2a, payload_2a, map_2a, "segunda")

    datos_informe_completo = {
        "primera": {
            "jornadas": datos_jornadas_1a,
            "totales": totales_1a,
            "clasificacion": clasificacion_1a,
            "capitanes": capitanes_1a
        },
        "segunda": {
            "jornadas": datos_jornadas_2a,
            "totales": totales_2a,
            "clasificacion": clasificacion_2a,
            "capitanes": capitanes_2a
        }
    }

    generar_pagina_html_completa(datos_informe_completo, "index.html")

    if all([GITHUB_TOKEN, GITHUB_USERNAME, GITHUB_REPO]):
        subir_informe_a_github("index.html", GITHUB_TOKEN, GITHUB_USERNAME, GITHUB_REPO)
    else:
        print("\n--- AVISO: Faltan variables de entorno de GitHub (.env) para la subida automática. ---")

    print("\n--- Proceso completado. ---")

if __name__ == '__main__':
    main()
