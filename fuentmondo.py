import requests
import json
import copy
import openpyxl
from collections import defaultdict

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
# Mapeos de equipos para las divisiones
#
MAPA_EQUIPOS_1A_DIV = {
    1: "Galácticos de la noche FC", 2: "AL-CARRER F.C.", 3: "QUE BARBARIDAD FC",
    4: "Fuentino Pérez", 5: "CALAMARES CON TORRIJAS🦑🍞", 6: "CD Congelados",
    7: "THE LIONS", 8: "EL CHOLISMO FC", 9: "Real Fermín C.F.", 10: "Real 🥚🥚 Bailarines 🪩F.C",
    11: "MORRITOS F.C.", 12: "Poli Ejido CF", 13: "Juaki la bomba", 14: "LA MARRANERA",
    15: "Larios Limon FC", 16: "PANAKOTA F.F.", 17: "Real Pezqueñines FC",
    18: "LOS POKÉMON 🐭🟡🐭", 19: "El Huracán CF", 20: "Cangrena F.C."
}

MAPA_EQUIPOS_2A_DIV = {
    1: "SANTA LUCIA FC", 2: "Osasuna N.S.R", 3: "Tetitas Colesterol . F.C",
    4: "Fuentino Pérez", 5: "Charo la  Picanta FC", 6: "Kostas Mariotas",
    7: "Real Pescados el Puerto Fc", 8: "Team pepino", 9: "🇧🇷Samba Rovinha 🇧🇷",
    10: "Banano Vallekano 🍌⚡", 11: "SICARIOS CF", 12: "Minabo De Kiev",
    13: "Todo por la camiseta 🇪🇸", 14: "parker f.c.", 15: "Molinardo fc",
    16: "Lazaroneta", 17: "ElBarto F.C", 18: "BANANEROS FC", 19: "Morenetes de la Giralda �",
    20: "Jamon York F.C.", 21: "Elche pero Peor", 22: "Motobetis a primera!",
    23: "MTB Drink Team", 24: "Patejas"
}

#
# Función para actualizar la hoja de Excel con los datos de la clasificación.
#
def actualizar_excel(datos_general, datos_teams, nombre_archivo_excel, sheet_name, fila_inicio, columna_inicio):
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
        for row in sheet.iter_rows(min_row=fila_inicio, max_row=sheet.max_row, min_col=columna_inicio, max_col=columna_inicio+3):
            for cell in row:
                cell.value = None

        # Escribir los nuevos datos
        for i, equipo in enumerate(ranking_general):
            fila_actual = fila_inicio + i
            nombre_equipo = equipo['name']

            # Escribir la posición (solo si la columna de inicio es A)
            if columna_inicio == 1:
                sheet.cell(row=fila_actual, column=1).value = f"{i+1}º"

            # Escribir el nombre del equipo, puntos totales y puntos generales
            sheet.cell(row=fila_actual, column=columna_inicio).value = nombre_equipo
            sheet.cell(row=fila_actual, column=columna_inicio + 1).value = equipo['points']
            sheet.cell(row=fila_actual, column=columna_inicio + 2).value = puntos_generales.get(nombre_equipo, 0)

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
def procesar_y_actualizar_division(payload_file, excel_file, sheet_name, start_row, start_col):
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
        actualizar_excel(datos_general, datos_teams, excel_file, sheet_name, start_row, start_col)
    else:
        print(f"\nFallo al obtener los datos necesarios para '{sheet_name}'. No se pudo actualizar el Excel.")

#
# Función que procesa y actualiza los resultados de una ronda específica para una división.
#
def procesar_ronda_completa(payload_file, output_file, round_number, division):
    """
    Coordina la carga del payload, la llamada a la API de ronda y el procesamiento de resultados.
    """
    print(f"\n--- Procesando resultados de la ronda para la {division} división ---")
    API_URL_ROUND = "https://api.futmondo.com/1/ranking/round"
    API_URL_LINEUP = "https://api.futmondo.com/1/userteam/roundlineup"

    payload_base = cargar_payload(payload_file)
    if not payload_base:
        return

    # Obtener el mapa de equipos correcto para la división
    mapa_equipos = MAPA_EQUIPOS_1A_DIV if division == "primera" else MAPA_EQUIPOS_2A_DIV

    # 1. Obtener los emparejamientos de la ronda
    payload_round = copy.deepcopy(payload_base)
    payload_round['query']['roundNumber'] = round_number
    if 'roundId' in payload_round['query']:
        payload_round['query'].pop('roundId')

    datos_ronda = llamar_api(API_URL_ROUND, payload_round)

    # NUEVA VERIFICACIÓN DE CLAVES
    if not datos_ronda or 'answer' not in datos_ronda or 'ranking' not in datos_ronda['answer'] or 'matches' not in datos_ronda['answer']:
        print("Error: La respuesta de la API no contiene las claves 'answer', 'ranking' o 'matches'. Saltando el procesamiento de esta ronda.")
        return

    # Aquí está la corrección: usar 'ranking' en lugar de 'teams'
    teams_in_round = datos_ronda['answer']['ranking']
    matches = datos_ronda['answer']['matches']

    # Mapear el índice de equipo a su ID y nombre para un acceso más fácil
    team_map_id = {i+1: team['_id'] for i, team in enumerate(teams_in_round)}
    team_map_name = {i+1: team['name'] for i, team in enumerate(teams_in_round)}

    resultados_finales = []

    # Listas para almacenar los datos de los equipos para encontrar los 3 peores
    puntos_equipos_por_ronda = []
    jugadores_ronda = []

    # 2. Procesar cada combate de la ronda
    for match in matches:
        team_idx_1 = match['p'][0]
        team_idx_2 = match['p'][1]

        userteamId_1 = team_map_id.get(team_idx_1)
        userteamId_2 = team_map_id.get(team_idx_2)
        nombre_equipo_1 = team_map_name.get(team_idx_1)
        nombre_equipo_2 = team_map_name.get(team_idx_2)

        puntos_equipo_1 = match['data']['partial'][0]
        puntos_equipo_2 = match['data']['partial'][1]

        # Almacenar puntos totales de los equipos para la lista de 3 peores
        puntos_equipos_por_ronda.append({"equipo": nombre_equipo_1, "puntos": puntos_equipo_1})
        puntos_equipos_por_ronda.append({"equipo": nombre_equipo_2, "puntos": puntos_equipo_2})

        # 3. Obtener alineación del Equipo 1
        payload_lineup_1 = copy.deepcopy(payload_base)
        payload_lineup_1['query']['round'] = round_number
        payload_lineup_1['query']['userteamId'] = userteamId_1
        datos_lineup_1 = llamar_api(API_URL_LINEUP, payload_lineup_1)

        # 4. Obtener alineación del Equipo 2
        payload_lineup_2 = copy.deepcopy(payload_base)
        payload_lineup_2['query']['round'] = round_number
        payload_lineup_2['query']['userteamId'] = userteamId_2
        datos_lineup_2 = llamar_api(API_URL_LINEUP, payload_lineup_2)

        if not datos_lineup_1 or not datos_lineup_2 or 'players' not in datos_lineup_1['answer'] or 'players' not in datos_lineup_2['answer']:
            print(f"Error al obtener las alineaciones para el combate {nombre_equipo_1} vs {nombre_equipo_2}. Saltando...")
            continue

        lineup_1_players_list = datos_lineup_1['answer']['players']
        lineup_2_players_list = datos_lineup_2['answer']['players']

        # Encontrar capitán del Equipo 1
        capitan_1 = ""
        for player in lineup_1_players_list:
            if player.get('cpt'):
                capitan_1 = player['name']
                break

        # Encontrar capitán del Equipo 2
        capitan_2 = ""
        for player in lineup_2_players_list:
            if player.get('cpt'):
                capitan_2 = player['name']
                break

        # Almacenar jugadores y capitanes para encontrar los peores
        for player in lineup_1_players_list:
            jugadores_ronda.append({"nombre": player['name'], "puntos": player['points'], "equipo": nombre_equipo_1, "es_capitan": player.get('cpt')})
        for player in lineup_2_players_list:
            jugadores_ronda.append({"nombre": player['name'], "puntos": player['points'], "equipo": nombre_equipo_2, "es_capitan": player.get('cpt')})

        # Encontrar jugadores repetidos
        jugadores_repetidos = [
            player['name'] for player in lineup_1_players_list if player['name'] in [p['name'] for p in lineup_2_players_list]
        ]

        # Construir el objeto de resultados
        resultado_combate = {
            "Combate": f"{nombre_equipo_1} vs {nombre_equipo_2}",
            f"{nombre_equipo_1}": {"Puntuacion": puntos_equipo_1, "Capitan": capitan_1},
            f"{nombre_equipo_2}": {"Puntuacion": puntos_equipo_2, "Capitan": capitan_2},
            "Jugadores repetidos": jugadores_repetidos
        }
        resultados_finales.append(resultado_combate)

    # 5. Encontrar los 3 peores equipos por puntuación total en la ronda
    peores_equipos = sorted(puntos_equipos_por_ronda, key=lambda x: x['puntos'])[:3]

    # Formatear la lista de peores equipos
    lista_peores_equipos = []
    for i, equipo in enumerate(peores_equipos):
        lista_peores_equipos.append({
            "posicion": i + 1,
            "equipo": equipo['equipo'],
            "puntos": equipo['puntos']
        })

    # Encontrar el/los peor(es) capitán(es) y unificarlos
    peores_capitanes_map = defaultdict(list)
    min_puntos_capitan = float('inf')

    for jugador in jugadores_ronda:
        if jugador['es_capitan']:
            if jugador['puntos'] < min_puntos_capitan:
                min_puntos_capitan = jugador['puntos']
                peores_capitanes_map.clear()
                peores_capitanes_map[jugador['nombre']].append(jugador['equipo'])
            elif jugador['puntos'] == min_puntos_capitan:
                if jugador['equipo'] not in peores_capitanes_map[jugador['nombre']]:
                    peores_capitanes_map[jugador['nombre']].append(jugador['equipo'])

    peores_capitanes_final = [
        {"nombre": nombre, "puntos": min_puntos_capitan, "equipos": equipos}
        for nombre, equipos in peores_capitanes_map.items()
    ]

    # Encontrar el/los peor(es) jugador(es) y unificarlos
    peores_jugadores_map = defaultdict(list)
    min_puntos_jugador = float('inf')

    for jugador in jugadores_ronda:
        if jugador['puntos'] < min_puntos_jugador:
            min_puntos_jugador = jugador['puntos']
            peores_jugadores_map.clear()
            peores_jugadores_map[jugador['nombre']].append(jugador['equipo'])
        elif jugador['puntos'] == min_puntos_jugador:
            if jugador['equipo'] not in peores_jugadores_map[jugador['nombre']]:
                peores_jugadores_map[jugador['nombre']].append(jugador['equipo'])

    peores_jugadores_final = [
        {"nombre": nombre, "puntos": min_puntos_jugador, "equipos": equipos}
        for nombre, equipos in peores_jugadores_map.items()
    ]

    # 6. Guardar la lista de resultados y el resumen
    resumen_final = {
        "Resultados por combate": resultados_finales,
        "Peor Capitan": peores_capitanes_final,
        "Peor Jugador": peores_jugadores_final,
        "Los 3 peores equipos de la ronda": lista_peores_equipos
    }

    guardar_respuesta(resumen_final, output_file)
    print(f"\nResumen de resultados de la ronda guardado en '{output_file}'.")


#
# Función principal del script.
#
def main():
    """
    Función principal que coordina las llamadas a la API y la gestión de archivos para ambas divisiones.
    """
    # Configuración de la 1a División
    PAYLOAD_PRIMERA = "payload_primera.json"
    EXCEL_FILE = "SuperLiga Fuentmondo 25-26.xlsx"
    SHEET_PRIMERA = "Clasificación 1a DIV"
    ROW_START_PRIMERA = 5 # Fila de inicio: 5. Columna de inicio: A
    COL_START_PRIMERA = 2 # A es la columna 1
    ROUND_NUMBER = "6868e3437775a433449ea12b"

    # Configuración de la 2a División
    PAYLOAD_SEGUNDA = "payload.json"
    SHEET_SEGUNDA = "Clasificación 2a DIV"
    ROW_START_SEGUNDA = 2 # Fila de inicio: 2. Columna de inicio: B
    COL_START_SEGUNDA = 3 # B es la columna 2

    # Procesar la 1a División
    procesar_y_actualizar_division(PAYLOAD_PRIMERA, EXCEL_FILE, SHEET_PRIMERA, ROW_START_PRIMERA, COL_START_PRIMERA)

    # Procesar la 2a División
    procesar_y_actualizar_division(PAYLOAD_SEGUNDA, EXCEL_FILE, SHEET_SEGUNDA, ROW_START_SEGUNDA, COL_START_SEGUNDA)

    # --- Proceso de ronda para la 1a División ---
    procesar_ronda_completa(
        payload_file=PAYLOAD_PRIMERA,
        output_file="resultados_ronda_1a_div.json",
        round_number=ROUND_NUMBER,
        division="primera"
    )

    # --- Proceso de ronda para la 2a División ---
    procesar_ronda_completa(
        payload_file=PAYLOAD_SEGUNDA,
        output_file="resultados_ronda_2a_div.json",
        round_number=ROUND_NUMBER, # NOTA: Usa el roundNumber correcto aquí para la 2a división
        division="segunda"
    )


if __name__ == '__main__':
    main()

