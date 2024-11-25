import aiohttp
import asyncio
import pandas as pd
from datetime import datetime, timedelta
import os
from openpyxl import load_workbook

# Función para obtener los datos de la API
async def fetch_twitch_data(date):
    url = "https://twitchtracker.com/api/channels/summary/rubius"  # URL de la API para el canal Rubius
    async with aiohttp.ClientSession() as session:
        async with session.get(url) as response:
            if response.status == 200:
                data = await response.json()  # Obtener datos en formato JSON
                return data
            else:
                print(f"Error al acceder a la API en {date}: {response.status}")
                return None

# Función para generar el rango de fechas
def generate_date_range(start_date, end_date):
    date_list = []
    current_date = start_date
    while current_date <= end_date:
        date_list.append(current_date)
        current_date += timedelta(days=1)
    return date_list

# Función principal para procesar y almacenar los datos
async def main():
    # Definir el rango de fechas (del 18/11/2024 al 24/11/2024)
    start_date = datetime(2024, 11, 18)
    end_date = datetime(2024, 11, 24)
    date_range = generate_date_range(start_date, end_date)

    all_data = []  # Lista para almacenar los datos de todos los días

    for date in date_range:
        # Obtener los datos de la API para esa fecha
        data = await fetch_twitch_data(date)
        if data:
            # Extraer datos relevantes del JSON (ajustar según la estructura de la API)
            extracted_data = {
                "Date": date.date(),  # Fecha específica de los datos
                "Rank": data.get("rank"),
                "Minutes Streamed": data.get("minutes_streamed"),
                "Average Viewers": data.get("avg_viwers"),  # Espectadores promedio
                "Peak Viewers": data.get("max_viewers"),  # Pico de espectadores
                "Followers": data.get("followers"),  # Total de seguidores
                "Total Followers": data.get("followers_total"),  # Total de seguidores
                "Ejecucion": datetime.now().time(),  # Hora actual de la ejecución
            }
            all_data.append(extracted_data)  # Almacenar los datos para ese día

    # Crear un DataFrame con todos los datos recopilados
    df = pd.DataFrame(all_data)

    # Guardar los datos en un archivo Excel
    file_name = "twitch_data.xlsx"

    # Verificar si el archivo ya existe
    if not os.path.exists(file_name):
        # Si el archivo no existe, se crea uno nuevo
        with pd.ExcelWriter(file_name, mode='w') as writer:
            df.to_excel(writer, index=False,header=True)  # Escribir sin índice
    else:
        # Si el archivo existe, se agrega una nueva hoja (con datos nuevos)
        with pd.ExcelWriter(file_name, mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, index=False,header=False)

    print(f"Datos guardados en {file_name}")
    
    wb = load_workbook(file_name)
    ws = wb.active
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Obtener la letra de la columna (ejemplo: "A", "B", etc.)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)  # Añadir un pequeño margen de 2
        ws.column_dimensions[column].width = adjusted_width

    # Guardamos el archivo con las celdas ajustadas
    wb.save(file_name)

# Ejecutar el script
if __name__ == "__main__":
    asyncio.run(main())
