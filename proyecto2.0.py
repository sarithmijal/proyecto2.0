import time
import sqlite3
import psutil
import pygetwindow as gw
from fpdf import FPDF
import googlemaps


# Crear una conexión a la base de datos (o abrir si ya existe)
conn = sqlite3.connect("ventanas.db")

# Crear una tabla si no existe
conn.execute('''CREATE TABLE IF NOT EXISTS ventanas
             (ventana TEXT PRIMARY KEY NOT NULL,
             tiempo_inactividad TEXT NOT NULL,
             tiempo_total_abierta TEXT NOT NULL)''')

# Crear un cursor para interactuar con la base de datos
cursor = conn.cursor()

def obtener_lista_ventanas_en_uso():
    """Obtiene la lista de nombres de ventanas en uso."""
    return set([ventana.title for ventana in gw.getAllTitles()])


def tiempo_transcurrido(ventana):
    """Calcula el tiempo transcurrido desde que la ventana se abrió por última vez."""
    try:
        procesos = [p.info for p in psutil.process_iter(['pid', 'name', 'create_time']) if
                    ventana.lower() in p.info['name'].lower()]
        if procesos:
            tiempo_creacion = min(proceso['create_time'] for proceso in procesos)
            tiempo_total_abierta = time.time() - tiempo_creacion
            return tiempo_total_abierta
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        pass
    return float('inf')  # Retornar infinito si no se puede determinar el tiempo


def formatear_tiempo(tiempo_segundos):
    """Convierte el tiempo de segundos a un formato legible (hh:mm:ss)."""
    if not tiempo_segundos or tiempo_segundos == float('inf'):
        return "Desconocido"

    horas, segundos_restantes = divmod(tiempo_segundos, 3600)
    minutos, segundos = divmod(segundos_restantes, 60)
    return f"{int(horas):02d}:{int(minutos):02d}:{int(segundos):02d}"


def tiempo_actividad_total(ventana):
    """Calcula el tiempo total que la ventana ha estado abierta."""
    try:
        procesos = [p.info for p in psutil.process_iter(['pid', 'name', 'create_time']) if
                    ventana.lower() in p.info['name'].lower()]
        if procesos:
            tiempo_creacion = min(proceso['create_time'] for proceso in procesos)
            tiempo_total_abierta = time.time() - tiempo_creacion
            return tiempo_total_abierta
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        pass
    return float('inf')  # Retornar infinito si no se puede determinar el tiempo


def obtener_lista_ventanas_no_utilizadas(tiempo_limite=600):
    """Obtiene la lista de ventanas abiertas que no están en uso."""
    ventanas_abiertas = gw.getAllTitles()
    ventanas_en_uso = obtener_lista_ventanas_en_uso()

    # Filtrar las ventanas no utilizadas que no están vacías
    ventanas_no_utilizadas = [ventana for ventana in ventanas_abiertas if ventana and ventana != "Program Manager"]

    # Filtrar las ventanas no utilizadas que no están en la lista de ventanas en uso
    ventanas_no_utilizadas = [ventana for ventana in ventanas_no_utilizadas if ventana not in ventanas_en_uso]

    # Filtrar las ventanas no utilizadas que han estado abiertas durante más tiempo del límite establecido
    ventanas_no_utilizadas = [ventana for ventana in ventanas_no_utilizadas if
                              tiempo_transcurrido(ventana) > tiempo_limite]

    return ventanas_no_utilizadas


def sugerir_ventanas_no_utilizadas(tiempo_limite=600, ventanas_por_pagina=1):
    """Sugiere ventanas que no se han utilizado en un tiempo específico."""
    info_ventanas = []  # Lista para almacenar la información de las ventanas no utilizadas
    num_ventanas = 0  # Contador de ventanas por página

    # Agregar título y descripción al informe
    info_ventanas.append("Informe de Ventanas No Utilizadas")
    info_ventanas.append("Este informe muestra las ventanas que no han sido utilizadas recientemente")
    info_ventanas.append("\nDetalles:")

    # Obtener la lista de ventanas no utilizadas con un límite de 10 minutos
    ventanas_no_utilizadas = obtener_lista_ventanas_no_utilizadas(tiempo_limite)

    # Agregar encabezado al PDF
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Agregar título y descripción al PDF
    pdf.set_font("Arial", 'B', size=16)
    pdf.multi_cell(0, 10, txt="Informe de Ventanas No Utilizadas", align='C')
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, txt="Este informe muestra las ventanas que no han sido utilizadas recientemente", align='L')
    pdf.ln(10)  # Agregar espacio

    # Agregar fecha de creación del informe al PDF
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Fecha de creación: {time.strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align='C')
    pdf.ln(10)  # Agregar espacio

    # Agregar detalles de ventanas no utilizadas al informe y mostrar en consola
    pdf.set_font("Arial", size=12)
    for ventana in ventanas_no_utilizadas:
        tiempo_inactividad = tiempo_transcurrido(ventana)
        tiempo_total_abierta = tiempo_actividad_total(ventana)
        tiempo_formateado_inactividad = formatear_tiempo(tiempo_inactividad)
        tiempo_formateado_total_abierta = formatear_tiempo(tiempo_total_abierta)

        # Agregar detalles de ventana al PDF
        pdf.cell(0, 10, txt=f"Ventana: {ventana.encode('latin-1', 'replace').decode('latin-1')}", ln=True)
        pdf.cell(0, 10, txt=f"Tiempo de Inactividad: {tiempo_formateado_inactividad}", ln=True)
        pdf.cell(0, 10, txt=f"Tiempo Total Abierta: {tiempo_formateado_total_abierta}", ln=True)
        pdf.ln(5)  # Agregar espacio entre ventanas

        # Verificar si se debe crear una nueva página
        num_ventanas += 1
        if num_ventanas >= ventanas_por_pagina:
            pdf.add_page()  # Agregar una nueva página
            num_ventanas = 0  # Reiniciar el contador de ventanas por página

        # Insertar datos en la base de datos
        cursor.execute(
            "INSERT OR REPLACE INTO ventanas (ventana, tiempo_inactividad, tiempo_total_abierta) VALUES (?, ?, ?)",
            (ventana.encode('latin-1', 'replace').decode('latin-1'), tiempo_formateado_inactividad,
             tiempo_formateado_total_abierta))

        # Mostrar en consola
        ventana_info = f"Ventana: {ventana}\nTiempo de Inactividad: {tiempo_formateado_inactividad}\nTiempo Total Abierta: {tiempo_formateado_total_abierta}"
        info_ventanas.append(ventana_info)
        print(ventana_info)

    # Confirmar los cambios y cerrar la conexión a la base de datos
    conn.commit()
    conn.close()

    # Guardar el archivo PDF en el escritorio
    pdf_file = "informe_ventanas_no_utilizadas.pdf"
    pdf_output_path = f"C:\\Users\\HP\\Desktop\\{pdf_file}"
    pdf.output(pdf_output_path)
    print(f"Se ha creado el archivo PDF en: {pdf_output_path}")




if __name__ == "__main__":
    sugerir_ventanas_no_utilizadas()


