# main.py
import pandas as pd
import openpyxl  # Necesario para leer formatos de Excel más recientes
import matplotlib.pyplot as plt
from twilio.rest import Client
from config import ACCOUNT_SID, AUTH_TOKEN, TWILIO_PHONE_NUMBER, DESTINATION_PHONE_NUMBER
import logging

# Configuración del logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='app.log',  # Guarda los logs en un archivo
    filemode='w'         # 'w' sobreescribe el archivo en cada ejecución, 'a' añade al final
)

def leer_datos_excel(archivo):
    """Lee datos de un archivo de Excel y devuelve un DataFrame de Pandas."""
    try:
        logging.info(f"Leyendo datos desde el archivo Excel: {archivo}")
        df = pd.read_excel(archivo)
        logging.info(f"Archivo Excel leído correctamente. Primeras filas:\n{df.head()}")
        return df
    except FileNotFoundError:
        logging.error(f"Error: El archivo '{archivo}' no fue encontrado.")
        return None
    except Exception as e:
        logging.exception(f"Error al leer el archivo Excel: {e}")
        return None

def realizar_analisis(df):
    """Realiza análisis financiero y estadístico sobre los datos."""
    if df is None:
        logging.warning("DataFrame es None, no se puede realizar el análisis.")
        return None

    try:
        # Ejemplo de análisis: Calcular el total de ventas por categoría
        ventas_por_categoria = df.groupby('Categoria')['Ventas'].sum()
        logging.info(f"Total de Ventas por Categoría:\n{ventas_por_categoria}")

        # Ejemplo de análisis: Calcular estadísticas descriptivas de las ventas
        estadisticas_ventas = df['Ventas'].describe()
        logging.info(f"Estadísticas Descriptivas de Ventas:\n{estadisticas_ventas}")

        return ventas_por_categoria, estadisticas_ventas

    except Exception as e:
        logging.exception(f"Error durante el análisis: {e}")
        return None

def generar_reporte(ventas_por_categoria, estadisticas_ventas):
    """Genera un reporte en texto basado en el análisis."""
    if ventas_por_categoria is None or estadisticas_ventas is None:
        logging.warning("No se pudo generar el reporte debido a errores en el análisis.")
        return "No se pudo generar el reporte debido a errores en el análisis."

    try:
        reporte = "--- Reporte de Ventas ---\n\n"
        reporte += "Total de Ventas por Categoría:\n"
        reporte += ventas_por_categoria.to_string() + "\n\n"
        reporte += "Estadísticas Descriptivas de Ventas:\n"
        reporte += estadisticas_ventas.to_string() + "\n"
        reporte += "\n--- Fin del Reporte ---"
        logging.info("Reporte generado correctamente.")
        return reporte
    except Exception as e:
        logging.exception(f"Error al generar el reporte: {e}")
        return "Error al generar el reporte."

def enviar_reporte_whatsapp(reporte):
    """Envía el reporte por WhatsApp usando la API de Twilio."""
    try:
        logging.info("Intentando enviar reporte por WhatsApp.")
        client = Client(ACCOUNT_SID, AUTH_TOKEN)

        message = client.messages.create(
            body=reporte,
            from_=f"whatsapp:{TWILIO_PHONE_NUMBER}", # Importante el prefijo 'whatsapp:'
            to=f"whatsapp:{DESTINATION_PHONE_NUMBER}"  # Importante el prefijo 'whatsapp:'
        )

        logging.info(f"Reporte enviado a WhatsApp. SID: {message.sid}")

    except Exception as e:
        logging.exception(f"Error al enviar el reporte por WhatsApp: {e}")

def generar_grafico(ventas_por_categoria, nombre_archivo="ventas_categoria.png"):
    """Genera un gráfico de barras de las ventas por categoría."""
    if ventas_por_categoria is None:
        logging.warning("No se pudo generar el gráfico.")
        return None

    try:
        plt.figure(figsize=(10, 6))  # Ajusta el tamaño del gráfico
        ventas_por_categoria.plot(kind='bar', color='skyblue')
        plt.title('Ventas por Categoría')
        plt.xlabel('Categoría')
        plt.ylabel('Ventas')
        plt.xticks(rotation=45, ha='right') # Rota las etiquetas del eje x para mejor legibilidad
        plt.tight_layout()  # Ajusta el diseño para que las etiquetas no se superpongan
        plt.savefig(f"reports/{nombre_archivo}")  # Guarda el gráfico en la carpeta 'reports'
        plt.close() # Cierra la figura para liberar memoria
        logging.info(f"Gráfico guardado como 'reports/{nombre_archivo}'")
        return f"reports/{nombre_archivo}"  # Retorna la ruta al archivo para poder adjuntarlo
    except Exception as e:
        logging.exception(f"Error al generar el gráfico: {e}")
        return None

def enviar_reporte_whatsapp_con_imagen(reporte, ruta_imagen):
    """Envía el reporte por WhatsApp con una imagen adjunta."""
    try:
        logging.info("Intentando enviar reporte con gráfico por WhatsApp.")
        client = Client(ACCOUNT_SID, AUTH_TOKEN)

        message = client.messages.create(
            body=reporte,
            media_url=[f"https://tu_servidor.com/{ruta_imagen}"],  # Requiere una URL accesible publicamente.
            from_=f"whatsapp:{TWILIO_PHONE_NUMBER}",
            to=f"whatsapp:{DESTINATION_PHONE_NUMBER}"
        )

        logging.info(f"Reporte y gráfico enviado a WhatsApp. SID: {message.sid}")
    except Exception as e:
        logging.exception(f"Error al enviar el reporte y gráfico por WhatsApp: {e}")


if __name__ == "__main__":
    archivo_excel = "data/Ventas/Fundamentos.xlsx"

    try:
        df = leer_datos_excel(archivo_excel)

        if df is not None:
            ventas_por_categoria, estadisticas_ventas = realizar_analisis(df)

            if ventas_por_categoria is not None and estadisticas_ventas is not None:
                reporte = generar_reporte(ventas_por_categoria, estadisticas_ventas)
                print("\nReporte generado:\n", reporte)
                logging.info(f"Reporte generado:\n{reporte}")  # Registra el reporte en el log

                ruta_grafico = generar_grafico(ventas_por_categoria)

                # Comentar o descomentar la función de envío de WhatsApp que quieras usar.
                # Si descomentas la función con imagen, recuerda configurar correctamente la URL pública.
                enviar_reporte_whatsapp(reporte) # Envia el reporte solo texto
                #enviar_reporte_whatsapp_con_imagen(reporte, ruta_grafico)  # Requiere una URL pública.

            else:
                logging.warning("No se pudo realizar el análisis o generar el reporte.")
        else:
            logging.error("No se pudo procesar el archivo Excel.")

    except Exception as e:
        logging.exception(f"Error inesperado en el flujo principal: {e}")