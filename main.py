import time
import smtplib
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def procesar_excel():
    
    try:
        archivo_excel = "Base Seguimiento Observ Auditoría al_30042021.xlsx"

        # Abrir el archivo Excel proporcionado
        wb = load_workbook(archivo_excel)
        sheet = wb.active
        
        # Recorrer las filas del archivo Excel
        for row in sheet.iter_rows(min_row=2):
            try:
                # Obtener los valores de las celdas
                proceso = row[0].value
                observacion = row[1].value
                tipo_riesgo = row[2].value
                severidad = row[3].value
                plan_accion = row[4].value
                fecha_compromiso = row[5].value
                responsable = row[6].value
                responsable_area = row[7].value
                correo_responsable = row[8].value
                estado = row[9].value
                
                # Si el proceso está regularizado, subir información al formulario
                if estado == 'Regularizado':
                    subir_informacion(proceso, tipo_riesgo, severidad, responsable, fecha_compromiso, observacion)
                
                # Si el proceso está atrasado, enviar correo al responsable
                elif estado == 'Atrasado':
                    enviar_correo(proceso, estado, observacion, fecha_compromiso, correo_responsable)
                
            except Exception as e:
                print(f"Error al procesar fila: {e}")
        
    except Exception as e:
        print(f"Error al abrir el archivo Excel: {e}")
        
    finally:
        try:
            wb.close()
        except UnboundLocalError:
            pass

def subir_informacion(proceso, tipo_riesgo, severidad, responsable, fecha_compromiso, observacion):
    try:
        # Iniciar el navegador Chrome
        driver = webdriver.Chrome()
        driver.get("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG")
        
        # Completar el formulario con la información
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "process")))
        
        #1
        # Convertir proceso a minúsculas ya que se generan conflictos con los desplegables del formulario
        proceso = proceso.lower()
        select_element = driver.find_element(By.ID, 'process')
        select = Select(select_element)
        
        for option in select.options:
            if option.text.lower() == proceso:
                option.click()
                break
            
        #2
        driver.find_element(By.ID, 'tipo_riesgo').send_keys(tipo_riesgo)
        
        #3
        # Elimina los espacios en blanco ya que genera conflictos con el desplegable del formulario
        severidad_texto = severidad.strip()
        severidad_dato = driver.find_element(By.ID, 'severidad')
        severidad_select = Select(severidad_dato)
        severidad_select.select_by_visible_text(severidad_texto)
        
        #4
        driver.find_element(By.ID, 'res').send_keys(responsable)
        #5
        fecha_formateada = fecha_compromiso.strftime('%d/%m/%Y')
        driver.find_element(By.ID, 'date').send_keys(fecha_formateada)
        #6
        driver.find_element(By.ID, 'obs').send_keys(observacion)
        
        time.sleep(4)
        # Click en submit
        driver.find_element(By.ID, 'submit').click()
        
        # Verificación de envío de formulario
        elemento_alerta = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "alert-success")))

        # Extraer el texto de la alerta
        texto_alerta = elemento_alerta.text

        # Verificar si el texto contiene "Data sent"
        if "Data sent" in texto_alerta:
            # Extraer el ID de la cola
            id_cola = texto_alerta.split("Queue ID ")[1]
            print("Se ha enviado la información correctamente. ID de la cola:", id_cola)
        else:
            print("No se ha enviado la información correctamente.")
        
    except Exception as e:
        print(f"Error al subir la información al formulario: {e}")
        
    finally:
        # Cerrar el navegador después de 5 segundos
        time.sleep(5)
        driver.quit()

def enviar_correo(proceso, estado, observacion, fecha_compromiso ,correo_responsable):
    
    try:
        # Configurar los parámetros del correo
        remitente = 'Tu email' ############ reemplaza con tu direccion de Gmail
        destinatario = correo_responsable
        asunto = f'Estado de proceso: {proceso}'
        cuerpo = f"Estado del proceso: {estado}\nObservación: {observacion}\nFecha de compromiso: {fecha_compromiso.strftime('%d/%m/%Y')}"
        
        # Crear el mensaje de correo
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = destinatario
        mensaje['Subject'] = asunto
        mensaje.attach(MIMEText(cuerpo, 'plain'))
        
        # Iniciar conexión con el servidor SMTP de Gmail
        servidor_smtp = smtplib.SMTP('smtp.gmail.com', 587)
        servidor_smtp.starttls()
        
        # Autenticarse con el servidor SMTP
        servidor_smtp.login(remitente, 'tu password') ########### aqui se tiene que poner el password para aplicaciones de terceros de Gmail
        
        # Enviar el correo electrónico
        servidor_smtp.send_message(mensaje)
        
        # Cerrar la conexión con el servidor SMTP
        servidor_smtp.quit()
        
        print("Correo enviado exitosamente a", destinatario)
    except Exception as e:
        print(f"Error al enviar el correo electrónico: {e}")

if __name__ == "__main__":
    procesar_excel()
    print("El programa ha finalizado la carga de formularios y envio de emails")
