import os
import time
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from unysoft_utils import (
    setup_driver, login_unysoft, seleccionar_empresa, buscar_nic,
    seleccionar_fila_trabajador, cerrar_sesion
)

# ========== CONFIG ==========
load_dotenv()
URL = os.getenv("UNYSOFT_URL")
CLIENTE = os.getenv("UNYSOFT_CLIENTE")
USUARIO = os.getenv("UNYSOFT_USUARIO")
PASSWORD = os.getenv("UNYSOFT_PASSWORD")
EMPRESA_OPERATIVA = os.getenv("EMPRESA_OPERATIVA")

# Cargar datos desde Excel
df = pd.read_excel("anexos_crear.xlsx")

# Crear archivo de log
log_file = open("log_crear_anexos.txt", "w", encoding="utf-8")

def log(msg):
    print(msg)
    log_file.write(msg + "\n")

def crear_anexo(driver, wait, nic, titulo, fecha_anexo, fecha_termino, descripcion, log):
    try:
        tabla = driver.find_element(By.ID, "TablaTrabajadores")
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")

        boton_anexo = None
        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if len(celdas) > 0 and celdas[0].text.strip() == nic:
                boton_anexo = fila.find_element(By.CLASS_NAME, "btn-Nuevo")
                break

        if not boton_anexo:
            log(f"❌ NIC {nic} no encontrado en la tabla.")
            driver.save_screenshot(f"NIC_no_encontrado_{nic}.png")
            return False

        boton_anexo.click()
        wait.until(EC.visibility_of_element_located((By.ID, "formularioAnexo")))
        time.sleep(1)

        driver.find_element(By.ID, "TituloAnexo").send_keys(titulo)
        driver.find_element(By.ID, "FechaIngresoAnexo").send_keys(fecha_anexo)

        driver.find_element(By.ID, "chkFechaTerminoAnexo").click()
        time.sleep(0.5)
        fecha_input = driver.find_element(By.ID, "FechaTerminoAnexo")
        fecha_input.clear()
        fecha_input.send_keys(fecha_termino)

        driver.execute_script("document.getElementById('EsAumentoPlazo').disabled = false;")
        driver.find_element(By.ID, "EsAumentoPlazo").click()
        driver.find_element(By.ID, "GlosaAnexo").send_keys(descripcion)

        driver.find_element(By.ID, "btnGuardarAnexo").click()
        time.sleep(2)

        log(f"✅ Anexo creado para NIC {nic}")
        return True
    except Exception as e:
        log(f"❌ Error al crear anexo para NIC {nic}: {e}")
        driver.save_screenshot(f"error_crear_{nic}.png")
        return False

# ========== EJECUCIÓN ==========
options = Options()
options.add_argument("--start-maximized")
driver, raw_driver = setup_driver(webdriver.Chrome, Service, options, ChromeDriverManager())
wait = WebDriverWait(driver, 15)

try:
    log("🚀 Iniciando bot de creación de anexos...")
    login_unysoft(driver, URL, CLIENTE, USUARIO, PASSWORD, log)
    seleccionar_empresa(driver, wait, EMPRESA_OPERATIVA, log)

    driver.get("https://www.unysofterp.cl/UnyRem/Anexo")
    wait.until(EC.presence_of_element_located((By.ID, "TablaTrabajadores")))
    log("📂 Página de Anexos cargada.")

    for idx, row in df.iterrows():
        nic = str(row["NIC"]).strip()
        titulo = str(row["Título"]).strip()
        fecha_anexo = pd.to_datetime(row["Fecha Anexo"]).strftime("%d-%m-%Y")
        fecha_termino = pd.to_datetime(row["Fecha Término"]).strftime("%d-%m-%Y")
        descripcion = str(row["Descripción"]).strip()

        try:
            buscar_nic(driver, wait, nic, log)
            if seleccionar_fila_trabajador(driver, nic):
                crear_anexo(driver, wait, nic, titulo, fecha_anexo, fecha_termino, descripcion, log)
            else:
                log(f"⚠️ No se pudo seleccionar trabajador con NIC {nic}")
        except Exception as e:
            log(f"❌ Error general con NIC {nic}: {e}")
            driver.save_screenshot(f"error_general_{nic}.png")

    log("\n🎉 Todos los anexos fueron ingresados exitosamente.")

except Exception as e:
    log(f"\n🚨 Error fatal: {e}")

finally:
    cerrar_sesion(driver, log)
    driver.quit()
    log_file.close()
