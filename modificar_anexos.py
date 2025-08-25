import os
import time
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC

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

df = pd.read_excel("anexos_modificar.xlsx")
log_file = open("log_modificar_anexos.txt", "w", encoding="utf-8")

def log(msg):
    print(msg)
    log_file.write(msg + "\n")

def editar_primer_anexo(driver, wait, nueva_fecha, nic, log):
    try:
        tabla_anexos = wait.until(EC.presence_of_element_located((By.ID, "TablaAnexos")))
        filas = tabla_anexos.find_elements(By.CSS_SELECTOR, "tbody tr")
        if not filas:
            log(f"⚠️ NIC {nic} no tiene anexos.")
            return False

        filas[0].find_element(By.CLASS_NAME, "btn-edit").click()
        wait.until(EC.visibility_of_element_located((By.ID, "formularioAnexo")))

        chk = driver.find_element(By.ID, "chkFechaTerminoAnexo")
        if not chk.is_selected():
            chk.click()
            time.sleep(0.5)

        campo_fecha = driver.find_element(By.ID, "FechaTerminoAnexo")
        campo_fecha.clear()
        campo_fecha.send_keys(nueva_fecha)

        driver.find_element(By.ID, "btnGuardarAnexo").click()
        time.sleep(2)

        log(f"✅ Fecha modificada para NIC {nic} → {nueva_fecha}")
        return True
    except Exception as e:
        log(f"❌ Error al editar anexo para NIC {nic}: {e}")
        driver.save_screenshot(f"error_modificar_{nic}.png")
        return False

# ========== EJECUCIÓN ==========
options = Options()
options.add_argument("--start-maximized")
# driver, raw_driver = setup_driver(webdriver.Chrome, Service, options, ChromeDriverManager)
driver, raw_driver = setup_driver(webdriver.Chrome, Service, options, ChromeDriverManager())
wait = WebDriverWait(driver, 15)

try:
    log("🚀 Iniciando bot de modificación de anexos...")
    login_unysoft(driver, URL, CLIENTE, USUARIO, PASSWORD, log)
    seleccionar_empresa(driver, wait, EMPRESA_OPERATIVA, log)

    driver.get("https://www.unysofterp.cl/UnyRem/Anexo")
    wait.until(EC.presence_of_element_located((By.ID, "TablaTrabajadores")))

    for idx, row in df.iterrows():
        nic = str(row["NIC"]).strip()
        nueva_fecha = pd.to_datetime(row["Fecha Término"]).strftime("%d-%m-%Y")

        try:
            buscar_nic(driver, wait, nic, log)
            if seleccionar_fila_trabajador(driver, nic):
                editar_primer_anexo(driver, wait, nueva_fecha, nic, log)
            else:
                log(f"⚠️ No se pudo seleccionar la fila del trabajador con NIC {nic}")
        except Exception as e:
            log(f"❌ Error general con NIC {nic}: {e}")
            driver.save_screenshot(f"error_general_{nic}.png")

except Exception as e:
    log(f"🚨 Error fatal: {e}")
finally:
    cerrar_sesion(driver, log)
    driver.quit()
    log_file.close()
