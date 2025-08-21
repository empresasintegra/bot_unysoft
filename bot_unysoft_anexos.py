import os
import time
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Cargar variables de entorno
load_dotenv()
URL = os.getenv("UNYSOFT_URL")
CLIENTE = os.getenv("UNYSOFT_CLIENTE")
USUARIO = os.getenv("UNYSOFT_USUARIO")
PASSWORD = os.getenv("UNYSOFT_PASSWORD")
EMPRESA_OPERATIVA = os.getenv("EMPRESA_OPERATIVA")

# Cargar datos desde Excel
df = pd.read_excel("anexos.xlsx")

# Crear archivo de log
log_file = open("log_anexos.txt", "w", encoding="utf-8")
log_file.write("📋 Registro de Anexos\n\n")

# Configurar navegador
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 15)

def log(mensaje):
    print(mensaje)
    log_file.write(mensaje + "\n")

def logout(driver):
    try:
        log("🔒 Cerrando sesión...")
        driver.get("https://www.unysofterp.cl/Login/Logout")
        time.sleep(2)
        log("✅ Sesión cerrada correctamente.")
    except Exception as e:
        log(f"⚠️ No se pudo cerrar sesión automáticamente: {e}")

try:
    log("🔑 Accediendo a login...")
    driver.get(URL)

    driver.find_element(By.ID, "Cnx").send_keys(CLIENTE)
    driver.find_element(By.ID, "Usuario").send_keys(USUARIO)
    driver.find_element(By.ID, "Contrase_a").send_keys(PASSWORD)

    log("🚀 Haciendo click en botón de login...")
    driver.find_element(By.CLASS_NAME, "login100-form-btn").click()

    wait.until(EC.presence_of_element_located((By.ID, "Empresa")))
    selector_empresa = Select(driver.find_element(By.ID, "Empresa"))
    empresa_actual = selector_empresa.first_selected_option.text.strip()

    log(f"✅ Login y acceso a empresa CLIENTE '{CLIENTE}' completado.")

    if empresa_actual != EMPRESA_OPERATIVA:
        log(f"🔄 Cambiando a EMPRESA_OPERATIVA: {EMPRESA_OPERATIVA}")
        selector_empresa.select_by_visible_text(EMPRESA_OPERATIVA)
        time.sleep(3)
    else:
        log(f"✅ EMPRESA_OPERATIVA ya seleccionada: {empresa_actual}")

    driver.get("https://www.unysofterp.cl/UnyRem")
    wait.until(EC.presence_of_element_located((By.ID, "Empresa")))
    selector_empresa = Select(driver.find_element(By.ID, "Empresa"))

    driver.get("https://www.unysofterp.cl/UnyRem/Anexo")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    log("📂 Página de Anexos cargada.")

    for idx, row in df.iterrows():
        nic = str(row["NIC"]).strip()
        titulo = str(row["Título"]).strip()
        fecha_anexo = pd.to_datetime(row["Fecha Anexo"]).strftime("%d-%m-%Y")
        fecha_termino = pd.to_datetime(row["Fecha Término"]).strftime("%d-%m-%Y")
        descripcion = str(row["Descripción"]).strip()

        log(f"📄 Procesando anexo para NIC: {nic}")

        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#TablaTrabajadores_filter input")))
        search_input.clear()
        search_input.send_keys(nic)
        time.sleep(2)  # Tiempo de espera inicial

        # Nueva espera: hasta que se cargue alguna fila en la tabla que contenga el NIC
        try:
            wait.until(lambda d: any(
                nic in row.text for row in d.find_elements(By.CSS_SELECTOR, "#TablaTrabajadores tbody tr")
            ))
        except:
            pass  # Si no lo encuentra, seguirá con la validación normal

        tabla = driver.find_element(By.ID, "TablaTrabajadores")
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")

        boton_anexo = None
        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if len(celdas) > 0 and celdas[0].text.strip() == nic:
                boton_anexo = fila.find_element(By.CLASS_NAME, "btn-Nuevo")
                break

        if not boton_anexo:
            screenshot_path = f"NIC_no_encontrado_{nic}.png"
            driver.save_screenshot(screenshot_path)
            log(f"❌ NIC {nic} no encontrado en la tabla. Screenshot guardado: {screenshot_path}")
            continue

        boton_anexo.click()
        wait.until(EC.visibility_of_element_located((By.ID, "formularioAnexo")))
        time.sleep(1)

        wait.until(EC.presence_of_element_located((By.ID, "TituloAnexo"))).send_keys(titulo)
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

        log(f"✅ Anexo guardado para NIC: {nic}")

    log("\n🎉 Todos los anexos fueron ingresados exitosamente.")

except Exception as e:
    log(f"\n❌ Error en la ejecución: {e}")
    logout(driver)

finally:
    logout(driver)
    driver.quit()
    log_file.close()
