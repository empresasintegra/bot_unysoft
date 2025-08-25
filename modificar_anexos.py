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
df = pd.read_excel("anexos_modificar.xlsx")

# Crear archivo de log
log_file = open("log_modificar_anexos.txt", "w", encoding="utf-8")
def log(mensaje):
    print(mensaje)
    log_file.write(mensaje + "\n")

def logout(driver):
    try:
        driver.get("https://www.unysofterp.cl/Login/Logout")
        time.sleep(2)
        log("✅ Sesión cerrada correctamente.")
    except Exception as e:
        log(f"⚠️ Error al cerrar sesión: {e}")

# Configurar navegador
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 15)

try:
    log("🔑 Iniciando sesión en Unysoft...")
    driver.get(URL)

    driver.find_element(By.ID, "Cnx").send_keys(CLIENTE)
    driver.find_element(By.ID, "Usuario").send_keys(USUARIO)
    driver.find_element(By.ID, "Contrase_a").send_keys(PASSWORD)
    driver.find_element(By.CLASS_NAME, "login100-form-btn").click()

    wait.until(EC.presence_of_element_located((By.ID, "Empresa")))
    selector_empresa = Select(driver.find_element(By.ID, "Empresa"))
    empresa_actual = selector_empresa.first_selected_option.text.strip()

    if empresa_actual != EMPRESA_OPERATIVA:
        selector_empresa.select_by_visible_text(EMPRESA_OPERATIVA)
        time.sleep(3)
    log(f"✅ Empresa seleccionada: {EMPRESA_OPERATIVA}")

    # Ir a página de Anexos
    driver.get("https://www.unysofterp.cl/UnyRem/Anexo")
    wait.until(EC.presence_of_element_located((By.ID, "TablaTrabajadores")))
    log("📂 Página de Anexos cargada.")

    for idx, row in df.iterrows():
        nic = str(row["NIC"]).strip()
        nueva_fecha_termino = pd.to_datetime(row["Fecha Término"]).strftime("%d-%m-%Y")
        log(f"🔍 Buscando NIC: {nic}")

        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#TablaTrabajadores_filter input")))
        search_input.clear()
        search_input.send_keys(nic)
        time.sleep(2)

        try:
            wait.until(lambda d: any(nic in r.text for r in d.find_elements(By.CSS_SELECTOR, "#TablaTrabajadores tbody tr")))
        except:
            log(f"❌ NIC {nic} no encontrado.")
            continue

        # Hacer clic en la fila del trabajador
        tabla = driver.find_element(By.ID, "TablaTrabajadores")
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")
        fila_encontrada = None

        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if len(celdas) > 0 and celdas[0].text.strip() == nic:
                fila_encontrada = fila
                break

        if not fila_encontrada:
            log(f"⚠️ No se encontró la fila del trabajador para NIC {nic}.")
            continue

        # Clic en la fila para cargar Anexos
        fila_encontrada.click()
        time.sleep(2)

        # Esperar a que aparezca la tabla de Anexos
        try:
            tabla_anexos = wait.until(EC.presence_of_element_located((By.ID, "TablaAnexos")))
            filas_anexos = tabla_anexos.find_elements(By.CSS_SELECTOR, "tbody tr")

            if not filas_anexos:
                log(f"⚠️ NIC {nic} no tiene anexos registrados.")
                continue

            # Editar el primer anexo
            boton_editar = filas_anexos[0].find_element(By.CLASS_NAME, "btn-edit")
            boton_editar.click()

            # Esperar a que cargue el formulario modal
            wait.until(EC.visibility_of_element_located((By.ID, "formularioAnexo")))

            # Activar checkbox si es necesario
            chk_fecha = driver.find_element(By.ID, "chkFechaTerminoAnexo")
            if not chk_fecha.is_selected():
                chk_fecha.click()
                time.sleep(0.5)

            # Cambiar Fecha Término
            input_fecha = driver.find_element(By.ID, "FechaTerminoAnexo")
            input_fecha.clear()
            input_fecha.send_keys(nueva_fecha_termino)

            # Guardar
            driver.find_element(By.ID, "btnGuardarAnexo").click()
            time.sleep(2)

            log(f"✅ Fecha modificada para NIC {nic} → {nueva_fecha_termino}")

        except Exception as e:
            log(f"❌ Error al modificar anexo para NIC {nic}: {e}")
            driver.save_screenshot(f"error_modificar_{nic}.png")

    log("\n🎉 Proceso finalizado correctamente.")

except Exception as e:
    log(f"\n❌ Error general: {e}")
finally:
    logout(driver)
    driver.quit()
    log_file.close()
