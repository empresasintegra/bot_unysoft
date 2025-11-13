import time
import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC

def setup_driver(driver_class, service_class, chrome_options, driver_manager_instance):
    driver = driver_class(service=service_class(driver_manager_instance.install()), options=chrome_options)
    return driver, driver

def login_unysoft(driver, url, cliente, usuario, password, log):
    log("🔑 Iniciando sesión en Unysoft...")
    driver.get(url)
    driver.find_element(By.ID, "Cnx").send_keys(cliente)
    driver.find_element(By.ID, "Usuario").send_keys(usuario)
    driver.find_element(By.ID, "Contrase_a").send_keys(password)
    driver.find_element(By.CLASS_NAME, "login100-form-btn").click()

def seleccionar_empresa(driver, wait, empresa_objetivo, log):
    wait.until(EC.presence_of_element_located((By.ID, "Empresa")))
    selector_empresa = Select(driver.find_element(By.ID, "Empresa"))
    empresa_actual = selector_empresa.first_selected_option.text.strip()
    if empresa_actual != empresa_objetivo:
        selector_empresa.select_by_visible_text(empresa_objetivo)
        time.sleep(3)
    log(f"🏢 Empresa activa: {empresa_objetivo}")

def configurar_filtro_anexos(driver, wait, log):
    """Configura el filtro en la página de Anexos:
       Estado = 'Todos', Fecha Finiquito = activa, rango de fechas dinámico y buscar.
    """
    log("⚙️ Configurando filtros de búsqueda de anexos...")

    # --- Seleccionar Estado = 'Todos' ---
    wait.until(EC.presence_of_element_located((By.ID, "ListEstados")))
    select_estado = Select(driver.find_element(By.ID, "ListEstados"))
    select_estado.select_by_visible_text("Todos")
    log("✅ Estado configurado en 'Todos'.")

    # --- Calcular fechas dinámicas ---
    hoy = datetime.date.today()
    # Día 15 del mes anterior
    mes_anterior = (hoy.replace(day=15) - datetime.timedelta(days=30))
    # Día 15 del mes siguiente
    mes_siguiente = (hoy.replace(day=15) + datetime.timedelta(days=30))

    fecha_desde = mes_anterior.strftime("%d-%m-%Y")
    fecha_hasta = mes_siguiente.strftime("%d-%m-%Y")

    # --- Activar checkbox Fecha Finiquito ---
    chk_finiquito = driver.find_element(By.ID, "chkFechaFiniquito")
    if not chk_finiquito.is_selected():
        chk_finiquito.click()
        time.sleep(1)
    log("✅ Checkbox 'Fecha Finiquito' activado.")

    # --- Esperar a que se muestren los campos de fecha ---
    wait.until(EC.visibility_of_element_located((By.ID, "FechasFiniquito")))

    # --- Ingresar fechas ---
    input_desde = driver.find_element(By.ID, "FechaFiniquitoDesde")
    input_hasta = driver.find_element(By.ID, "FechaFiniquitoHasta")

    input_desde.clear()
    input_desde.send_keys(fecha_desde)
    input_hasta.clear()
    input_hasta.send_keys(fecha_hasta)
    log(f"📅 Rango de fechas configurado: {fecha_desde} → {fecha_hasta}")

    # --- Hacer clic en el botón Buscar ---
    btn_buscar = driver.find_element(By.ID, "btnBuscar")
    btn_buscar.click()
    log("🔎 Ejecutando búsqueda inicial...")
    time.sleep(3)

def buscar_nic(driver, wait, nic, log):
    log(f"🔍 Buscando NIC: {nic}")
    search = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#TablaTrabajadores_filter input")))
    search.clear()
    search.send_keys(nic)
    time.sleep(2)
    wait.until(lambda d: any(nic in r.text for r in d.find_elements(By.CSS_SELECTOR, "#TablaTrabajadores tbody tr")))

def seleccionar_fila_trabajador(driver, nic):
    tabla = driver.find_element(By.ID, "TablaTrabajadores")
    filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")
    for fila in filas:
        celdas = fila.find_elements(By.TAG_NAME, "td")
        if len(celdas) > 0 and celdas[0].text.strip() == nic:
            fila.click()
            time.sleep(2)
            return True
    return False

def cerrar_sesion(driver, log):
    try:
        driver.get("https://www.unysofterp.cl/Login/Logout")
        time.sleep(2)
        log("🔒 Sesión cerrada.")
    except Exception as e:
        log(f"⚠️ Error al cerrar sesión: {e}")
