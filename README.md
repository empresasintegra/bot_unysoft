## Bot Anexos UnySOFT ERP

## ️Requisitos

* Python 3.10 o superior

* pip

* Windows (para ejecutar scripts .bat y entornos virtuales con venv)

## Instalación

1. Clona el repositorio:

```
git clone https://github.com/empresasintegra/bot_unysoft
cd bot_unysoft
```


2. Crea y activa el entorno virtual:

```
python -m venv .envs
.envs\Scripts\activate
```


3. Instala las dependencias:

`pip install -r requirements.txt`

## Crear ejecutable

Ejecuta el script para crear el ejecutable:

`python crear_ejecutable_corregido.py`

## Notas adicionales

Verifica que el ERP UnySOFT esté accesible y funcional durante la ejecución.


### Archivo necesarios .env, anexos_crear.xlsx y anexos_modificar.xlsx

<!-- 1. CONFIGURAR CREDENCIALES:
     * UNYSOFT_URL=https://www.unysofterp.cl/
     * UNYSOFT_CLIENTE=tu_cliente
     * UNYSOFT_USUARIO=tu_usuario  
     * UNYSOFT_PASSWORD=tu_contraseña
     * EMPRESA_OPERATIVA=tu_empresa

2. ARCHIVO EXCEL:
   - Tu archivo debe llamarse exactamente "anexos_crear.xlsx"
   - Columnas:
     * NIC
     * Título
     * Fecha Anexo
     * Fecha Término  
     * Descripción -->

## 📁 Estructura de archivos
bot_unysoft/
├── crear_anexos.py                 ← Crear anexos
├── modificar_anexos.py             ← Ejecuta la modificación de anexos
├── unysoft_utils.py                ← Funciones reutilizables (login, buscar, etc.)
├── anexos_crear.xlsx               ← Excel con NIC, Título, Fecha Anexo, Fecha Término y Descripción
├── anexos_modificar.xlsx           ← Excel con NIC y Fecha Término
├── log_crear_anexos.txt            ← Log generado automáticamente en la creación
├── log_modificar_anexos.txt        ← Log generado automáticamente en la modificación
└── crear_ejecutable_corregido.py   ← Crea el archivo ejecutable
