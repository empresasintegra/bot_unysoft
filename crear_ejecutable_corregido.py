import subprocess
import sys
import os
import shutil

def clean_previous_builds():
    """Limpia builds anteriores"""
    folders_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = ['BotAnexos.spec']
    
    for folder in folders_to_clean:
        if os.path.exists(folder):
            print(f"🧹 Limpiando carpeta: {folder}")
            shutil.rmtree(folder, ignore_errors=True)
    
    for file in files_to_clean:
        if os.path.exists(file):
            print(f"🧹 Eliminando archivo: {file}")
            os.remove(file)

def install_pyinstaller():
    """Instala PyInstaller si no está instalado"""
    try:
        import PyInstaller
        print("✅ PyInstaller ya está instalado")
    except ImportError:
        print("📦 Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

def create_executable():
    """Crea el archivo ejecutable"""
    print("🔨 Creando ejecutable del bot...")
    
    # Verificar que existe el archivo principal
    if not os.path.exists("bot_unysoft_anexos.py"):
        print("❌ ERROR: No se encuentra 'bot_unysoft_anexos.py'")
        return False
    
    # Comando para crear ejecutable (SIMPLIFICADO para evitar errores)
    cmd = [
        "pyinstaller",
        "--onefile",                    # Un solo archivo .exe
        "--console",                    # Con ventana de consola para debug
        "--name=BotAnexos",            # Nombre del ejecutable
        "bot_unysoft_anexos.py"        # Archivo principal
    ]
    
    try:
        subprocess.run(cmd, check=True)
        print("✅ Ejecutable creado exitosamente!")
        
        # Verificar que se creó correctamente
        exe_path = "dist/BotAnexos.exe"
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024*1024)
            print(f"📁 Archivo creado: {exe_path} ({size_mb:.1f} MB)")
            
            # Crear archivos auxiliares
            create_support_files()
            return True
        else:
            print("❌ No se encontró el archivo ejecutable generado")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Error creando ejecutable: {e}")
        return False

def create_support_files():
    """Crea archivos de soporte para el usuario"""
    print("📄 Creando archivos de soporte...")
    
    # Crear .env de ejemplo si no existe
    if not os.path.exists("dist/.env"):
        with open("dist/.env", "w", encoding="utf-8") as f:
            f.write('''# Configuración del Bot de Anexos UnySOFT
# IMPORTANTE: Completa estos datos antes de ejecutar el bot

UNYSOFT_URL=https://www.unysofterp.cl/Login/Login
UNYSOFT_CLIENTE=TU_CLIENTE
UNYSOFT_USUARIO=TU_USUARIO
UNYSOFT_PASSWORD=TU_PASSWORD
EMPRESA_OPERATIVA=TU_EMPRESA
''')
    
    # Crear archivo bat mejorado
    with open("dist/Ejecutar_Bot_Anexos.bat", "w", encoding="utf-8") as f:
        f.write('''@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo.
echo ==========================================
echo    🤖 BOT DE ANEXOS - UNYSOFT ERP
echo ==========================================
echo.

REM Verificar archivo .env
if not exist ".env" (
    echo ❌ ERROR: No se encuentra el archivo ".env"
    echo.
    echo Por favor configura tus credenciales en el archivo .env
    echo.
    pause
    exit /b 1
)

REM Verificar archivo Excel
if not exist "anexos.xlsx" (
    echo ❌ ERROR: No se encuentra el archivo "anexos.xlsx"
    echo.
    echo Coloca tu archivo Excel con los anexos en esta carpeta
    echo y asegúrate de que se llame exactamente "anexos.xlsx"
    echo.
    pause
    exit /b 1
)

echo ✅ Archivos verificados correctamente
echo.
echo 🚀 Iniciando bot de anexos...
echo ⏳ Esto puede tomar varios minutos, por favor espera...
echo.

BotAnexos.exe

echo.
echo ==========================================
if exist "log_anexos.txt" (
    echo ✅ Ejecución completada
    echo 📋 Revisa el archivo "log_anexos.txt" para ver los resultados
) else (
    echo ⚠️  La ejecución terminó pero no se encontró el log
)
echo.
echo Presiona cualquier tecla para cerrar...
pause >nul''')
    
    # Crear archivo README
    with open("dist/README_Usuario.txt", "w", encoding="utf-8") as f:
        f.write('''🤖 BOT DE ANEXOS UNYSOFT - MANUAL DE USUARIO
=====================================================

ARCHIVOS NECESARIOS:
--------------------
✅ BotAnexos.exe          - El programa principal
✅ .env                   - Configuración de credenciales  
✅ anexos.xlsx           - Tu archivo Excel con los datos
✅ Ejecutar_Bot_Anexos.bat - Ejecutor fácil

¡IMPORTANTE!
- No muevas ni elimines archivos durante la ejecución
- Asegúrate de tener conexión a internet estable
- Cierra otros navegadores antes de ejecutar el bot
''')
    
    # Crear archivo Excel de ejemplo
    try:
        import pandas as pd
        df_ejemplo = pd.DataFrame({
            'NIC': ['12345678-9', '98765432-1'],
            'Título': ['Anexo Ejemplo 1', 'Anexo Ejemplo 2'], 
            'Fecha Anexo': ['2024-01-15', '2024-01-20'],
            'Fecha Término': ['2024-12-31', '2024-12-31'],
            'Descripción': ['Descripción del anexo 1', 'Descripción del anexo 2']
        })
        df_ejemplo.to_excel("dist/anexos_ejemplo.xlsx", index=False)
        print("📊 Archivo Excel de ejemplo creado")
    except:
        print("⚠️ No se pudo crear el Excel de ejemplo")

if __name__ == "__main__":
    print("🚀 Generador de Ejecutable - Bot de Anexos (CORREGIDO)")
    print("=" * 55)
    
    clean_previous_builds()
    install_pyinstaller()
    
    if create_executable():
        print("\n" + "="*50)
        print("🎉 ¡EJECUTABLE CREADO EXITOSAMENTE!")
        print("="*50)
        print("\n📋 INSTRUCCIONES PARA EL USUARIO FINAL:")
        print("1. Copia toda la carpeta 'dist' al PC del administrativo")
        print("2. Dentro de 'dist', configura el archivo '.env' con las credenciales")
        print("3. Coloca tu archivo 'anexos.xlsx' dentro de 'dist'")
        print("4. Haz doble click en 'Ejecutar_Bot_Anexos.bat'")
        print("5.\n✨ ¡El bot se ejecutará automáticamente!")
    else:
        print("\n❌ Falló la creación del ejecutable")
        print("Revisa los errores mostrados arriba")