# crear_ejecutable.py
import subprocess
import sys
import os
import shutil

def clean_previous_builds(name):
    """Limpia builds anteriores"""
    folders_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = [f'{name}.spec']

    for folder in folders_to_clean:
        if os.path.exists(folder):
            print(f"🧹 Limpiando carpeta: {folder}")
            shutil.rmtree(folder, ignore_errors=True)

    for file in files_to_clean:
        if os.path.exists(file):
            print(f"🧹 Eliminando archivo: {file}")
            os.remove(file)

def create_executable(main_file, exe_name):
    """Crea el archivo ejecutable"""
    print("🔨 Creando ejecutable del bot...", main_file, exe_name)
    
    # Verificar que existe el archivo principal
    if not os.path.exists(main_file):
        print(f"❌ ERROR: No se encuentra {main_file}")
        return False
    
    # Comando para crear ejecutable (SIMPLIFICADO para evitar errores)
    cmd = [
        "pyinstaller",
        "--onefile",                    # Un solo archivo .exe
        "--console",                    # Con ventana de consola para debug
        f"--name={exe_name}",           # Nombre del ejecutable
        main_file                       # Archivo principal
    ]

    try:
        subprocess.run(cmd, check=True)
        print(f"✅ Ejecutable {exe_name}.exe creado correctamente")
        return True
    except subprocess.CalledProcessError:
        print("❌ Error creando el ejecutable")
        return False

if __name__ == "__main__":
    # Selección del bot (crear o modificar)
    if len(sys.argv) < 2:
        print("❌ Uso: python crear_ejecutable.py [crear|modificar]")
        sys.exit(1)
    
    bot = sys.argv[1].lower()
    
    if bot == "crear":
        clean_previous_builds("BotCrearAnexos")
        create_executable("crear_anexos.py", "BotCrearAnexos")
    elif bot == "modificar":
        clean_previous_builds("BotModificarAnexos")
        create_executable("modificar_anexos.py", "BotModificarAnexos")
    else:
        print("❌ Opción inválida. Usa: crear o modificar")
