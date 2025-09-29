#!/usr/bin/env python3
"""
Script para crear ejecutable de Windows desde cualquier sistema operativo.
Este script debe ejecutarse en Windows para crear el ejecutable .exe
"""

import subprocess
import sys
import os

def instalar_dependencias():
    """Instala las dependencias necesarias"""
    print("🔧 Instalando dependencias...")
    dependencias = [
        'pandas',
        'openpyxl', 
        'pyinstaller'
    ]
    
    for dep in dependencias:
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', dep])
            print(f"✅ {dep} instalado correctamente")
        except subprocess.CalledProcessError:
            print(f"❌ Error instalando {dep}")
            return False
    return True

def crear_ejecutable():
    """Crea el ejecutable usando PyInstaller"""
    print("🚀 Creando ejecutable...")
    
    # Comandos para crear diferentes tipos de ejecutables
    comandos = [
        # Ejecutable simple con ventana
        [
            'pyinstaller',
            '--onefile',
            '--windowed', 
            '--name', 'ProcesadorLotes',
            '--distpath', './dist_windows',
            '--workpath', './build_windows',
            'procesar_lotes.py'
        ],
        # Ejecutable con consola (para ver mensajes de debug)
        [
            'pyinstaller',
            '--onefile',
            '--console',
            '--name', 'ProcesadorLotes_Debug',
            '--distpath', './dist_windows',
            '--workpath', './build_windows',
            'procesar_lotes.py'
        ]
    ]
    
    for i, cmd in enumerate(comandos):
        print(f"\n📦 Creando ejecutable {i+1}/2...")
        try:
            subprocess.check_call(cmd)
            print(f"✅ Ejecutable {i+1} creado correctamente")
        except subprocess.CalledProcessError as e:
            print(f"❌ Error creando ejecutable {i+1}: {e}")
            return False
    
    return True

def main():
    print("🔨 CREADOR DE EJECUTABLE PARA WINDOWS")
    print("=" * 50)
    
    # Verificar que estamos en Windows
    if os.name != 'nt':
        print("⚠️  ADVERTENCIA: Este script está diseñado para ejecutarse en Windows")
        print("   El ejecutable resultante solo funcionará en Windows")
        respuesta = input("¿Continuar de todos modos? (y/n): ")
        if respuesta.lower() != 'y':
            return
    
    # Verificar que existe el archivo fuente
    if not os.path.exists('procesar_lotes.py'):
        print("❌ Error: No se encontró el archivo 'procesar_lotes.py'")
        print("   Asegúrate de ejecutar este script en el mismo directorio que procesar_lotes.py")
        return
    
    # Instalar dependencias
    if not instalar_dependencias():
        print("❌ Error instalando dependencias. Abortando.")
        return
    
    # Crear ejecutable
    if not crear_ejecutable():
        print("❌ Error creando ejecutable. Abortando.")
        return
    
    print("\n🎉 ¡PROCESO COMPLETADO!")
    print("📁 Los ejecutables se encuentran en: ./dist_windows/")
    print("   • ProcesadorLotes.exe - Versión con interfaz gráfica")
    print("   • ProcesadorLotes_Debug.exe - Versión con consola (para debugging)")
    print("\n📋 INSTRUCCIONES DE USO:")
    print("   1. Copia cualquiera de los .exe a la computadora del cliente")
    print("   2. El cliente solo necesita hacer doble clic en el .exe")
    print("   3. No es necesario instalar Python ni otras dependencias")

if __name__ == "__main__":
    main()
