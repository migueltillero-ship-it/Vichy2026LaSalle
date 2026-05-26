import os
import shutil
from pathlib import Path

# Ruta exacta del Directorio Institucional
ruta_proyecto = Path(r"C:\Users\ALIANZA FRANCESA\Desktop\Vichy2026-LaSalle")

def optimizar_boveda(ruta_base):
    print("Iniciando auditoría y consolidación de la bóveda de Vichy 2026...\n")
    
    if not ruta_base.exists():
        print(f"Error crítico: No se encontró la ruta {ruta_base}")
        return

    carpeta_public = ruta_base / "public"
    
    # FASE 1: Extracción segura de recursos multimedia desde 'public'
    if carpeta_public.exists():
        print("[1] Extrayendo recursos multimedia a la raíz...")
        for archivo in carpeta_public.rglob('*'):
            if archivo.is_file():
                destino = ruta_base / archivo.name
                
                # Regla de oro: Proteger archivos maestros. Si no existe en raíz, se mueve.
                if not destino.exists():
                    shutil.move(str(archivo), str(destino))
                    print(f"  ✓ Rescatado: {archivo.name}")
                else:
                    print(f"  ⚠ Duplicado ignorado (se conserva el maestro intacto): {archivo.name}")
        
        # Eliminación de la estructura public ahora vacía
        try:
            shutil.rmtree(carpeta_public)
            print("  ✓ Carpeta 'public' eliminada exitosamente.")
        except Exception as e:
            print(f"  ⚠ Aviso: No se pudo eliminar 'public' (Asegúrese de cerrar la carpeta en Windows).")
    else:
        print("[1] Estructura óptima: La carpeta 'public' ya no existe.")

    # FASE 2: Erradicación de deuda técnica y archivos de código remanentes
    print("\n[2] Purgando carpetas y archivos estructurales obsoletos...")
    
    carpetas_obsoletas = ['app', 'components', 'styles', 'node_modules', '.next']
    archivos_obsoletos = [
        'package.json', 'package-lock.json', 'next.config.mjs', 
        'postcss.config.js', 'tailwind.config.js', 'tailwind.config.ts', 
        '.eslintrc.json', 'README.md', 'components.json'
    ]

    for carpeta in carpetas_obsoletas:
        ruta_carpeta = ruta_base / carpeta
        if ruta_carpeta.exists():
            try:
                shutil.rmtree(ruta_carpeta)
                print(f"  ✓ Directorio eliminado: {carpeta}")
            except Exception:
                pass

    for archivo in archivos_obsoletos:
        ruta_archivo = ruta_base / archivo
        if ruta_archivo.exists():
            try:
                os.remove(ruta_archivo)
                print(f"  ✓ Archivo eliminado: {archivo}")
            except Exception:
                pass

    print("\n=======================================================================")
    print(" OPERACIÓN EXITOSA: La bóveda ha sido purgada y consolidada.")
    print(" Su proyecto ahora es una plataforma estática pura y de alto rendimiento.")
    print("=======================================================================")

if __name__ == "__main__":
    optimizar_boveda(ruta_proyecto)