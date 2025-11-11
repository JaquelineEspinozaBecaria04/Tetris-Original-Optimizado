# backend/main.py

import uvicorn
import shutil
import tempfile
import os
import asyncio
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles  
from fastapi.middleware.cors import CORSMiddleware
from typing import List
from . import logic  

# 1. Inicializar la aplicación FastAPI
app = FastAPI(title="API del Generador Tetris")

# 2. Configurar CORS
# Aunque servimos todo desde el mismo dominio, es una buena
# práctica tenerlo por si el frontend crece o cambia.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================================================
# SECCIÓN 1: ENDPOINT DE LA API
# ==========================================================

@app.post("/process-tetris/")
async def handle_tetris_processing(
    # QUÉ HACE: Define la ruta de la API.
    # CÓMO FUNCIONA: Solo acepta archivos .txt subidos por el usuario.
    # Los otros archivos (diccionario, catálogo) ya están en el backend.
    #txt_file: UploadFile = File(...) # Para un solo archivo
    txt_files: List[UploadFile] = File(...) # Para múltiples archivos
):
    """
    Este endpoint recibe los archivos .txt, los guarda en una carpeta
    temporal y llama a la función principal del backend para
    procesarlos. Luego devuelve los resultados al frontend.
    """
    
    # 1. Crear un directorio temporal seguro
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # 2. Guardar los .txt subidos en el directorio temporal
            print(f"Guardando {len(txt_files)} archivos .txt en {temp_dir}")
            
            site_name = "Sitio" # Valor por defecto

            for i, txt_file in enumerate(txt_files): 
                relative_path = os.path.normpath(txt_file.filename)

                # Extraer el nombre del sitio (carpeta) del primer archivo
                if i == 0:
                    folder_name = os.path.split(relative_path)[0]
                    if folder_name:
                        site_name = folder_name

                if not relative_path.endswith(".txt"):
                    print(f"Omitiendo archivo no .txt: {relative_path}")
                    continue
                # 1. Crear la ruta de destino completa
                txt_path = os.path.join(temp_dir, relative_path)
                
                # 2. Obtener el nombre del directorio donde irá el archivo
                txt_dir = os.path.dirname(txt_path)
                
                # 3. ¡Crear ese directorio (y padres) si no existe!
                os.makedirs(txt_dir, exist_ok=True)
                # 4. Guarda el archivo
                with open(txt_path, "wb") as f:
                    shutil.copyfileobj(txt_file.file, f)
            print(f"Archivos .txt guardados. Nombre del sitio detectado: {site_name}")
            

            # 3. Ejecutar la lógica pesada en un hilo separado
            # CÓMO FUNCIONA: asyncio.to_thread evita que el servidor
            # se bloquee mientras pandas y plotly trabajan.
            print("Iniciando procesamiento en hilo de fondo...")
            html_results = await asyncio.to_thread(
                logic.generar_reportes_tetris, 
                temp_dir,
                site_name  # Pasa el nombre del sitio a la lógica
            )
            print("Procesamiento completado.")

            # 4. Devolver los resultados al frontend
            return JSONResponse(
                content={"results": html_results}
            )

        except Exception as e:
            # 5. Manejo de errores
            print(f"¡ERROR! {str(e)}")
            raise HTTPException(
                status_code=500, 
                detail=f"Ocurrió un error interno: {str(e)}"
            )

# ==========================================================
# SECCIÓN 2: SERVIR EL FRONTEND
# ==========================================================

# QUÉ HACE: Esta línea une el backend y el frontend.
# CÓMO FUNCIONA:
# 1. "Monta" la carpeta 'frontend' en la ruta raíz "/".
# 2. html=True le dice que sirva 'index.html' si alguien visita "/".
# 3. Automáticamente también servirá 'style.css', 'app.js', 'logot.jpg', etc.
# IMPORTANTE: Debe ir DESPUÉS de las rutas de tu API.
app.mount("/", StaticFiles(directory="frontend", html=True), name="static")

