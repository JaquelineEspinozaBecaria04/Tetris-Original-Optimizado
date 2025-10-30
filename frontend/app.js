// frontend/app.js

// QUÉ HACE: Es el punto de entrada. Ejecuta el código solo
// cuando el HTML (DOM) se ha cargado completamente.
document.addEventListener('DOMContentLoaded', () => {
    
    // --- 1. Obtener referencias a elementos del HTML ---
    // CÓMO FUNCIONA: Guardamos los elementos en variables
    // para poder interactuar con ellos (leer valores, ocultar/mostrar).
    const btn = document.getElementById('process-btn');
    const txtInput = document.getElementById('txt-file');
    // Los inputs de diccionario y catálogo se eliminaron.
    
    const resultsContainer = document.getElementById('results-container');
    const errorBox = document.getElementById('error-box');
    const loadingMsg = document.getElementById('loading-msg');

    // --- 2. Escuchar el clic en el botón "Generar" ---
    // CÓMO FUNCIONA: 'async' permite usar 'await' para
    // esperar la respuesta del servidor sin congelar el navegador.
    btn.addEventListener('click', async () => {
        
        // --- 3. Validación (simplificada) ---
        // QUÉ HACE: Verifica que el usuario haya seleccionado archivos.
        if (txtInput.files.length === 0) {
            showError("Por favor, selecciona un archivo .txt.");
            return; 
        }

        // --- 4. Preparar FormData ---
        // QUÉ HACE: FormData es el "paquete" que enviamos al
        // backend. Es la forma de enviar archivos.
        const formData = new FormData();
        
        // CÓMO FUNCIONA: Recorre la lista de archivos .txt
        // y los añade todos al "paquete" con la misma clave.
        // FastAPI entenderá esto como una Lista.
        formData.append('txt_file', txtInput.files[0]);

        // --- 5. Actualizar UI a "Cargando" ---
        showLoading(true); // Muestra "Procesando..."
        showError(null); // Oculta errores antiguos
        resultsContainer.innerHTML = ''; // Limpia resultados antiguos
        btn.disabled = true; // Deshabilita el botón

        // --- 6. Llamar al API (Backend) ---
        try {
            // CÓMO FUNCIONA: fetch() hace la petición de red.
            // 'await' pausa la función aquí hasta que el servidor responda.
            // La URL '/process-tetris/' es relativa (al mismo servidor).
            const response = await fetch('/process-tetris/', {
                method: 'POST',
                body: formData 
                // No se pone 'Content-Type', el navegador lo hace solo.
            });

            // Si el servidor devuelve un error (ej. 500)
            if (!response.ok) {
                const errData = await response.json();
                // Lanza un error para ser capturado por el 'catch'
                throw new Error(errData.detail || "Ocurrió un error en el servidor.");
            }

            // ¡Éxito! Leer la respuesta JSON
            const data = await response.json(); // Espera: { results: [...] }
            displayResults(data.results); // Dibuja los resultados

        } catch (error) {
            // --- 7. Manejo de Errores ---
            // QUÉ HACE: Captura errores de red o del 'throw' anterior.
            console.error('Error en la petición fetch:', error);
            showError(error.message); // Muestra el error al usuario
        } finally {
            // --- 8. Limpiar Estado (siempre se ejecuta) ---
            showLoading(false); // Oculta "Procesando..."
            btn.disabled = false; // Vuelve a habilitar el botón
        }
    });

    // --- Funciones de Ayuda ---

    /**
     * QUÉ HACE: Dibuja los resultados (links de descarga e iframes).
     * CÓMO FUNCIONA: Itera la lista de resultados. Por cada uno,
     * crea un 'Blob' (archivo en memoria), genera una URL
     * para él, y usa esa URL tanto para un link de descarga <a>
     * como para un iframe de vista previa.
     */
    function displayResults(results) {
        resultsContainer.innerHTML = ''; // Limpiar
        if (results.length === 0) {
            resultsContainer.innerHTML = '<p class="hint">Proceso completado, no se generaron gráficos.</p>';
            return;
        }
        
        results.forEach(file => {
            // 1. Crear archivo en memoria
            const blob = new Blob([file.content], { type: 'text/html' });
            // 2. Crear URL para ese archivo
            const url = URL.createObjectURL(blob);
            
            // 3. Crear link de descarga
            const link = document.createElement('a');
            link.href = url;
            link.download = file.filename;
            link.textContent = `Descargar ${file.filename}`;
            link.className = 'btn success';
            link.style.margin = '5px';
            resultsContainer.appendChild(link);
            
            // 4. Crear vista previa en iframe
            const iframeWrapper = document.createElement('div');
            iframeWrapper.className = 'layout html-wrapper'; 
            const iframe = document.createElement('iframe');
            iframe.className = 'html-frame';
            iframe.src = url; 
            iframe.style.height = '600px'; 
            iframe.style.marginTop = '15px';
            iframe.style.marginBottom = '15px';
            iframeWrapper.appendChild(iframe);
            resultsContainer.appendChild(iframeWrapper);
        });
    }

    /** Muestra/Oculta el mensaje de error */
    function showError(message) {
        if (message) {
            errorBox.textContent = message;
            errorBox.style.display = 'block';
        } else {
            errorBox.style.display = 'none';
        }
    }

    /** Muestra/Oculta el mensaje de carga */
    function showLoading(isLoading) {
        if (loadingMsg) {
            loadingMsg.style.display = isLoading ? 'block' : 'none';
        }
    }
});