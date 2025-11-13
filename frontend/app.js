// frontend/app.js

// --- INICIO: BLOQUE DE FUNCIONES AUXILIARES FALTANTES ---
// Estas funciones son necesarias para renderizar las estadísticas
const fmtPct = (x) => (isFinite(x) ? (x * 100).toFixed(1) + "%" : "N/A");
const safe = (x) => (x === undefined || x === null ? "N/A" : x);
// --- FIN: BLOQUE DE FUNCIONES AUXILIARES FALTANTES ---


document.addEventListener('DOMContentLoaded', () => {
    
    // --- 1. Obtener referencias a elementos del HTML ---
    const btn = document.getElementById('process-btn');
    const txtInput = document.getElementById('txt-files'); // ID plural
    
    const resultsContainer = document.getElementById('results-container');
    const errorBox = document.getElementById('error-box');
    const loadingMsg = document.getElementById('loading-msg');

    // --- 2. Escuchar el clic en el botón "Generar" ---
    btn.addEventListener('click', async () => {
        
        // --- 3. Validación ---
        if (txtInput.files.length === 0) {
            showError("Por favor, selecciona los archivos .txt o la carpeta.");
            return; 
        }

        // --- 4. Preparar FormData ---
        const formData = new FormData();
        for (const file of txtInput.files) {
            formData.append('txt_files', file);
        }

        // --- 5. Actualizar UI a "Cargando" ---
        showLoading(true); 
        showError(null); 
        resultsContainer.innerHTML = ''; 
        btn.disabled = true; 

        // --- 6. Llamar al API (Backend) ---
        try {
            const response = await fetch('/process-tetris/', {
                method: 'POST',
                body: formData 
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || "Ocurrió un error en el servidor.");
            }

            const data = await response.json(); 
            displayResults(data.results); 

        } catch (error) {
            console.error('Error en la petición fetch:', error);
            showError(error.message);
        } finally {
            showLoading(false); 
            btn.disabled = false; 
        }
    });

    // --- Funciones de Ayuda ---

    /**
     * QUÉ HACE: Dibuja los resultados, agrupando botones
     * y mostrando iframes por separado.
     */
    function displayResults(results) {
        resultsContainer.innerHTML = ''; // Limpiar
        if (results.length === 0) {
            resultsContainer.innerHTML = '<p class="hint">Proceso completado, no se generaron gráficos.</p>';
            return;
        }

        // 1. Agrupar archivos por nombre base (ej. "Sitio_tetris_CEE1")
        const groupedResults = {};
        results.forEach(file => {
            // Elimina .html, .csv o .xlsx para obtener el nombre base
            const baseName = file.filename.replace(/(_stats)?\.(html|csv|xlsx|json)$/, '');
            if (!groupedResults[baseName]) {
                groupedResults[baseName] = {};
            }
            if (file.filename.endsWith('.html')) {
                groupedResults[baseName].html = file;
            } else if (file.filename.endsWith('.csv')) { // Detectar archivo con formato csv
                groupedResults[baseName].csv = file;
            } else if (file.filename.endsWith('.xlsx')) { // Detectar archivo con formato xlsx
                groupedResults[baseName].xlsx = file;
            } else if (file.filename.endsWith('_stats.json')) { // Detectar el json de las estadísticas
                groupedResults[baseName].stats = file;
            }
        });

        // 2. Iterar sobre cada grupo y crear los elementos
        for (const baseName in groupedResults) {
            const group = groupedResults[baseName];

            // Crear un contenedor para este grupo de resultados
            const groupContainer = document.createElement('div');
            groupContainer.style.marginBottom = '20px';

            if (group.stats) {
                try {
                    const s = JSON.parse(group.stats.content);
                    
                    // Usamos la clase 'stats' que ya tiene el estilo grid/flex oscuro en tu CSS
                    const statsContainer = document.createElement('div');
                    statsContainer.className = 'stats'; 
                    
                    // Función interna para crear cada tarjeta usando la clase 'stat'
                    const createStatCard = (label, value) => {
                        const card = document.createElement('div');
                        card.className = 'stat'; 
                        // Eliminamos los estilos inline de background y color para que tome el CSS oscuro
                        card.innerHTML = `<b>${label}</b><div>${value}</div>`;
                        return card;
                    };

                    // Añadir tarjetas
                    statsContainer.appendChild(createStatCard('Estado', safe(s.status)));
                    statsContainer.appendChild(createStatCard('VMs', safe(s.n_vms)));
                    statsContainer.appendChild(createStatCard('Hosts', safe(s.hosts)));
                    statsContainer.appendChild(createStatCard('Capacidad', safe(s.total_capacity)));
                    statsContainer.appendChild(createStatCard('Usado', safe(s.total_used)));
                    statsContainer.appendChild(createStatCard('Utilización', fmtPct(s.utilization)));
                    statsContainer.appendChild(createStatCard('Sin Uso', fmtPct(s.empty_pct)));
                    
                    if (s.chips_used !== undefined) {
                        statsContainer.appendChild(createStatCard('Chips usados', safe(s.chips_used)));
                    }
                    
                    if (s.host_utilization) {
                        // Formato compacto para la tarjeta de utilización
                        const hostUtilHTML = `Prom: ${fmtPct(s.host_utilization.avg)}<br>Max: ${fmtPct(s.host_utilization.max)}<br>Min: ${fmtPct(s.host_utilization.min)}`;
                        statsContainer.appendChild(createStatCard('Utilización/host', hostUtilHTML));
                    }
                    
                    // AZ Detectadas (Si son muchas, el CSS debe manejar el wrap)
                    if (s.az_values && s.az_values.length) {
                        statsContainer.appendChild(createStatCard('AZ detectadas', s.az_values.join(", ")));
                    }
                    
                    // Detalles por AZ
                    if (s.per_az) {
                        for (const [az, v] of Object.entries(s.per_az)) {
                            const azHTML = `hosts: ${safe(v.hosts)}<br>utilidad: ${fmtPct(v.utilization)}`;
                            statsContainer.appendChild(createStatCard(az, azHTML));
                        }
                    }
                    
                    groupContainer.appendChild(statsContainer);

                } catch (e) {
                    console.error("Error al renderizar estadísticas:", e);
                }
            }
            
            // Crear un contenedor para los botones
            const buttonContainer = document.createElement('div');
            buttonContainer.className = 'row'; // Usa tu clase CSS
            
            let html_url = null; // Guardar la URL del blob HTML

            // 3. Crear botón de descarga HTML (si existe)
            if (group.html) {
                const blob = new Blob([group.html.content], { type: 'text/html' });
                html_url = URL.createObjectURL(blob); // Guardar URL para el iframe
                
                const link = document.createElement('a');
                link.href = html_url;
                link.download = group.html.filename;
                link.textContent = `Descargar ${group.html.filename}`;
                link.className = 'btn success';
                link.style.margin = '5px';
                buttonContainer.appendChild(link);
            }

            // 4. Crear botón de descarga CSV (si existe)
            if (group.csv) {
                const blob = new Blob([group.csv.content], { type: 'text/csv;charset=utf-8-sig;' });
                const url = URL.createObjectURL(blob);
                
                const link = document.createElement('a');
                link.href = url;
                link.download = group.csv.filename;
                link.textContent = `Descargar ${group.csv.filename}`;
                link.className = 'btn primary';
                link.style.margin = '5px';
                buttonContainer.appendChild(link);
            }
            

            // 4.5. Crear botón de descarga EXCEL (si existe)
            if (group.xlsx) {
                // El contenido es base64, hay que decodificarlo
                const byteCharacters = atob(group.xlsx.content);
                const byteNumbers = new Array(byteCharacters.length);
                for (let i = 0; i < byteCharacters.length; i++) {
                    byteNumbers[i] = byteCharacters.charCodeAt(i);
                }
                const byteArray = new Uint8Array(byteNumbers);
                
                // Crear el Blob con el MIME type correcto para .xlsx
                const blob = new Blob([byteArray], { 
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
                });
                const url = URL.createObjectURL(blob);
                
                const link = document.createElement('a');
                link.href = url;
                link.download = group.xlsx.filename;
                link.textContent = `Descargar ${group.xlsx.filename}`;
                link.className = 'btn warning'; 
                link.style.margin = '5px';
                buttonContainer.appendChild(link);
            }
            // --- FIN: NUEVO BLOQUE PARA EXCEL ---

            // Añadir el contenedor de botones al contenedor del grupo
            groupContainer.appendChild(buttonContainer);

            // 5. Crear el iframe (solo si había un HTML)
            if (html_url) {
                const iframeWrapper = document.createElement('div');
                iframeWrapper.className = 'layout html-wrapper'; 
                const iframe = document.createElement('iframe');
                iframe.className = 'html-frame';
                iframe.src = html_url; // Reusar la URL del blob
                iframe.style.height = '600px'; 
                iframe.style.marginTop = '15px';
                iframe.style.marginBottom = '15px';
                iframeWrapper.appendChild(iframe);
                groupContainer.appendChild(iframeWrapper);
            }
            
            // Añadir este grupo completo al contenedor principal
            resultsContainer.appendChild(groupContainer);
        }
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

    /** MMuestra/Oculta el mensaje de carga */
    function showLoading(isLoading) {
        if (loadingMsg) {
            loadingMsg.style.display = isLoading ? 'block' : 'none';
        }
    }
});