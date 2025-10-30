# backend/logic.py

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
import numpy as np
import re
import unicodedata
import plotly.graph_objects as go
from typing import Optional, Tuple, List

# --- RUTAS A ARCHIVOS FIJOS ---
# Esta es la parte clave. Leemos los archivos desde la misma
# carpeta donde vive este script (logic.py).
BASE_DIR = os.path.dirname(__file__)
RUTA_DICC = os.path.join(BASE_DIR, "data\Diccionario NFV_2.xlsx")
RUTA_CATALOGO = os.path.join(BASE_DIR, "data\CID_Diferencias_catalogo_Vs_Antiago_250825.xlsx")

# ==========================================================
# SECCIÓN 1: LÓGICA DE LA MACRO (Leer TXT -> Excel)
# ==========================================================

def importar_txt_a_excel(path_dir, nombre_sitio="default"):
    """
    QUÉ HACE: Lee todos los .txt de una carpeta (path_dir),
    los procesa con pandas y los guarda en un nuevo archivo Excel.
    CÓMO FUNCIONA:
    1. Lista los .txt en path_dir.
    2. Crea un Workbook (un archivo Excel en memoria).
    3. Itera cada .txt, lo lee con pandas (saltando filas, usando '|'
       como delimitador) y lo limpia.
    4. Añade los datos limpios a una nueva hoja en el Workbook.
    5. Guarda el Workbook en la misma path_dir.
    DEVUELVE: La ruta completa al archivo .xlsx que se creó.
    """
    archivos = [f for f in os.listdir(path_dir) if f.endswith('.txt')]
    wb = Workbook()
    wb.remove(wb.active)
    for archivo in archivos:
        ruta = os.path.join(path_dir, archivo)
        nombre_hoja = os.path.splitext(archivo)[0][:31]

        df = pd.read_csv(ruta, delimiter='|', encoding='utf-8',header=1)
        df = df.iloc[1:].reset_index(drop=True)
        df = df.drop(df.columns[0], axis=1)
        df = df.drop(df.columns[-1], axis=1)
        df = df.dropna(how='all')
        ws = wb.create_sheet(title=nombre_hoja)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        df2 = pd.read_csv(ruta, delimiter='|', usecols=[0], header=None, encoding='utf-8')
        df2.columns = [""]
        df2 = df2.dropna(how='all')
        df2 = df2.iloc[3:].reset_index(drop=True)
        for r in dataframe_to_rows(df2, index=False, header=True):
            ws.append(r)

    wb.create_sheet(title="Consolidado")
    ruta_salida = os.path.join(path_dir, f"Macro_Impresos_Generada.xlsx")
    wb.save(ruta_salida)
    return ruta_salida

def consolidar_datos(path_excel):
    """
    QUÉ HACE: Abre el Excel generado y combina todas las hojas
    en la hoja "Consolidado".
    """
    wb = load_workbook(path_excel)
    ws_consolidado = wb["Consolidado"]
    ws_consolidado.append(["Region", "Sitio", "ID", "Name", "OS-EXT-SRV-ATTR: Hypervisor Hostname",
                           "OS-EXT-SRV-ATTR: Instance Name", "OS-EXT-AZ: Availability Zone",
                           "flavor: Ram", "flavor: Vcpus", "flavor: Disk", "Status"])
    
    for sheet in wb.sheetnames:
        if sheet == "Consolidado": continue
        ws = wb[sheet]
        region = sheet[:2]
        sitio = sheet.split('_', 1)[1] if '_' in sheet else sheet
        sitio = sitio.split('-', 1)[0] if '-' in sitio else sitio
        data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row): break
            fila = [region, sitio] + list(row[:10])
            data.append(fila)
        for row in data:
            ws_consolidado.append(row)
    wb.save(path_excel)

def crear_segunda_parte(path_excel):
    """
    QUÉ HACE: Lee la segunda parte de cada hoja (después de la
    fila vacía) y la consolida en una nueva hoja "2daParte".
    """
    wb = load_workbook(path_excel)
    data = []
    for sheet in wb.sheetnames:
        if sheet in ["Consolidado", "2daParte"]: continue
        ws = wb[sheet]
        region = sheet[:2]
        sitio = sheet.split('_', 1)[1] if '_' in sheet else sheet
        sitio = sitio.split('-', 1)[0] if '-' in sitio else sitio
        comenzar = False
        for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            if not any(row):
                comenzar = True
                continue
            if not comenzar: continue
            col_c = row[0]
            col_c = col_c.strip() if isinstance(col_c, str) else ''
            data.append([region, sitio, col_c])

    df = pd.DataFrame(data, columns=["Region", "Sitio", "Azone Hostname"])
    df["Azone"] = df["Azone Hostname"].apply(lambda x: x.split(" ")[0] if isinstance(x, str) else "")
    df["Hostname"] = df["Azone Hostname"].apply(lambda x: x.split(" ", 1)[1] if isinstance(x, str) and " " in x else x)
    df["AZ"] = df["Azone"].apply(lambda x: x if x in ["internal", "nova"] else x[-4:])
    df["Host"] = df.apply(lambda row: row["Hostname"].split(".")[0] if isinstance(row["Hostname"], str) and "." in row["Hostname"] else row["Azone Hostname"], axis=1)
    df = df.drop(columns=["Azone Hostname"])
    
    ws_2da = wb.create_sheet(title="2daParte")
    for r in dataframe_to_rows(df, index=False, header=True):
        ws_2da.append(r)
    wb.save(path_excel)

def ejecutar_macro_completa(path_dir):
    """
    QUÉ HACE: Orquesta los 3 pasos anteriores.
    DEVUELVE: La ruta al excel final.
    """
    archivo_excel = importar_txt_a_excel(path_dir)
    consolidar_datos(archivo_excel)
    crear_segunda_parte(archivo_excel)
    print("✅ Consolidado completado. Archivo guardado en:", archivo_excel)
    return archivo_excel

# ==========================================================
# SECCIÓN 2: LÓGICA DE MAPEADO (VNF/VNFC)
# ==========================================================

def _norm(s: str) -> str:
    """Utilidad: Normaliza texto (Mayúsculas, sin acentos, sin espacios)."""
    if pd.isna(s): return ""
    s = str(s).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return " ".join(s.split())

def _best_match(texto: str, candidatos, whole_word: bool = False):
    """Utilidad: Encuentra la mejor coincidencia de un texto en una lista."""
    textoN = _norm(texto)
    mejor, mejor_len = None, -1
    for cand in candidatos:
        cN = _norm(cand)
        if not cN: continue
        if whole_word:
            pat = rf'(?<![A-Z0-9]){re.escape(cN)}(?![A-Z0-9])'
            if re.search(pat, textoN):
                L = len(cN)
                if L > mejor_len:
                    mejor, mejor_len = cand, L
        else:
            if cN in textoN:
                L = len(cN)
                if L > mejor_len:
                    mejor, mejor_len = cand, L
    return mejor

def mapear_vnf_vnfc(df_names: pd.DataFrame, df_dicc: pd.DataFrame,
                    col_name="Name") -> pd.DataFrame:
    """
    QUÉ HACE: Lógica principal de mapeo VNF/VNFC.
    CÓMO FUNCIONA: Usa el diccionario para construir un mapa de
    "firmas" y "tipos de VM" por VNF, y luego lo aplica a cada
    fila en el dataframe de datos.
    """
    dicc = df_dicc.copy()
    necesarios = {"VNF", "VNFC", "VNF2", "VM Type", "Name contiene"}
    if not necesarios <= set(dicc.columns):
        raise ValueError(f"Faltan columnas en el diccionario: {necesarios - set(dicc.columns)}")

    firmas_por_vnf2 = {}
    for _, row in dicc.iterrows():
        vnf2 = row["VNF2"]
        if pd.isna(vnf2) or not str(vnf2).strip(): continue
        vnf2 = str(vnf2)
        firma = row.get("Name contiene", None)
        firmas_por_vnf2.setdefault(vnf2, set())
        if pd.notna(firma) and str(firma).strip():
            firmas_por_vnf2[vnf2].add(str(firma))
        firmas_por_vnf2[vnf2].add(vnf2)

    vmtypes_por_vnf2 = {}
    vnf_por_vnf2 = {}
    for _, row in dicc.iterrows():
        vnf2 = row["VNF2"]
        if pd.isna(vnf2) or not str(vnf2).strip(): continue
        vnf2 = str(vnf2)
        vnf = row["VNF"]
        if vnf_por_vnf2.get(vnf2) in (None, "") and pd.notna(vnf) and str(vnf).strip():
            vnf_por_vnf2[vnf2] = str(vnf)
        vm_t = row["VM Type"]
        vnfc = row["VNFC"]
        if pd.notna(vm_t) and str(vm_t).strip() and pd.notna(vnfc) and str(vnfc).strip():
            vmtypes_por_vnf2.setdefault(vnf2, [])
            vmtypes_por_vnf2[vnf2].append((str(vm_t), str(vnfc)))

    def resolver_name(name: str) -> Tuple[Optional[str], Optional[str]]:
        if pd.isna(name) or not str(name).strip():
            return (None, None)
        mejor_vnf2, mejor_len = None, -1
        for vnf2, firmas in firmas_por_vnf2.items():
            firma_match = _best_match(name, firmas, whole_word=False)
            if firma_match is not None:
                L = len(_norm(firma_match))
                if L > mejor_len:
                    mejor_vnf2, mejor_len = vnf2, L
        if not mejor_vnf2:
            return (None, None)
        vnf = vnf_por_vnf2.get(mejor_vnf2, None)
        pares_vmtype_vnfc = vmtypes_por_vnf2.get(mejor_vnf2, [])
        if not pares_vmtype_vnfc:
            return (vnf, None)
        mejor_vmtype, mejor_vnfc, mejor_len2 = None, None, -1
        for vm_t, vnfc_cand in pares_vmtype_vnfc:
            vm_match = _best_match(name, [vm_t], whole_word=True)
            if vm_match is not None:
                L2 = len(_norm(vm_match))
                if L2 > mejor_len2:
                    mejor_len2, mejor_vmtype, mejor_vnfc = L2, vm_t, vnfc_cand
        if mejor_vnfc is None:
            for vm_t, vnfc_cand in pares_vmtype_vnfc:
                vm_match = _best_match(name, [vm_t], whole_word=False)
                if vm_match is not None:
                    L2 = len(_norm(vm_match))
                    if L2 > mejor_len2:
                        mejor_len2, mejor_vmtype, mejor_vnfc = L2, vm_t, vnfc_cand
        return (vnf, mejor_vnfc)

    vnf_list, vnfc_list = [], []
    for name in df_names[col_name]:
        v, c = resolver_name(name)
        vnf_list.append(v)
        vnfc_list.append(c)
    return pd.DataFrame({"VNF": vnf_list, "VNFC": vnfc_list})

def completar_excepciones(df_names: pd.DataFrame,
                          mapeo: pd.DataFrame,
                          df_exc: pd.DataFrame,
                          col_name="Name") -> pd.DataFrame:
    """
    QUÉ HACE: Usa la hoja "Excepciones" del diccionario para
    rellenar los mapeos que fallaron.
    """
    out = mapeo.copy()
    if not {"VNF", "TIPO VM"} <= set(df_exc.columns):
        raise ValueError("La hoja 'Excepciones' debe tener columnas: 'VNF' y 'TIPO VM'.")
    lista_vnf_exc = [str(v) for v in df_exc["VNF"].dropna().astype(str) if str(v).strip()]

    for i, (vnf_val, vnfc_val) in enumerate(zip(out["VNF"], out["VNFC"])):
        if pd.isna(vnf_val) and pd.isna(vnfc_val):
            name = df_names.iloc[i][col_name]
            if pd.isna(name) or not str(name).strip(): continue
            vnf_match = _best_match(name, lista_vnf_exc, whole_word=True)
            if vnf_match is None:
                vnf_match = _best_match(name, lista_vnf_exc, whole_word=False)
            if vnf_match is not None:
                fila = df_exc[df_exc["VNF"].astype(str) == str(vnf_match)]
                tipo_vm = None
                if not fila.empty:
                    tipo_vm_series = fila["TIPO VM"].dropna().astype(str)
                    if not tipo_vm_series.empty:
                        tipo_vm = tipo_vm_series.iloc[0]
                out.at[i, "VNF"] = str(vnf_match)
                out.at[i, "VNFC"] = tipo_vm if tipo_vm is not None else None
    return out

# ==========================================================
# SECCIÓN 3: LÓGICA DE DIBUJADO (Plotly)
# ==========================================================

def f_tetris_plot(tetris_piezas: pd.DataFrame, tamaño_chip: int, sin_uso: int):
    """
    QUÉ HACE: Toma un DataFrame de "piezas de tetris" (calculado
    por build_tetris_for_cee) y lo dibuja con Plotly.
    CÓMO FUNCIONA:
    1. Crea una figura de Plotly (go.Figure).
    2. Itera cada fila del DataFrame y la añade como una barra
       horizontal (go.Bar) en la figura.
    3. Configura los ejes, el título y los estilos (hover, etc.).
    4. NO USA fig.show().
    DEVUELVE: El objeto 'fig' de Plotly.
    """
    df = tetris_piezas.copy()
    df['host_chip'] = df['HOST'].astype(str) + ' - ' + df['AZ'] + ' - ' + df['Host'].astype(str)
    ordered_hosts = (
        df[['AZ', 'Host', 'host_chip']]
        .drop_duplicates()
        .sort_values(by=['AZ', 'Host'])['host_chip']
        .tolist()[::-1]
    )
    fig = go.Figure()
    for _, row in df.iterrows():
        fig.add_trace(go.Bar(
            x=[row['Length']],
            y=[row['host_chip']],
            text=row['VM'],
            base=row['Start'],
            name=row['VM'],
            orientation='h',
            marker=dict(color=row['Color']),
            hovertemplate=(
                f'<b>{row["VM"]}</b><br>'
                f'Longitud: {row["Length"]}<br>'
                f'Anti-Afinidad: {int(row["Anti-Afinidad"])}<br>'
                f'Pieza: {row["Numero"]}<extra></extra>'
            ) if pd.notnull(row['VM']) else '<extra></extra>',
            showlegend=False
        ))
    fig.update_layout(
        barmode='stack',
        title='TETRIS actual',
        xaxis=dict(
            title='22 sept 2025',
            tickmode='linear',
            dtick=1,
            range=[0, 2 * tamaño_chip + 2 * sin_uso],
            showgrid=True
        ),
        yaxis=dict(
            title='Host',
            categoryorder='array',
            categoryarray=ordered_hosts
        ),
        height = 2000,
        width=1500,
        showlegend=False
    )
    fig.update_traces(
        textposition='inside',
        insidetextanchor='middle',
        textfont=dict(color='black', size=10),
        marker=dict(line=dict(width=1.5, color='black'))
    )
    # ¡IMPORTANTE! No hay fig.show()
    return fig

def build_tetris_for_cee(data, zonas, cee_value, ancho_cee=42):
    """
    QUÉ HACE: Filtra el DataFrame 'data' principal para un
    CEE específico y lo transforma en un formato de "piezas
    de tetris" listo para ser dibujado.
    DEVUELVE: Un DataFrame listo para f_tetris_plot.
    """
    tetris = []
    for zona in zonas:
        data_zona = data[data['OS-EXT-AZ: Availability Zone'] == zona]
        if cee_value is None:
            data_cee = data_zona[data_zona['CEE'].isna()]
        else:
            data_cee = data_zona[data_zona['CEE'].astype(str) == str(cee_value)]
        if data_cee.empty: continue
        
        tetris.append({'AZ': zona, 'Host': ' ', 'Chip': ' ', 'VM': zona,
                       'Start': 0, 'Length': ancho_cee, 'Color': '#696969',
                       'Inst.': ' ', 'Anti-Afinidad': 0, 'Requerimiento': ' ', 'HOST': ' '})
        hosts = data_cee['OS-EXT-SRV-ATTR: Hypervisor Hostname'].unique()
        for host_name in hosts:
            data_host = data_cee[data_cee['OS-EXT-SRV-ATTR: Hypervisor Hostname'] == host_name] \
                .sort_values(by='flavor: Vcpus', ascending=False)
            chip1 = 0
            chip2 = 20
            cee_tag = f'CEE{cee_value}' if cee_value is not None else 'CEE_NA'
            tetris.append({'AZ': zona, 'Host': host_name, 'Chip': ' ', 'VM': ' ',
                           'Start': 17, 'Length': 3, 'Color': '#000000',
                           'Inst.': ' ', 'Anti-Afinidad': 0, 'Requerimiento': ' ', 'HOST': cee_tag})
            tetris.append({'AZ': zona, 'Host': host_name, 'Chip': ' ', 'VM': ' ',
                           'Start': 37, 'Length': 3, 'Color': '#000000',
                           'Inst.': ' ', 'Anti-Afinidad': 0, 'Requerimiento': ' ', 'HOST': cee_tag})

            for _, row in data_host.iterrows():
                tamaño = row['flavor: Vcpus'] / 2
                if chip1 + tamaño <= 17:
                    chip = 'Chip 1'
                    start_rel = chip1
                    chip1 += tamaño
                else:
                    chip = 'Chip 2'
                    start_rel = chip2
                    chip2 += tamaño
                tetris.append({'AZ': zona, 'Host': host_name, 'Chip': chip,
                               'VM': row['Name'], 'Start': start_rel, 'Length': tamaño,
                               'Color': row['Color'], 'Inst.': ' ', 'Anti-Afinidad': 0,
                               'Requerimiento': ' ', 'HOST': cee_tag})
    tetris_df = pd.DataFrame(tetris)
    if tetris_df.empty: return tetris_df
    if 'HOST' not in tetris_df.columns: tetris_df['HOST'] = ' '
    tetris_df['Numero'] = ' '
    tetris_df = tetris_df.sort_values(by=['AZ', 'HOST', 'Host', 'Chip', 'Start']).reset_index(drop=True)
    return tetris_df

# ==========================================================
# SECCIÓN 4: FUNCIÓN "CEREBRO" PRINCIPAL
# ==========================================================

def generar_reportes_tetris(temp_dir_path: str) -> List[dict]:
    """
    QUÉ HACE: Es la función "cerebro" que orquesta todo.
    CÓMO FUNCIONA:
    1. Llama a ejecutar_macro_completa() para procesar los .txt.
    2. Carga los archivos FIJOS (diccionario y catálogo) desde las
       rutas estáticas (RUTA_DICC, RUTA_CATALOGO).
    3. Carga los datos consolidados del Excel generado.
    4. Ejecuta toda la lógica de mapeo (mapear_vnf_vnfc, etc.).
    5. Ejecuta la lógica de asignación de colores.
    6. Ejecuta la lógica de extracción de CEEs.
    7. Itera por cada CEE, llama a build_tetris_for_cee y f_tetris_plot.
    8. Convierte cada gráfico de Plotly a un string HTML.
    DEVUELVE: Una lista de diccionarios, donde cada dict tiene
              el 'filename' y el 'content' (HTML) del gráfico.
    """
    
    # 1. Ejecutar macro (sobre los .txt subidos en temp_dir_path)
    ruta_macro = ejecutar_macro_completa(temp_dir_path)
    
    # 2. Cargar diccionarios (desde rutas fijas)
    print("Cargando diccionarios...")
    df_dicc = pd.read_excel(RUTA_DICC, sheet_name="Diccionario")
    df_exc  = pd.read_excel(RUTA_DICC, sheet_name="Excepciones")

    # 3. Cargar Excel generado
    print("Cargando datos consolidados...")
    data = pd.read_excel(ruta_macro, sheet_name="Consolidado")
    zonas = sorted(data["OS-EXT-AZ: Availability Zone"].unique())

    # 4. Mapeo de VNF/VNFC
    print("Mapeando VNF/VNFC...")
    mapeo_inicial = mapear_vnf_vnfc(data, df_dicc, col_name="Name")
    mapeo_final = completar_excepciones(data, mapeo_inicial, df_exc, col_name="Name")
    v = mapeo_final["VNF"].astype(str).str.strip().replace({"<NA>": np.nan, "nan": np.nan, "": np.nan})
    c = mapeo_final["VNFC"].astype(str).str.strip().replace({"<NA>": np.nan, "nan": np.nan, "": np.nan})
    data["Name"] = np.where(~v.isna() & ~c.isna(), v + " " + c,
                        np.where(~v.isna(), v, data["Name"]))

    # 5. Asignar color (desde ruta fija)
    print("Asignando colores...")
    base = pd.read_excel(RUTA_CATALOGO, header=1)
    base["VNF"] = base["VNF"].astype(str)
    data["Name"] = data["Name"].astype(str)
    def obtener_color(name):
        for _, fila in base.iterrows():
            if fila["VNF"] in name:
                return fila["Color"]
        return None
    data["Color"] = data["Name"].apply(obtener_color)
    data["Color"].fillna("#808080", inplace=True) # Color gris por defecto

    # 6. Extraer CEE
    print("Extrayendo CEEs...")
    data["CEE"] = data["OS-EXT-AZ: Availability Zone"].astype(str).apply(
        lambda x: re.search(r"vapp(\d+)", x).group(1) if re.search(r"vapp(\d+)", x) else None
    )

    # 7. Generar dibujos
    print("Generando gráficos...")
    cee_vals = pd.to_numeric(data["CEE"], errors="coerce")
    unique_cees = sorted([int(x) for x in cee_vals.dropna().unique()])
    if data["CEE"].isna().any():
        unique_cees.append(None)
    
    tamaño_chip = 17
    sin_uso = 3
    
    html_outputs = [] # Lista para guardar nuestros resultados

    for cee in unique_cees:
        cee_tag = f"CEE{cee}" if cee is not None else "CEE_NA"
        print(f"Procesando {cee_tag}...")
        
        tetris_cee = build_tetris_for_cee(data, zonas, cee, ancho_cee=42)
        if tetris_cee.empty:
            print(f"No hay piezas para {cee_tag}, se omite.")
            continue

        fig = f_tetris_plot(tetris_cee, tamaño_chip, sin_uso)
        
        # Convertir figura a string HTML
        html_content = fig.to_html(full_html=True, include_plotlyjs='cdn')
        
        html_outputs.append({
            "filename": f"tetris_{cee_tag}.html",
            "content": html_content
        })

    print("✅ Proceso completado.")
    return html_outputs