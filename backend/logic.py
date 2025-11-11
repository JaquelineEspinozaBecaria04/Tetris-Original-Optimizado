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
import io
import base64
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- RUTAS A ARCHIVOS FIJOS ---
BASE_DIR = os.path.dirname(__file__)
RUTA_DICC = os.path.join(BASE_DIR, "data", "Diccionario NFV_2.xlsx")
RUTA_CATALOGO = os.path.join(BASE_DIR, "data", "CID_Diferencias_catalogo_Vs_Antiago_250825.xlsx")

# ==========================================================
# SECCIÓN 1: LÓGICA DE LA MACRO (Leer TXT -> Excel)
# ==========================================================
def importar_txt_a_excel(path_dir, nombre_sitio="default"):
    archivos = []
    for root, dirs, files in os.walk(path_dir):
        for file in files:
            if file.endswith('.txt'):
                full_path = os.path.join(root, file)
                sheet_name = os.path.splitext(file)[0][:31]
                archivos.append((full_path, sheet_name))
    wb = Workbook()
    wb.remove(wb.active)
    print(f"Encontrados {len(archivos)} archivos .txt para procesar.")
    for ruta, nombre_hoja in archivos:
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
        sitio = sitio.split('-', 1)[0] if '-' in sitio else sheet
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
    Formatea la AZ como 'AZ##'.
    """
    wb = load_workbook(path_excel)
    data = []
    for sheet in wb.sheetnames:
        if sheet in ["Consolidado", "2daParte"]: continue
        ws = wb[sheet]
        region = sheet[:2]
        sitio = sheet.split('_', 1)[1] if '_' in sheet else sheet
        sitio = sitio.split('-', 1)[0] if '-' in sitio else sheet
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
    
    # --- Función de ayuda para extraer AZ ---
    def format_az_from_azone(x):
        if x in ["internal", "nova"]:
            return x
        az_str = str(x)
        
        # Intenta encontrar 'az##' (ej. 'ceeaz01' -> 'az01')
        match = re.search(r"(az\d{2})", az_str, re.IGNORECASE)
        if match:
            return match.group(1).upper() # -> "AZ01"
        
        # Si no, intenta encontrar 'cee##' (ej. 'cee01' -> '01')
        match_cee = re.search(r"cee(\d{2})", az_str, re.IGNORECASE)
        if match_cee:
            return "AZ" + match_cee.group(1) # -> "AZ01"
        
        # Fallback si no encuentra patrones conocidos (toma los últimos 2 dígitos y espera que sean el número)
        match_num = re.search(r"(\d{2})$", az_str)
        if match_num:
             return "AZ" + match_num.group(1)
            
        return "AZ_ND" # Fallback final

    df["AZ"] = df["Azone"].apply(format_az_from_azone)
    
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
def f_tetris_plot(tetris_piezas: pd.DataFrame, tamaño_chip: int, sin_uso: int, site_name: str, cee_tag: str):
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
    # Rellenar valores nulos para evitar errores de concatenación
    df['AZ'] = df['AZ'].fillna('AZ_ND') # Asegura que AZ nunca sea nulo
    df['Host'] = df['Host'].fillna('')
    df['CEE'] = df['CEE'].fillna('') 
    
    # Cambiado de 'HOST' (que es el número 1,2,3) a 'CEE' (que es "CEE3")
    df['host_chip'] = df['CEE'].astype(str) + ' - ' + df['AZ'] + ' - ' + df['Host'].astype(str)
    
    ordered_hosts = (
        df[['AZ', 'Host', 'host_chip']]
        .drop_duplicates()
        .sort_values(by=['AZ', 'Host'])['host_chip']
        .tolist()[::-1]
    )
    fig = go.Figure()
    
    # Construir el título dinámico
    title_text = f'TETRIS actual - {site_name} - {cee_tag}'
    
    for _, row in df.iterrows():
        # 1. Definir el texto de la pieza (solo si es una pieza real)
        if row["VM_Individual_Count"] > 0:
            pieza_str = f'Pieza: {row["VM_Individual_Count"]} de {row["VM_Total_Count"]}'
        else:
            pieza_str = ''

        # 2. Construir el 'hovertemplate'
        hover_template = (
            f'<b>{row["VM"]}</b><br>'
            f'Longitud: {row["Length"]}<br>'
            f'Anti-Afinidad: {int(row["Anti-Afinidad"])}<br>'
            f'{pieza_str}<extra></extra>'
        ) if pd.notnull(row['VM']) else '<extra></extra>'

        fig.add_trace(go.Bar(
            x=[row['Length']],
            y=[row['host_chip']],
            text=row['VM'],
            base=row['Start'],
            name=row['VM'],
            orientation='h',
            marker=dict(color=row['Color']),
            hovertemplate=hover_template,
            showlegend=False
        ))
    fig.update_layout(
        barmode='stack',
        title=title_text, # Asignar el título dinámico
        xaxis=dict(
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
    return fig

def build_tetris_for_cee(data, zonas, cee_value, ancho_cee=42):
    """
    QUÉ HACE: Filtra el DataFrame 'data' principal para un
    CEE específico y lo transforma en un formato de "piezas
    de tetris" listo para ser dibujado.
    DEVUELVE: Un DataFrame listo para f_tetris_plot.
    """
    tetris = []
    
    # Obtiene los datos para el CEE actual
    if cee_value is None:
        data_cee_full = data[data['CEE'].isna()]
    else:
        data_cee_full = data[data['CEE'].astype(str) == str(cee_value)]

    # Obtiene la lista de zonas únicas *dentro de este CEE*
    # Usa la columna 'OS-EXT-AZ_Formatted' que ya tiene "AZ01", etc.
    zonas_en_cee = sorted(data_cee_full["OS-EXT-AZ_Formatted"].dropna().unique())

    for zona in zonas_en_cee: # Itera sobre las zonas formateadas (ej. 'AZ01')
        # Filtra el DF del CEE por la zona actual
        data_cee = data_cee_full[data_cee_full["OS-EXT-AZ_Formatted"] == zona]
        if data_cee.empty: continue
        
        # 'zona' ya es el string 'AZ##' formateado
        short_zona = zona 
       
        zona_larga_series = data_cee['OS-EXT-AZ: Availability Zone'].dropna()
        zona_larga = zona_larga_series.iloc[0] if not zona_larga_series.empty else short_zona
        
        tetris.append({'AZ': short_zona, 'Host': ' ', 'Chip': ' ', 'VM': zona_larga,
                       'Start': 0, 'Length': ancho_cee, 'Color': '#696969',
                       'Inst.': ' ', 'Anti-Afinidad': 0, 'Requerimiento': ' ', 'HOST': ' '})
        hosts = data_cee['OS-EXT-SRV-ATTR: Hypervisor Hostname'].unique()
        for host_name in hosts:
            data_host = data_cee[data_cee['OS-EXT-SRV-ATTR: Hypervisor Hostname'] == host_name] \
                .sort_values(by='flavor: Vcpus', ascending=False)
            chip1 = 0
            chip2 = 20
            cee_tag = f'CEE{cee_value}' if cee_value is not None else 'CEE_NA'
            
            tetris.append({'AZ': short_zona, 'Host': host_name, 'Chip': ' ', 'VM': ' ',
                           'Start': 17, 'Length': 3, 'Color': '#000000',
                           'Inst.': ' ', 'Anti-Afinidad': 0, 'Requerimiento': ' ', 'HOST': cee_tag})

            tetris.append({'AZ': short_zona, 'Host': host_name, 'Chip': ' ', 'VM': ' ',
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

                tetris.append({'AZ': short_zona, 'Host': host_name, 'Chip': chip,
                               'VM': row['Name'], 'Start': start_rel, 'Length': tamaño,
                               'Color': row['Color'], 'Inst.': ' ', 'Anti-Afinidad': 0,
                               'Requerimiento': ' ', 'HOST': cee_tag})
    
    tetris_df = pd.DataFrame(tetris)
    if tetris_df.empty: return tetris_df
    if 'HOST' not in tetris_df.columns: tetris_df['HOST'] = ' '
    
   
    # 1. Ordenar ANTES del cumcount para numeración correcta
    tetris_df = tetris_df.sort_values(by=['AZ', 'HOST', 'Host', 'Chip', 'Start']).reset_index(drop=True)

    # 2. Calcular 'Numero' (X de Y)
    vm_rows = tetris_df['VM'].str.strip().fillna('').ne('') & tetris_df['Length'].gt(0)
    vm_groups = tetris_df[vm_rows].groupby('VM')
    tetris_df.loc[vm_rows, 'VM_Total_Count'] = vm_groups['VM'].transform('count')
    tetris_df.loc[vm_rows, 'VM_Individual_Count'] = vm_groups.cumcount() + 1
    tetris_df['VM_Total_Count'] = tetris_df['VM_Total_Count'].fillna(0).astype(int)
    tetris_df['VM_Individual_Count'] = tetris_df['VM_Individual_Count'].fillna(0).astype(int)
    
    tetris_df['Numero'] = tetris_df.apply(
        lambda row: f"{row['VM_Individual_Count']} de {row['VM_Total_Count']}" if row['VM_Individual_Count'] > 0 else '',
        axis=1
    )
    
    # 3. CREAR COLUMNA CEE Y RE-NUMERAR HOST (para 'host_chip' y CSV)
    tetris_df['CEE'] = tetris_df['HOST']
    unique_hosts = tetris_df['Host'].unique()
    real_hosts = [h for h in unique_hosts if h.strip() != '']
    host_map = {host_name: i + 1 for i, host_name in enumerate(real_hosts)}
    tetris_df['HOST'] = tetris_df['Host'].map(host_map)
    tetris_df['HOST'] = tetris_df['HOST'].astype('object').fillna(' ') 

    # 4. CALCULAR HOST_RELATIVO (para el Excel)
    host_relativos = []
    for zona, grupo in tetris_df.groupby("AZ", sort=False):
        unique_hosts_zona = grupo['Host'].unique()
        real_hosts_zona = [h for h in unique_hosts_zona if h.strip() != '']
        host_map_zona = {host_name: i + 1 for i, host_name in enumerate(real_hosts_zona)}
        host_relativos.extend(grupo['Host'].map(host_map_zona))
        
    tetris_df['HOST_RELATIVO'] = host_relativos
  
    return tetris_df
# ==========================================================
# SECCIÓN 4: FUNCIÓN "CEREBRO" PRINCIPAL
# ==========================================================

def generar_reportes_tetris(temp_dir_path: str, site_name: str) -> List[dict]:
    """
    QUÉ HACE: Es la función "cerebro" que orquesta todo.
    ...
    11. (NUEVO) Genera un Excel estilizado para cada gráfico.
    DEVUELVE: Una lista de diccionarios (HTMLs, CSVs y Excels).
    """
    
    # 1. Ejecutar macro (sobre los .txt subidos en temp_dir_path)
    ruta_macro = ejecutar_macro_completa(temp_dir_path)
    
    # 2. Cargar diccionarios (desde rutas fijas)
    print("Cargando diccionarios...")
    df_dicc = pd.read_excel(RUTA_DICC, sheet_name="Diccionario")
    df_exc  = pd.read_excel(RUTA_DICC, sheet_name="Excepciones")

    # 3. Cargar Excel generado y ENRIQUECER datos de AZ faltantes
    print("Cargando datos consolidados...")
    data = pd.read_excel(ruta_macro, sheet_name="Consolidado") 
    
    # --- Lógica para rellenar AZ faltantes ---
    print("Enriqueciendo datos de AZ...")
    df_2daparte = pd.read_excel(ruta_macro, sheet_name="2daParte")
    # Crear nombre corto de Host en el DF principal
    data['Host_short'] = data['OS-EXT-SRV-ATTR: Hypervisor Hostname'].astype(str).str.split('.').str[0]
    # Crear mapa de Host -> AZ desde la 2daParte (ya formateada como 'AZ##')
    df_2daparte_clean = df_2daparte[['Host', 'AZ']].drop_duplicates().dropna(subset=['Host'])
    host_az_map = df_2daparte_clean.set_index('Host')['AZ'].to_dict()
    
    # Función para aplicar el mapa
    def fill_missing_az(row):
        az = row['OS-EXT-AZ: Availability Zone']
        if pd.isna(az) or str(az).strip() == '':
            # Si está vacío, buscar en el mapa.
            return host_az_map.get(row['Host_short'], az) # Devuelve 'AZ##' o 'None'
        return az # Si no está vacío, devolver el original

    data['OS-EXT-AZ: Availability Zone'] = data.apply(fill_missing_az, axis=1)
    
    # --- Función de formateo de AZ (Corregida BUGS 1 y 3) ---
    def format_az_column(az):
        if pd.isna(az) or str(az).strip() == '':
            return "AZ_ND" # Zona No Definida
        az_str = str(az)
        if az_str in ["internal", "nova"]:
            return az_str
        
        # Intenta encontrar 'az##' (ej. 'r3...ceeaz01' -> 'az01')
        match = re.search(r"(az\d{2})", az_str, re.IGNORECASE)
        if match:
            return match.group(1).upper() # -> "AZ01"
        
        # Si no, intenta encontrar 'cee##' (ej. 'r3...cee01' -> '01')
        match_cee = re.search(r"cee(\d{2})", az_str, re.IGNORECASE)
        if match_cee:
            return "AZ" + match_cee.group(1) # -> "AZ01"
        
        # Fallback si no encuentra patrones conocidos (toma los últimos 2 dígitos y espera que sean el número)
        match_num = re.search(r"(\d{2})$", az_str)
        if match_num:
             return "AZ" + match_num.group(1)

        return "AZ_ND" 

    data["OS-EXT-AZ_Formatted"] = data['OS-EXT-AZ: Availability Zone'].apply(format_az_column)
    
    # Obtener lista de zonas únicas (ORIGINALES) para pasar a build_tetris
    zonas = sorted(data["OS-EXT-AZ: Availability Zone"].dropna().unique())

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
    
    html_outputs = [] 
    for cee in unique_cees:
        cee_tag = f"CEE{cee}" if cee is not None else "CEE_NA"
        print(f"Procesando {cee_tag}...")

        tetris_cee = build_tetris_for_cee(data, None, cee, ancho_cee=42)
        if tetris_cee.empty:
            print(f"No hay piezas para {cee_tag}, se omite.")
            continue
        
        # Reemplazar ' ' por 'INFRA'
        tetris_cee.loc[tetris_cee["VM"] == " ", "VM"] = "INFRA"
    

        df = tetris_cee.copy()
        df['AZ'] = df['AZ'].fillna('AZ_ND')
        df['Host'] = df['Host'].fillna('')
        df['CEE'] = df['CEE'].fillna('')
        df['host_chip'] = df['CEE'].astype(str) + ' - ' + df['AZ'] + ' - ' + df['Host'].astype(str)
        tetris_cee['host_chip'] = df['host_chip']
        
        # Generar nombre de archivo base con el nombre del sitio
        base_filename = f"{site_name}_tetris_{cee_tag}"
        
        # --- 1. Generar HTML ---
        fig = f_tetris_plot(tetris_cee, tamaño_chip, sin_uso, site_name, cee_tag)
        html_content = fig.to_html(full_html=True, include_plotlyjs=True)
        html_outputs.append({
            "filename": f"{base_filename}.html",
            "content": html_content
        })

        # --- 2. Generar CSV ---
        print(f"Generando CSV para {cee_tag}...")
        
        columnas_map = {
            'AZ': 'AZ',
            'Host': 'HOST_RELATIVO', # ID relativo de AZ
            'Chip': 'Chip',
            'VM': 'VM',
            'Start': 'Start',
            'Length': 'Length',
            'Color': 'Color',
            'Inst.': 'Inst.',
            'Anti-Afinidad': 'Anti-Afinidad',
            'HOST': 'HOST', # ID total del CRV
            'Numero': 'Numero',
            'host_chip': 'host_chip'
        }

        df_csv_final = pd.DataFrame()
        for col_deseada, col_original in columnas_map.items():
            if col_original in tetris_cee.columns:
                df_csv_final[col_deseada] = tetris_cee[col_original]
            else:
                df_csv_final[col_deseada] = np.nan 

        csv_content = df_csv_final.to_csv(index=False, encoding='utf-8-sig')
        
        html_outputs.append({
            "filename": f"{base_filename}.csv",
            "content": csv_content
        })

        # --- 3. Generar Excel ---
        print(f"Generando Excel para {cee_tag}...")
        try:
            # 1. Generar diccionario de piezas (para colores y tamaños)
            diccionario_excel = generar_diccionario(tetris_cee)
            # 2. Generar matriz de datos ancha
            matriz_excel = generar_matriz(tetris_cee)
            # 3. Definir columnas
            columnas_excel = definir_columnas(col_info)
            # 4. Obtener DataFrame final para Excel
            df_excel_final = obtener_df_final(matriz_excel, columnas_excel)
            
            # 5. Generar y Estilar el Excel en memoria
            excel_bytes = generar_excel_bytes(site_name, df_excel_final, diccionario_excel)
            
            # 6. Codificar en Base64 para enviarlo por JSON
            excel_base64 = base64.b64encode(excel_bytes).decode('utf-8')
            
            # 7. Añadir a los resultados
            html_outputs.append({
                "filename": f"{base_filename}.xlsx",
                "content": excel_base64
            })
            print(f"✅ Excel generado para {cee_tag}")

        except Exception as e:
            print(f"❌ Error generando Excel para {cee_tag}: {e}")
            # El proceso continúa aunque falle el Excel


    print("✅ Proceso completado.")
    return html_outputs

# ==========================================================
# SECCIÓN 5: LÓGICA DE EXCEL FINAL
# ==========================================================

# --- Constantes ---
NUMERO_CHIPS = 2
NUMERO_CORES = 20
col_info = ['# S. TOTAL', '# S. ZONA', 'SERVIDOR']


def generar_diccionario(df_solucion):
    """
    Crea un diccionario de VM -> [Length, Color] a partir
    del dataframe 'tetris_cee'.
    """
    df_diccionario = df_solucion[['VM', 'Length', 'Color']]
    df_unique = df_diccionario.drop_duplicates(subset="VM")
    diccionario = {fila.VM: [fila.Length, fila.Color] for fila in df_unique.itertuples(index=False)}

    diccionario['INFRA'] = [3, '#FBE2D5'] # Color crema/infra 
    # Remover la entrada ' ' que ya no se usa (si existiera)
    if ' ' in diccionario:
        del diccionario[' ']
    
    # Manejar VMs sin color (como 'DataMigrationPCRF None')
    for vm, (length, color) in diccionario.items():
        if pd.isna(color):
            diccionario[vm] = [length, "#808080"] # Gris por defecto si falta color
            
    return diccionario

def generar_matriz(df_solucion):
    """
    Convierte el dataframe 'tetris_cee' (formato largo) en
    una matriz (formato ancho) lista para el Excel.
    """
    # Rellenar NaNs en HOST_RELATIVO para evitar errores
    df_solucion['HOST_RELATIVO'] = df_solucion['HOST_RELATIVO'].fillna(0)

    servidores = df_solucion['HOST'].unique()
    matriz = [[np.nan for _ in range((NUMERO_CORES*NUMERO_CHIPS) + len(col_info) + 1)] for _ in range(len(servidores))]
    
    # Crear un mapa de Host (nombre largo) a su # S. TOTAL (HOST)
    host_to_num = df_solucion.drop_duplicates(subset='Host').set_index('Host')['HOST'].to_dict()

    for fila in df_solucion.itertuples(index=False):
        # Usar el # S. TOTAL (entero) como índice de fila
        host_num = host_to_num.get(fila.Host)
        if host_num is None or host_num == ' ': continue # Omitir filas sin host (ej. cabecera AZ)
        
        idx_fila = int(host_num) - 1 # Convertir a índice 0-based
        
        # Asegurarse de que el 'Start' es un número
        try:
            idx_col = int(fila.Start)
        except ValueError:
            continue

        # Escribir la VM ('INFRA' o el nombre)
        matriz[idx_fila][idx_col] = fila.VM
            
        # --- (MODIFICADO) Asegurarse de que AZ nunca sea NaN en la matriz ---
        matriz[idx_fila][-4] = fila.AZ if pd.notna(fila.AZ) else "AZ_ND"
        matriz[idx_fila][-3] = fila.HOST # # S. TOTAL
        matriz[idx_fila][-2] = int(fila.HOST_RELATIVO) # # S. ZONA
        matriz[idx_fila][-1] = fila.Host # SERVIDOR
    
    # Filtrar filas vacías que pudieron crearse por 'servidores'
    matriz = [fila for fila in matriz if not all(pd.isna(c) for c in fila)]
    return matriz

def definir_columnas(col_info):
    """Define las columnas para el DataFrame ancho. """
    return [f'chip_{i}_{j}' for i in range(1, NUMERO_CHIPS + 1) for j in range(1, NUMERO_CORES + 1)] + ['AZ'] + col_info

def obtener_df_final(matriz, columnas):
    """Crea el DataFrame final y añade las filas de separación por AZ."""
    df = pd.DataFrame(matriz, columns=columnas)
    
    if 'SERVIDOR' in df.columns:
        df['SERVIDOR'] = df['SERVIDOR'].replace(' ', 'INFRA')

    # Insertar filas vacías entre cambios de AZ
    if not df.empty:
        cambios = df["AZ"] != df["AZ"].shift()
        filas_cambio = df.index[cambios].tolist()[1:]
        for idx in reversed(filas_cambio):  # en reversa
            df = pd.concat([
                df.iloc[:idx],
                pd.DataFrame([{}], columns=df.columns), # Fila vacía
                df.iloc[idx:]
            ], ignore_index=True)
            
    # Reordenar columnas 
    cols = df.columns
    try:
        df = df[cols[-len(col_info):].tolist() + cols[:NUMERO_CORES].tolist() + [cols[-(len(col_info)+1)]] + cols[NUMERO_CORES:-(len(col_info)+1)].tolist()]
    except Exception as e:
        print(f"Error reordenando columnas, cayendo a método simple: {e}")
        # Fallback a lógica simple
        cols_info_list = ['AZ'] + col_info
        cols_chip = [c for c in df.columns if c.startswith("chip_")]
        df = df[cols_info_list + cols_chip]
    
    return df

# --- Función de Estilizado y Guardado  ---
def generar_excel_bytes(sitio: str, df: pd.DataFrame, diccionario: dict) -> bytes:
    """
    Toma el DataFrame final, lo escribe en un Workbook de openpyxl,
    aplica todo el estilo de la Celda 18, y devuelve los bytes
    del archivo .xlsx.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Tetris"

    # 1. Escribir el DataFrame en la hoja
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # 2. Insertar fila de título
    filas_recorridas = 2 # 1 (header) + 1 (excel 1-based)
    columnas_recorridas = 1 # Para el índice de 1 de openpyxl
    ws.insert_rows(1)
    filas_recorridas += 1 # Ahora filas_recorridas es 3 
    ws.cell(row=1, column=1, value=f"TETRIS ACTUAL {sitio}")

    # 3. Definir estilos
    fuente = Font(name="Aptos Narrow", size=11)
    alineacion = Alignment(horizontal="center", vertical="center")
    borde = Side(style="thin", color="000000")
    borde_completo = Border(left=borde, right=borde, top=borde, bottom=borde)

    # 4. Aplicar estilos generales (bordes, fuente, alineación)
    fila_inicio = 1
    fila_fin = df.shape[0] + 2 # +1 por header, +1 por título
    columna_inicio = 1
    columna_fin = df.shape[1]

    for fila in range(fila_inicio, fila_fin + 1):
        for col in range(columna_inicio, columna_fin + 1):
            celda = ws.cell(row=fila, column=col)
            celda.font = fuente
            celda.alignment = alineacion
            celda.border = borde_completo

    # 5. Combinar celdas y colorear piezas
    cols = [c for c in df.columns if c.startswith("chip_")] 
    df_stack = df[cols].stack()

    for (idx, col), val in df_stack.items():
        if pd.isna(val):
            continue
            
        col_num = df.columns.get_loc(col)
        
        # --- Usar filas_recorridas (3) y columnas_recorridas (1) ---
        fila_excel = idx + filas_recorridas 
        col_excel = col_num + columnas_recorridas 

        try:
            pieza_info = diccionario.get(val)
            if pieza_info is None:
                print(f"Advertencia: VM '{val}' no encontrada en diccionario de colores.")
                continue
                
            longitud = int(pieza_info[0])
            color_hex = str(pieza_info[1])[1:] # Quitar '#'
            
            if longitud <= 0:
                continue

            fila_inicio_merge = fila_excel
            fila_fin_merge = fila_excel
            col_inicio_merge = col_excel
            col_fin_merge = col_excel + longitud - 1

            if col_fin_merge > columna_fin:
                col_fin_merge = columna_fin
            
            ws.merge_cells(start_row=fila_inicio_merge, start_column=col_inicio_merge, end_row=fila_fin_merge, end_column=col_fin_merge)
            
            relleno = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
            ws.cell(row=fila_inicio_merge, column=col_inicio_merge).fill = relleno

        except Exception as e:
            print(f"Error al procesar pieza {val} en fila {idx}: {e}")

    # 6. Colorear Headers (Fila 2)
    fila_headers = 2
    color_header = 'FBE2D5' # Crema
    relleno_header = PatternFill(start_color=color_header, end_color=color_header, fill_type="solid")
    
    for col in range(columna_inicio, columna_fin + 1):
        celda = ws.cell(row=fila_headers, column=col)
        texto = celda.value
        if texto and "chip" in str(texto):
            try:
                celda.value = int(texto.split("_")[-1])
            except:
                pass 

    # 7. Colorear columnas de info
    fila_inicio_datos = 3 # Fila 3 del excel
    fila_fin_datos = fila_inicio_datos + df.shape[0] -1 # Corrección: df.shape[0] ya incluye separadores
    
    # Bucle 1: Colorear columnas de info (excepto SERVIDOR)
    for fila in range(fila_inicio_datos, fila_fin_datos + 1):
        for col in range(columna_inicio, len(col_info)): # col 1 y 2
            ws.cell(row=fila, column=col).fill = relleno_header

    # Bucle 2: Poner en negrita la columna SERVIDOR
    fuente_negrita = Font(name="Aptos Narrow", size=11, bold=True)
    col_servidor_idx = len(col_info) # Columna 3
    for fila in range(fila_inicio_datos, fila_fin_datos + 1):
         ws.cell(row=fila, column=col_servidor_idx).font = fuente_negrita

    # 9. Colorear filas vacías (separadores de AZ)
    indices_nulos = df.index[df.isnull().all(axis=1)].tolist()
    color_separador = "000000" # Negro
    relleno_separador = PatternFill(start_color=color_separador, end_color=color_separador, fill_type="solid")
    
    for idx_nulo in indices_nulos:
        fila_excel = idx_nulo + filas_recorridas
        ws.merge_cells(start_row=fila_excel, start_column=columna_inicio, end_row=fila_excel, end_column=columna_fin)
        celda_merge = ws.cell(row=fila_excel, column=columna_inicio)
        celda_merge.fill = relleno_separador
        celda_merge.value = "INFRA"
        
    # 10. Estilo del Título (Fila 1)
    fuente_titulo = Font(name="Aptos Narrow", size=16, bold=True, color="FF0000")
    alineacion_titulo = Alignment(horizontal="center", vertical="center")
    color_titulo = "CAEDFB" # Azul claro
    relleno_titulo = PatternFill(start_color=color_titulo, end_color=color_titulo, fill_type="solid")
    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=columna_fin)
    celda_titulo = ws.cell(row=1, column=1)
    celda_titulo.font = fuente_titulo
    celda_titulo.alignment = alineacion_titulo
    celda_titulo.fill = relleno_titulo
    
    # 11.  Estilizar y fusionar columna AZ (vertical)
    try:
        col_az_idx = df.columns.get_loc('AZ') + 1
        fuente_az = Font(name="Aptos Narrow", size=14, bold=True)
        alineacion_az = Alignment(horizontal="center", vertical="center", textRotation=90)
        color_az = "B5E6A2" # Verde 
        relleno_az = PatternFill(start_color=color_az, end_color=color_az, fill_type="solid")
        
        fila_inicio_az = 3 # Empezar desde la primera fila de datos (Fila 3 en Excel)
        
        # Recorrer las filas nulas + la última fila para definir los bloques a fusionar
        filas_de_corte = indices_nulos + [df.shape[0]] 
        
        for idx_corte in filas_de_corte:
            # +2 porque filas_recorridas es 3, y el índice de df es 0-based
            fila_fin_az = idx_corte + 2 
            
            if fila_inicio_az > fila_fin_az:
                continue # Omitir si el bloque está vacío
                
            # Aplicar estilo a todas las celdas ANTES de fusionar
            for fila in range(fila_inicio_az, fila_fin_az + 1):
                celda = ws.cell(row=fila, column=col_az_idx)
                celda.font = fuente_az
                celda.alignment = alineacion_az
                celda.fill = relleno_az
            
            # Fusionar el bloque
            if fila_inicio_az != fila_fin_az:
                ws.merge_cells(start_row=fila_inicio_az, start_column=col_az_idx, end_row=fila_fin_az, end_column=col_az_idx)
            
            # Preparar la siguiente fila de inicio (+2 para saltar la fila negra separadora)
            fila_inicio_az = fila_fin_az + 2 
            
    except Exception as e:
        print(f"Error estilizando la columna AZ vertical: {e}")
   
    # 12. Guardar en un stream de bytes en lugar de un archivo
    with io.BytesIO() as bio:
        wb.save(bio)
        return bio.getvalue()