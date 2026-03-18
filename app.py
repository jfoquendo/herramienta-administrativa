import streamlit as st
import pandas as pd
import pikepdf
import re
import io
from openpyxl.styles import PatternFill

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Multi-Tool Administrativa 2026", layout="wide")

def limpiar_extremo(dato):
    if pd.isna(dato): return ""
    s = str(dato).strip()
    s = re.sub(r'\.0$', '', s)
    s = re.sub(r'\D', '', s)
    return s

# ... (Funciones de PDF y Teléfonos se mantienen igual) ...

# --- INTERFAZ ---
st.title("⚙️ Multi-Tool Administrativa v3.0")
tab1, tab2, tab3, tab4 = st.tabs(["📱 Teléfonos", "🔓 PDFs", "👥 Cruce", "📊 Organizador"])

# --- PESTAÑA 3: CRUCE (Mejorada) ---
with tab3:
    st.header("👥 Cruce de Empleados")
    m_f = st.file_uploader("1. Archivo Maestro (Activos)", type="xlsx", key="upload_m")
    
    col_cedula = None
    if m_f:
        # Leemos solo las columnas primero para que sea rápido
        df_temp = pd.read_excel(m_f, skiprows=1, nrows=0)
        col_cedula = st.selectbox("Selecciona la columna de Cédula en el Maestro:", df_temp.columns)

    b_f = st.file_uploader("2. Lista Búsqueda (Cédulas a encontrar)", type="xlsx", key="upload_b")
    
    if m_f and b_f and col_cedula and st.button("Ejecutar Cruce"):
        bar = st.progress(0); msg = st.empty()
        
        msg.text("Cargando Maestro...")
        df_a = pd.read_excel(m_f, skiprows=1); bar.progress(20)
        cols_originales = df_a.columns.tolist()
        
        msg.text("Cargando Búsqueda...")
        df_b = pd.read_excel(b_f, header=None); bar.progress(40)
        
        # Limpieza
        df_a['ID_LIMPIO'] = df_a[col_cedula].apply(limpiar_extremo)
        ceds = set([limpiar_extremo(v) for v in df_b.values.flatten() if limpiar_extremo(v) != ""])
        
        enc = df_a[df_a['ID_LIMPIO'].isin(ceds)].copy()
        enc = enc[cols_originales] 
        
        faltantes = [c for c in ceds if c not in set(df_a['ID_LIMPIO'])]
        
        # Generar Excel
        out_enc = io.BytesIO()
        with pd.ExcelWriter(out_enc, engine='openpyxl') as writer:
            enc.to_excel(writer, index=False)
            ws = writer.book.active
            fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            # Usamos la columna seleccionada para marcar duplicados
            dup_mask = enc[col_cedula].duplicated(keep=False).tolist()
            for i, is_dup in enumerate(dup_mask):
                if is_dup:
                    for cell in ws[i+2]: cell.fill = fill

        st.download_button("📥 Descargar EXISTENTES", out_enc.getvalue(), "Existentes.xlsx")
        if faltantes:
            out_f = io.BytesIO()
            pd.DataFrame(faltantes, columns=['Cédula_No_Encontrada']).to_excel(out_f, index=False)
            st.download_button("📥 Descargar FALTANTES", out_f.getvalue(), "Faltantes.xlsx")
        bar.progress(100); msg.success("Cruce finalizado")

# --- PESTAÑA 4: ORGANIZADOR (Mejorada) ---
with tab4:
    st.header("📊 Organizador de Excel")
    d_f = st.file_uploader("1. Archivo de Datos (Toda la info)", type="xlsx", key="org_data")
    
    sel_d = None
    if d_f:
        df_temp_d = pd.read_excel(d_f, nrows=0)
        sel_d = st.selectbox("Selecciona columna ID en el archivo de Datos:", df_temp_d.columns)

    o_f = st.file_uploader("2. Archivo de Orden (IDs en el orden deseado)", type="xlsx", key="org_order")
    
    sel_o = None
    if o_f:
        df_temp_o = pd.read_excel(o_f, nrows=0)
        sel_o = st.selectbox("Selecciona columna ID en el archivo de Orden:", df_temp_o.columns)

    if d_f and o_f and sel_d and sel_o and st.button("Organizar Excel"):
        bar = st.progress(0); msg = st.empty()
        
        df_datos = pd.read_excel(d_f); bar.progress(30)
        cols_originales = df_datos.columns.tolist()
        
        df_orden = pd.read_excel(o_f); bar.progress(50)
        
        # Procesamiento seguro
        df_datos[sel_d] = df_datos[sel_d].astype(str).str.strip()
        df_orden[sel_o] = df_orden[sel_o].astype(str).str.strip()
        
        df_res = pd.merge(df_orden[[sel_o]], df_datos, left_on=sel_o, right_on=sel_d, how='left')
        
        # Mantener solo columnas originales y el orden del archivo de entrada
        df_res = df_res[cols_originales]
        
        out = io.BytesIO()
        df_res.to_excel(out, index=False)
        bar.progress(100); msg.success("Organización completa")
        st.download_button("📥 Descargar Organizado", out.getvalue(), "Organizado.xlsx")
