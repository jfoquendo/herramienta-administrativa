import streamlit as st
import pandas as pd
import pikepdf
import re
import io
import time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Multi-Tool Administrativa 2026", layout="wide")

def limpiar_extremo(dato):
    if pd.isna(dato): return ""
    s = str(dato).strip()
    s = re.sub(r'\.0$', '', s)
    s = re.sub(r'\D', '', s)
    return s

def unlock_pdf(file_bytes, passwords):
    try:
        with pikepdf.open(io.BytesIO(file_bytes)) as pdf:
            out = io.BytesIO(); pdf.save(out)
            return True, "Sin protección.", out.getvalue()
    except pikepdf.PasswordError:
        for pw in passwords:
            try:
                with pikepdf.open(io.BytesIO(file_bytes), password=pw.strip()) as pdf:
                    out = io.BytesIO(); pdf.save(out)
                    return True, "Desbloqueado.", out.getvalue()
            except pikepdf.PasswordError: continue
        return False, "Contraseña incorrecta.", None
    except Exception as e: return False, f"Error: {str(e)}", None

# --- INTERFAZ ---
st.title("⚙️ Multi-Tool Administrativa v3.0")

tab1, tab2, tab3, tab4 = st.tabs(["📱 Teléfonos", "🔓 PDFs", "👥 Cruce", "📊 Organizador"])

# --- PESTAÑA 1: TELÉFONOS ---
with tab1:
    st.header("📱 Extractor WhatsApp")
    f = st.file_uploader("Excel Base", type=["xlsx", "xls"], key="t1")
    if f and st.button("Procesar Teléfonos"):
        bar = st.progress(0); msg = st.empty()
        msg.text("Leyendo datos..."); df = pd.read_excel(f); bar.progress(30)
        
        cats = {'Cédula':['cedula'],'Firma':['firma'],'Foto':['foto'],'Carta':['carta'],'Cesantias':['cesantias'],'EPS':['eps'],'ADRES':['adres'],'Bancario':['bancario','cuenta'],'Incompleto':['incompleto'],'Acta de Grado':['acta'],"Ruaf":['ruaf']}
        res = {c: [] for c in cats}
        
        for i, fila in df.iterrows():
            obs = str(fila.get('VALIDACION DE DOCUMENTOS', '')).lower()
            tel = fila.get('TELEFONO CELULAR')
            if pd.isna(obs) or obs in ['nan','ok',''] or pd.isna(tel): continue
            num = f"A,57{limpiar_extremo(tel)}"
            for c, pws in cats.items():
                if any(p in obs for p in pws): res[c].append(num)
        
        bar.progress(80); msg.text("Formateando tabla...")
        max_l = max(len(v) for v in res.values()) if res.values() else 0
        for c in res: res[c] += [None]*(max_l - len(res[c]))
        
        out = io.BytesIO()
        pd.DataFrame(res).to_excel(out, index=False)
        bar.progress(100); msg.success("¡Listo!")
        st.download_button("📥 Descargar", out.getvalue(), "Telefonos.xlsx")

# --- PESTAÑA 2: PDFS ---
with tab2:
    st.header("🔓 Desbloqueo Masivo")
    pws = st.text_input("Contraseñas (separadas por coma)")
    p_files = st.file_uploader("PDFs", type="pdf", accept_multiple_files=True)
    if p_files and st.button("Desbloquear"):
        bar = st.progress(0); msg = st.empty()
        list_p = [p.strip() for p in pws.split(',')]
        for i, pf in enumerate(p_files):
            msg.text(f"Desbloqueando: {pf.name}")
            ok, m, content = unlock_pdf(pf.read(), list_p)
            if ok: st.download_button(f"📥 {pf.name}", content, f"unlocked_{pf.name}")
            else: st.error(f"{pf.name}: {m}")
            bar.progress((i+1)/len(p_files))
        st.balloons()

# --- PESTAÑA 3: CRUCE (Mantiene Orden) ---
with tab3:
    st.header("👥 Cruce de Empleados")
    m_f = st.file_uploader("Maestro (Activos)", type="xlsx")
    b_f = st.file_uploader("Lista Búsqueda", type="xlsx")
    if m_f and b_f and st.button("Ejecutar Cruce"):
        bar = st.progress(0); msg = st.empty()
        msg.text("Cargando Maestro..."); df_a = pd.read_excel(m_f, skiprows=1); bar.progress(20)
        cols_originales = df_a.columns.tolist() # GUARDAMOS EL ORDEN ORIGINAL
        
        msg.text("Cargando Búsqueda..."); df_b = pd.read_excel(b_f, header=None); bar.progress(40)
        
        c_ced = next((c for c in df_a.columns if any(x in str(c).lower() for x in ['ced','doc','id'])), df_a.columns[0])
        df_a['ID_LIMPIO'] = df_a[c_ced].apply(limpiar_extremo)
        ceds = set([limpiar_extremo(v) for v in df_b.values.flatten() if limpiar_extremo(v) != ""])
        
        enc = df_a[df_a['ID_LIMPIO'].isin(ceds)].copy()
        # REORDENAR COLUMNAS AL ESTADO ORIGINAL
        enc = enc[cols_originales] 
        
        faltantes = [c for c in ceds if c not in set(df_a['ID_LIMPIO'])]
        bar.progress(80); msg.text("Generando archivos...")
        
        # Generar Excel con colores para duplicados
        out_enc = io.BytesIO()
        with pd.ExcelWriter(out_enc, engine='openpyxl') as writer:
            enc.to_excel(writer, index=False)
            # Lógica de color de tu código original
            ws = writer.book.active
            fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            dup_mask = enc[c_ced].duplicated(keep=False).tolist()
            for i, is_dup in enumerate(dup_mask):
                if is_dup:
                    for cell in ws[i+2]: cell.fill = fill

        st.download_button("📥 Descargar EXISTENTES (Color en duplicados)", out_enc.getvalue(), "Existentes.xlsx")
        if faltantes:
            out_f = io.BytesIO()
            pd.DataFrame(faltantes, columns=['Cédula_No_Encontrada']).to_excel(out_f, index=False)
            st.download_button("📥 Descargar FALTANTES", out_f.getvalue(), "Faltantes.xlsx")
        bar.progress(100); msg.success("Cruce finalizado")

# --- PESTAÑA 4: ORGANIZADOR (Mantiene Orden) ---
with tab4:
    st.header("📊 Organizador")
    d_f = st.file_uploader("Archivo de Datos", type="xlsx")
    c_d = st.text_input("Columna ID Datos", "Número de identificación")
    o_f = st.file_uploader("Archivo de Orden", type="xlsx")
    c_o = st.text_input("Columna ID Orden", "Número de identificación")

    if d_f and o_f and st.button("Organizar"):
        bar = st.progress(0); msg = st.empty()
        df_datos = pd.read_excel(d_f); bar.progress(30)
        cols_originales = df_datos.columns.tolist() # GUARDAMOS ORDEN
        
        df_orden = pd.read_excel(o_f); bar.progress(50)
        
        df_datos[c_d] = df_datos[c_d].astype(str).str.strip()
        df_orden[c_o] = df_orden[c_o].astype(str).str.strip()
        
        # Merge manteniendo la estructura del archivo de orden
        df_res = pd.merge(df_orden[[c_o]], df_datos, left_on=c_o, right_on=c_d, how='left')
        
        # Eliminar la columna extra si los nombres eran distintos y reordenar
        if c_o != c_d and c_o in df_res.columns: df_res = df_res.drop(columns=[c_o])
        df_res = df_res[cols_originales]
        
        out = io.BytesIO()
        df_res.to_excel(out, index=False)
        bar.progress(100); msg.success("Organización completa")
        st.download_button("📥 Descargar Organizado", out.getvalue(), "Organizado.xlsx")
