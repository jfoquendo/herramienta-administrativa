import streamlit as st
import pandas as pd
import pikepdf
import re
import io
from pathlib import Path

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Multi-Tool Administrativa 2026", layout="wide")

# --- FUNCIONES DE LÓGICA ---

def limpiar_extremo(dato):
    if pd.isna(dato): return ""
    s = str(dato).strip()
    s = re.sub(r'\.0$', '', s)  # Quita decimales de Excel
    s = re.sub(r'\D', '', s)     # Quita todo lo que no sea dígito
    return s

def unlock_pdf(file_bytes, passwords):
    try:
        # Intentar abrir sin contraseña primero
        try:
            with pikepdf.open(io.BytesIO(file_bytes)) as pdf:
                output = io.BytesIO()
                pdf.save(output)
                return True, "Sin protección.", output.getvalue()
        except pikepdf.PasswordError:
            pass 

        # Probar con la lista de contraseñas
        for pw in passwords:
            try:
                with pikepdf.open(io.BytesIO(file_bytes), password=pw.strip()) as pdf:
                    output = io.BytesIO()
                    pdf.save(output)
                    return True, "Desbloqueado.", output.getvalue()
            except pikepdf.PasswordError:
                continue
        return False, "Contraseña incorrecta.", None
    except Exception as e:
        return False, f"Error: {str(e)}", None

# --- INTERFAZ DE USUARIO ---

st.title("⚙️ Multi-Tool Administrativa v3.0")
st.info("Versión Web Cloud - 2026")

tab1, tab2, tab3, tab4 = st.tabs([
    "📱 Agrupar Teléfonos", 
    "🔓 Desbloquear PDFs", 
    "👥 Cruce Empleados", 
    "📊 Organizador Excel"
])

# --- PESTAÑA 1: AGRUPAR TELÉFONOS ---
with tab1:
    st.header("📱 Extractor para WhatsApp")
    st.write("Agrupa números telefónicos por categoría según la columna 'VALIDACION DE DOCUMENTOS'.")
    
    archivo_base = st.file_uploader("Cargar Excel Base", type=["xlsx", "xls"], key="agr_tel")
    
    if archivo_base:
        df = pd.read_excel(archivo_base)
        categorias_keywords = {
            'Cédula': ['cedula'], 'Firma': ['firma'], 'Foto': ['foto'],
            'Carta': ['carta'], 'Cesantias': ['cesantias'], 'EPS': ['eps'],
            'ADRES': ['adres'], 'Bancario': ['bancario', 'cuenta'],
            'Incompleto': ['incompleto'], 'Acta de Grado': ['acta'], "Ruaf": ['ruaf']
        }
        
        col_obs = 'VALIDACION DE DOCUMENTOS'
        col_tel = 'TELEFONO CELULAR'

        if col_obs in df.columns and col_tel in df.columns:
            data_columnas = {cat: [] for cat in categorias_keywords}
            for _, fila in df.iterrows():
                obs = str(fila.get(col_obs, '')).lower()
                tel = fila.get(col_tel)
                if pd.isna(obs) or obs in ['nan', 'ok', ''] or pd.isna(tel): continue
                
                num = f"A,57{limpiar_extremo(tel)}"
                for cat, palabras in categorias_keywords.items():
                    if any(p in obs for p in palabras):
                        data_columnas[cat].append(num)

            st.success("✅ Procesamiento completado")
            
            # Ajustar longitudes para el DataFrame
            max_len = max([len(v) for v in data_columnas.values()]) if data_columnas.values() else 0
            for cat in data_columnas: 
                data_columnas[cat] += [None] * (max_len - len(data_columnas[cat]))
            
            res_df = pd.DataFrame(data_columnas)
            st.dataframe(res_df, use_container_width=True)
            
            # Botón de descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Descargar Resultado", output.getvalue(), "Telefonos_Agrupados.xlsx")
        else:
            st.error(f"⚠️ No se encontraron las columnas '{col_obs}' o '{col_tel}'")

# --- PESTAÑA 2: DESBLOQUEAR PDF ---
with tab2:
    st.header("🔓 Desbloqueador de PDFs")
    pws_input = st.text_input("Ingresa las contraseñas posibles (separadas por coma)")
    uploaded_pdfs = st.file_uploader("Sube uno o varios archivos PDF", type="pdf", accept_multiple_files=True)
    
    if uploaded_pdfs and st.button("Iniciar Desbloqueo"):
        pws = [p.strip() for p in pws_input.split(',')] if pws_input else []
        for pdf_file in uploaded_pdfs:
            ok, msg, content = unlock_pdf(pdf_file.read(), pws)
            if ok:
                st.success(f"✅ {pdf_file.name}: {msg}")
                st.download_button(f"Descargar {pdf_file.name} desbloqueado", content, f"unlocked_{pdf_file.name}")
            else:
                st.error(f"❌ {pdf_file.name}: {msg}")

# --- PESTAÑA 3: CRUCE EMPLEADOS ---
with tab3:
    st.header("👥 Cruce Masivo de Empleados")
    c1, c2 = st.columns(2)
    with c1: m_file = st.file_uploader("Archivo Maestro (Activos)", type=["xlsx"], key="m_cruce")
    with c2: b_file = st.file_uploader("Lista de Búsqueda (Cédulas)", type=["xlsx"], key="b_cruce")

    if m_file and b_file and st.button("Ejecutar Cruce"):
        # Leemos saltando la primera fila como en tu código original
        df_activos = pd.read_excel(m_file, skiprows=1)
        df_busqueda = pd.read_excel(b_file, header=None)
        
        col_cedula = next((c for c in df_activos.columns if any(x in str(c).lower() for x in ['ced', 'doc', 'id'])), df_activos.columns[0])
        df_activos['ID_LIMPIO'] = df_activos[col_cedula].apply(limpiar_extremo)
        
        busqueda_valores = df_busqueda.values.flatten()
        cedulas_a_buscar = set([limpiar_extremo(v) for v in busqueda_valores if limpiar_extremo(v) != ""])
        
        encontrados = df_activos[df_activos['ID_LIMPIO'].isin(cedulas_a_buscar)].copy()
        id_hallados = set(encontrados['ID_LIMPIO'])
        faltantes = [c for c in cedulas_a_buscar if c not in id_hallados]

        st.write(f"🔎 Cédulas buscadas: **{len(cedulas_a_buscar)}**")
        st.write(f"✅ Encontradas: **{len(encontrados)}**")
        st.write(f"⚠️ Faltantes: **{len(faltantes)}**")

        if not encontrados.empty:
            out_ok = io.BytesIO()
            encontrados.to_excel(out_ok, index=False)
            st.download_button("📥 Descargar EXISTENTES", out_ok.getvalue(), "1_EXISTENTES.xlsx")
        
        if faltantes:
            out_fail = io.BytesIO()
            pd.DataFrame(faltantes, columns=['Cédula_No_Encontrada']).to_excel(out_fail, index=False)
            st.download_button("📥 Descargar NO ENCONTRADOS", out_fail.getvalue(), "2_NO_ENCONTRADOS.xlsx")

# --- PESTAÑA 4: ORGANIZADOR EXCEL ---
with tab4:
    st.header("📊 Organizador por Orden Específico")
    f_datos = st.file_uploader("Archivo de Datos Completo", type=["xlsx"], key="org_d")
    c_datos = st.text_input("Columna ID en Datos", "Número de identificación")
    
    f_orden = st.file_uploader("Archivo con el ORDEN deseado", type=["xlsx"], key="org_o")
    c_orden = st.text_input("Columna ID en Orden", "Número de identificación")

    if f_datos and f_orden and st.button("Generar Excel Organizado"):
        df_d = pd.read_excel(f_datos)
        df_o = pd.read_excel(f_orden)
        
        # Limpiar columnas
        df_d.columns = [str(c).strip() for c in df_d.columns]
        df_o.columns = [str(c).strip() for c in df_o.columns]
        
        if c_datos in df_d.columns and c_orden in df_o.columns:
            df_d[c_datos] = df_d[c_datos].astype(str).str.strip()
            df_o[c_orden] = df_o[c_orden].astype(str).str.strip()
            
            df_d = df_d.drop_duplicates(subset=[c_datos])
            df_final = pd.merge(df_o[[c_orden]], df_d, left_on=c_orden, right_on=c_datos, how='left')
            
            out_org = io.BytesIO()
            df_final.to_excel(out_org, index=False)
            st.success("✅ Excel organizado con éxito")
            st.download_button("📥 Descargar Excel Organizado", out_org.getvalue(), "Reporte_Ordenado.xlsx")
        else:
            st.error("❌ No se encontraron las columnas de ID especificadas.")
