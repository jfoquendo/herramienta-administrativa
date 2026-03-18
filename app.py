import streamlit as st
import pandas as pd
import pikepdf
import re
import io
from openpyxl.styles import PatternFill

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Multi-Tool Administrativa 2026",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS PERSONALIZADOS (CSS) ---
st.markdown("""
    <style>
    .main {
        background-color: #f5f7f9;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #007bff;
        color: white;
    }
    .stProgress .st-bo {
        background-color: #007bff;
    }
    </style>
    """, unsafe_allow_密w=True)

# --- FUNCIONES DE LÓGICA ---

def limpiar_extremo(dato):
    if pd.isna(dato): return ""
    s = str(dato).strip()
    s = re.sub(r'\.0$', '', s)
    s = re.sub(r'\D', '', s)
    return s

def unlock_pdf(file_bytes, passwords):
    try:
        try:
            with pikepdf.open(io.BytesIO(file_bytes)) as pdf:
                out = io.BytesIO(); pdf.save(out)
                return True, "Sin protección o ya liberado.", out.getvalue()
        except pikepdf.PasswordError:
            pass 
        for pw in passwords:
            try:
                with pikepdf.open(io.BytesIO(file_bytes), password=pw.strip()) as pdf:
                    out = io.BytesIO(); pdf.save(out)
                    return True, "Desbloqueado con éxito.", out.getvalue()
            except pikepdf.PasswordError:
                continue
        return False, "Contraseña incorrecta o no suministrada.", None
    except Exception as e:
        return False, f"Error técnico: {str(e)}", None

# --- MENÚ LATERAL ---
with st.sidebar:
    st.title("🛠️ Panel de Control")
    st.markdown("---")
    opcion = st.radio(
        "Seleccione una Herramienta:",
        ["📱 Agrupar Teléfonos", "🔓 Desbloqueo PDF", "👥 Cruce Empleados", "📊 Organizador Excel"],
        index=0
    )
    st.markdown("---")
    st.info("**Estado del Sistema:** Conectado ✅")
    st.caption("v3.5 - Edición 2026")

# --- CONTENIDO PRINCIPAL ---

if opcion == "📱 Agrupar Teléfonos":
    st.header("📱 Extractor de Teléfonos para WhatsApp")
    st.write("Clasifica números según las observaciones del archivo Excel.")
    
    archivo = st.file_uploader("Subir Excel Base", type=["xlsx", "xls"])
    
    if archivo:
        df_temp = pd.read_excel(archivo, nrows=5)
        st.success("Archivo cargado correctamente.")
        
        col1, col2 = st.columns(2)
        with col1:
            col_obs = st.selectbox("Columna de Observaciones:", df_temp.columns, index=0)
        with col2:
            col_tel = st.selectbox("Columna de Teléfonos:", df_temp.columns, index=1)

        if st.button("🚀 Procesar y Agrupar"):
            bar = st.progress(0)
            df = pd.read_excel(archivo)
            cats = {'Cédula':['cedula'],'Firma':['firma'],'Foto':['foto'],'Carta':['carta'],'Cesantias':['cesantias'],'EPS':['eps'],'ADRES':['adres'],'Bancario':['bancario','cuenta'],'Incompleto':['incompleto'],'Acta de Grado':['acta'],"Ruaf":['ruaf']}
            res = {c: [] for c in cats}
            
            for i, fila in df.iterrows():
                obs = str(fila.get(col_obs, '')).lower()
                tel = fila.get(col_tel)
                if pd.isna(obs) or obs in ['nan','ok',''] or pd.isna(tel): continue
                num = f"A,57{limpiar_extremo(tel)}"
                for c, pws in cats.items():
                    if any(p in obs for p in pws): res[c].append(num)
                bar.progress((i+1)/len(df))
            
            max_l = max(len(v) for v in res.values()) if res.values() else 0
            for c in res: res[c] += [None]*(max_l - len(res[c]))
            
            df_final = pd.DataFrame(res)
            st.dataframe(df_final, use_container_width=True)
            
            out = io.BytesIO()
            df_final.to_excel(out, index=False)
            st.download_button("📥 Descargar Reporte Agrupado", out.getvalue(), "Telefonos_WhatsApp.xlsx")

elif opcion == "🔓 Desbloqueo PDF":
    st.header("🔓 Desbloqueador Masivo de PDFs")
    st.write("Sube tus archivos protegidos y aplica las contraseñas conocidas.")
    
    pws_raw = st.text_input("Contraseñas (separadas por coma)", placeholder="ej: 1234, clave2026, admin")
    p_files = st.file_uploader("Selecciona los archivos PDF", type="pdf", accept_multiple_files=True)
    
    if p_files:
        if st.button("🔓 Iniciar Desbloqueo"):
            lista_pws = [p.strip() for p in pws_raw.split(',')] if pws_raw else []
            bar = st.progress(0)
            
            for i, pf in enumerate(p_files):
                ok, msg, content = unlock_pdf(pf.read(), lista_pws)
                with st.expander(f"Archivo: {pf.name}"):
                    if ok:
                        st.success(msg)
                        st.download_button(f"Descargar {pf.name}", content, f"liberado_{pf.name}", key=f"pdf_{i}")
                    else:
                        st.error(msg)
                bar.progress((i+1)/len(p_files))
            st.balloons()

elif opcion == "👥 Cruce Empleados":
    st.header("👥 Cruce Masivo de Empleados")
    st.write("Busca una lista de IDs dentro de un archivo Maestro conservando todo el formato.")
    
    c1, c2 = st.columns(2)
    with c1:
        m_f = st.file_uploader("1. Subir Archivo Maestro", type="xlsx")
    with c2:
        b_f = st.file_uploader("2. Subir Lista de Búsqueda", type="xlsx")
    
    if m_f and b_f:
        df_m_head = pd.read_excel(m_f, skiprows=1, nrows=0)
        col_id_m = st.selectbox("Selecciona columna ID del Maestro:", df_m_head.columns)
        
        if st.button("🔎 Ejecutar Cruce de Datos"):
            bar = st.progress(0)
            msg = st.empty()
            
            msg.info("Cargando y procesando datos...")
            df_a = pd.read_excel(m_f, skiprows=1)
            cols_orig = df_a.columns.tolist()
            df_b = pd.read_excel(b_f, header=None)
            
            df_a['ID_LIMPIO'] = df_a[col_id_m].apply(limpiar_extremo)
            ceds_busqueda = set([limpiar_extremo(v) for v in df_b.values.flatten() if limpiar_extremo(v) != ""])
            
            # Realizar el cruce
            encontrados = df_a[df_a['ID_LIMPIO'].isin(ceds_busqueda)].copy()
            encontrados = encontrados[cols_orig] # Mantener orden original
            
            id_hallados = set(encontrados['ID_LIMPIO'])
            faltantes = [c for c in ceds_busqueda if c not in id_hallados]
            
            bar.progress(70)
            msg.info("Generando archivos de salida...")
            
            # Excel con resaltado
            out_enc = io.BytesIO()
            with pd.ExcelWriter(out_enc, engine='openpyxl') as writer:
                encontrados.to_excel(writer, index=False)
                ws = writer.book.active
                fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                dup_mask = encontrados[col_id_m].duplicated(keep=False).tolist()
                for i, is_dup in enumerate(dup_mask):
                    if is_dup:
                        for cell in ws[i+2]: cell.fill = fill
            
            col_res1, col_res2 = st.columns(2)
            with col_res1:
                st.metric("Encontrados", len(encontrados))
                st.download_button("📥 Descargar EXISTENTES", out_enc.getvalue(), "Cruce_Existentes.xlsx")
            with col_res2:
                st.metric("No Encontrados", len(faltantes))
                if faltantes:
                    out_f = io.BytesIO()
                    pd.DataFrame(faltantes, columns=['ID_No_Encontrado']).to_excel(out_f, index=False)
                    st.download_button("📥 Descargar FALTANTES", out_f.getvalue(), "Cruce_Faltantes.xlsx")
            
            bar.progress(100)
            msg.success("Proceso de cruce terminado con éxito.")

elif opcion == "📊 Organizador Excel":
    st.header("📊 Organizador por Lista Específica")
    st.write("Reordena el archivo de datos basándose en el orden de otro archivo.")
    
    col_a, col_b = st.columns(2)
    with col_a:
        d_f = st.file_uploader("Archivo Principal (Datos)", type="xlsx")
    with col_b:
        o_f = st.file_uploader("Archivo de Referencia (Orden)", type="xlsx")
        
    if d_f and o_f:
        df_d_head = pd.read_excel(d_f, nrows=0)
        df_o_head = pd.read_excel(o_f, nrows=0)
        
        c1, c2 = st.columns(2)
        with c1:
            sel_d = st.selectbox("Columna ID en Datos:", df_d_head.columns)
        with c2:
            sel_o = st.selectbox("Columna ID en Orden:", df_temp_o_cols := df_o_head.columns)

        if st.button("⚙️ Reorganizar Documento"):
            bar = st.progress(20)
            df_datos = pd.read_excel(d_f)
            cols_orig = df_datos.columns.tolist()
            df_orden = pd.read_excel(o_f)
            
            df_datos[sel_d] = df_datos[sel_d].astype(str).str.strip()
            df_orden[sel_o] = df_orden[sel_o].astype(str).str.strip()
            
            bar.progress(60)
            # Cruce para ordenar
            df_res = pd.merge(df_orden[[sel_o]], df_datos, left_on=sel_o, right_on=sel_d, how='left')
            df_res = df_res[cols_orig]
            
            out = io.BytesIO()
            df_res.to_excel(out, index=False)
            bar.progress(100)
            st.success("¡Archivo reorganizado!")
            st.download_button("📥 Descargar Excel Organizado", out.getvalue(), "Resultado_Ordenado.xlsx")
