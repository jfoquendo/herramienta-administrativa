import pandas as pd
import pikepdf
import threading
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

# --- FUNCIONES DE UTILIDAD PARA CRUCE ---

def limpiar_extremo(dato):
    if pd.isna(dato): return ""
    s = str(dato).strip()
    s = re.sub(r'\.0$', '', s)  # Quita decimales de Excel
    s = re.sub(r'\D', '', s)     # Quita todo lo que no sea dígito
    return s

# --- LÓGICA DE AGRUPACIÓN DE TELÉFONOS (Pestaña 1) ---

def agrupar_telefonos_logic(ruta_excel, txt_widget):
    try:
        df = pd.read_excel(ruta_excel)
        categorias_keywords = {
            'Cédula': ['cedula'], 
            'Firma': ['firma'], 
            'Foto': ['foto'],
            'Carta': ['carta'], 
            'Certificado': ['certificado'], 
            'EPS': ['eps'],
            'ADRES': ['adres'], 
            'Bancario': ['bancario', 'cuenta'],
            'Incompleto': ['incompleto'], 
            'Acta de Grado': ['acta'], 
            "Ruaf": ['ruaf']
        }
        data_columnas = {cat: [] for cat in categorias_keywords}
        col_obs = 'VALIDACION DE DOCUMENTOS'
        col_tel = 'TELEFONO CELULAR'

        if col_obs not in df.columns or col_tel not in df.columns:
            txt_widget.insert(tk.END, f"ERROR: No se encontraron las columnas requeridas.\n")
            return

        for _, fila in df.iterrows():
            obs = str(fila.get(col_obs, '')).lower()
            tel = fila.get(col_tel)
            if pd.isna(obs) or obs in ['nan', 'ok', ''] or pd.isna(tel): continue
            
            num = f"A,57{limpiar_extremo(tel)}"
            for cat, palabras in categorias_keywords.items():
                if any(p in obs for p in palabras):
                    data_columnas[cat].append(num)

        txt_widget.delete(1.0, tk.END)
        for cat, lista in data_columnas.items():
            if lista:
                txt_widget.insert(tk.END, f"\n--- {cat.upper()} ({len(lista)}) ---\n" + "\n".join(lista) + "\n")
        
        archivo_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if archivo_salida:
            max_len = max([len(v) for v in data_columnas.values()]) if data_columnas.values() else 0
            for cat in data_columnas: data_columnas[cat] += [None] * (max_len - len(data_columnas[cat]))
            pd.DataFrame(data_columnas).to_excel(archivo_salida, index=False)
            messagebox.showinfo("Éxito", "Archivo de teléfonos guardado.")
    except Exception as e:
        txt_widget.insert(tk.END, f"\n[ERROR]: {str(e)}\n")

# --- LÓGICA DE DESBLOQUEO PDF (Pestaña 2) ---

def unlock_pdf(src, dst, passwords):
    try:
        try:
            with pikepdf.open(src) as pdf:
                return True, "Sin protección."
        except pikepdf.PasswordError:
            pass 
        for pw in passwords:
            try:
                with pikepdf.open(src, password=pw.strip()) as pdf:
                    pdf.save(dst); return True, "Desbloqueado."
            except pikepdf.PasswordError:
                continue
        return False, "Contraseña incorrecta."
    except Exception as e: return False, f"Error: {str(e)}"

# --- LÓGICA DE CRUCE DE EMPLEADOS (Pestaña 3) ---

def cruce_empleados_logic(ruta_maestro, ruta_busqueda, txt_widget):
    try:
        txt_widget.delete(1.0, tk.END)
        txt_widget.insert(tk.END, "🚀 Cargando archivos para el cruce...\n")
        df_activos = pd.read_excel(ruta_maestro, skiprows=1)
        df_activos.columns = [str(c).strip() for c in df_activos.columns]
        df_busqueda = pd.read_excel(ruta_busqueda, header=None)
        col_cedula = next((c for c in df_activos.columns if any(x in c.lower() for x in ['ced', 'doc', 'id'])), df_activos.columns[0])
        df_activos['ID_LIMPIO'] = df_activos[col_cedula].apply(limpiar_extremo)
        busqueda_valores = df_busqueda.values.flatten()
        cedulas_a_buscar = set([limpiar_extremo(v) for v in busqueda_valores if limpiar_extremo(v) != ""])
        encontrados = df_activos[df_activos['ID_LIMPIO'].isin(cedulas_a_buscar)].copy()
        id_hallados = set(encontrados['ID_LIMPIO'])
        faltantes = [c for c in cedulas_a_buscar if c not in id_hallados]

        txt_widget.insert(tk.END, f"🔎 Cédulas únicas buscadas: {len(cedulas_a_buscar)}\n")
        txt_widget.insert(tk.END, f"✅ Encontrados: {len(encontrados)}\n")
        txt_widget.insert(tk.END, f"⚠️ Faltantes: {len(faltantes)}\n\n")

        if not encontrados.empty:
            cols_ok = [c for c in [col_cedula, 'Cod_proveedor', 'Nombre', 'Cta_Bancaria', 'Sociedad_FI', 'FUNCION_BLOQUEADA'] if c in df_activos.columns]
            encontrados[cols_ok].to_excel('1_EXISTENTES_COMPLETO.xlsx', index=False)
            txt_widget.insert(tk.END, "✔️ Generado: 1_EXISTENTES_COMPLETO.xlsx\n")
        if faltantes:
            pd.DataFrame(faltantes, columns=['Cédula_No_Encontrada']).to_excel('2_NO_ENCONTRADOS.xlsx', index=False)
            txt_widget.insert(tk.END, "✔️ Generado: 2_NO_ENCONTRADOS.xlsx\n")
        messagebox.showinfo("Cruce Finalizado", "Se han generado los archivos de reporte.")
    except Exception as e:
        txt_widget.insert(tk.END, f"\n[ERROR EN CRUCE]: {str(e)}\n")

# --- INTERFAZ PRINCIPAL ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Tool Administrativa 2026")
        self.root.geometry("950x750")
        self.root.configure(bg="#1e1e1e")

        style = ttk.Style()
        style.theme_use('default')
        style.configure("TNotebook", background="#1e1e1e", borderwidth=0)
        style.configure("TNotebook.Tab", background="#333333", foreground="white", padding=[15, 5])
        style.map("TNotebook.Tab", background=[("selected", "#4CAF50")], foreground=[("selected", "white")])
        
        # Estilos para la nueva pestaña (Organizador)
        style.configure("TFrame", background="#1e1e1e")
        style.configure("TLabel", background="#1e1e1e", foreground="white")
        style.configure("TEntry", fieldbackground="#333333", foreground="white")

        self.tab_control = ttk.Notebook(root)
        
        self.tab1 = tk.Frame(self.tab_control, bg="#1e1e1e")
        self.tab2 = tk.Frame(self.tab_control, bg="#1e1e1e")
        self.tab3 = tk.Frame(self.tab_control, bg="#1e1e1e")
        self.tab4 = tk.Frame(self.tab_control, bg="#1e1e1e")

        self.tab_control.add(self.tab1, text='📱 Agrupar Teléfonos')
        self.tab_control.add(self.tab2, text='🔓 Desbloquear PDFs')
        self.tab_control.add(self.tab3, text='👥 Cruce Empleados')
        self.tab_control.add(self.tab4, text='📊 Organizador Excel')
        self.tab_control.pack(expand=1, fill="both")

        self.setup_tab_agrupador()
        self.setup_tab_pdf()
        self.setup_tab_cruce()
        self.setup_tab_organizador()

    # --- PESTAÑA 1, 2, 3 (Igual a tu código original) ---

    def setup_tab_agrupador(self):
        lbl = tk.Label(self.tab1, text="EXTRACTOR PARA WHATSAPP (VALIDACIÓN)", fg="#4CAF50", bg="#1e1e1e", font=("Consolas", 12, "bold"))
        lbl.pack(pady=10)
        tk.Button(self.tab1, text="CARGAR EXCEL", command=self.run_agrupar, bg="#4CAF50", fg="white", width=25).pack(pady=5)
        self.consola_agr = scrolledtext.ScrolledText(self.tab1, bg="black", fg="#00FF00", font=("Consolas", 10))
        self.consola_agr.pack(padx=10, pady=10, fill="both", expand=True)

    def run_agrupar(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if f: threading.Thread(target=agrupar_telefonos_logic, args=(f, self.consola_agr), daemon=True).start()

    def setup_tab_pdf(self):
        frm = tk.Frame(self.tab2, bg="#1e1e1e", padx=20, pady=20)
        frm.pack(fill="x")
        self.ent_pdf_dir = self.crear_campo(frm, "Carpeta PDFs:", 0)
        self.ent_pw = self.crear_campo(frm, "Contraseñas (comas):", 1)
        self.var_rec = tk.BooleanVar(value=True)
        tk.Checkbutton(frm, text="Incluir subcarpetas", variable=self.var_rec, bg="#1e1e1e", fg="white", selectcolor="#333").grid(row=2, column=1, sticky="w")
        tk.Button(frm, text="DESBLOQUEAR AHORA", command=self.run_pdf, bg="#2196F3", fg="white", font=("Bold")).grid(row=3, column=1, pady=15)
        self.consola_pdf = scrolledtext.ScrolledText(self.tab2, bg="black", fg="#00FF00", font=("Consolas", 10))
        self.consola_pdf.pack(padx=10, pady=5, fill="both", expand=True)

    def run_pdf(self):
        root_dir = self.ent_pdf_dir.get()
        pws = self.ent_pw.get().split(',')
        if not root_dir: return
        def worker():
            files = list(Path(root_dir).rglob("*.pdf")) if self.var_rec.get() else list(Path(root_dir).glob("*.pdf"))
            for f in files:
                if "_unlocked" in f.name: continue
                ok, msg = unlock_pdf(f, f.with_name(f.stem + "_unlocked.pdf"), pws)
                self.consola_pdf.insert(tk.END, f"[{'OK' if ok else 'FAIL'}] {f.name}: {msg}\n")
                self.consola_pdf.see(tk.END)
        threading.Thread(target=worker, daemon=True).start()

    def setup_tab_cruce(self):
        frm = tk.Frame(self.tab3, bg="#1e1e1e", padx=20, pady=20)
        frm.pack(fill="x")
        tk.Label(frm, text="Cruce Masivo de Empleados", fg="#FF9800", bg="#1e1e1e", font=("Consolas", 14, "bold")).grid(row=0, column=0, columnspan=3, pady=10)
        self.ent_maestro = self.crear_campo(frm, "Archivo Maestro (Activos):", 1)
        self.ent_busqueda = self.crear_campo(frm, "Lista de Búsqueda:", 2)
        tk.Button(frm, text="EJECUTAR CRUCE", command=self.run_cruce, bg="#FF9800", fg="black", font=("Consolas", 11, "bold"), width=30).grid(row=3, column=1, pady=20)
        self.consola_cruce = scrolledtext.ScrolledText(self.tab3, bg="black", fg="#FF9800", font=("Consolas", 10))
        self.consola_cruce.pack(padx=10, pady=5, fill="both", expand=True)

    def run_cruce(self):
        m, b = self.ent_maestro.get(), self.ent_busqueda.get()
        if not m or not b:
            messagebox.showwarning("Faltan archivos", "Selecciona ambos archivos.")
            return
        threading.Thread(target=cruce_empleados_logic, args=(m, b, self.consola_cruce), daemon=True).start()

    # --- NUEVA PESTAÑA 4: ORGANIZADOR DE EXCEL ---

    def setup_tab_organizador(self):
        container = tk.Frame(self.tab4, bg="#1e1e1e", padx=30, pady=30)
        container.pack(fill="both", expand=True)

        tk.Label(container, text="ORGANIZADOR POR LISTA ESPECÍFICA", fg="#E91E63", bg="#1e1e1e", font=("Consolas", 14, "bold")).pack(pady=(0,20))

        # Sección 1
        tk.Label(container, text="1. Archivo Base (Toda la información):", fg="white", bg="#1e1e1e", font=("Arial", 10, "bold")).pack(anchor="w")
        f1 = tk.Frame(container, bg="#1e1e1e")
        f1.pack(fill="x", pady=5)
        self.ruta_datos = tk.Entry(f1, bg="#333", fg="white", insertbackground="white")
        self.ruta_datos.pack(side="left", fill="x", expand=True, padx=(0,5))
        tk.Button(f1, text="Buscar", command=lambda: self.seleccionar_generico(self.ruta_datos)).pack(side="right")

        tk.Label(container, text="Nombre de columna ID en Base:", fg="#bbb", bg="#1e1e1e").pack(anchor="w")
        self.col_datos = tk.Entry(container, bg="#333", fg="white", insertbackground="white")
        self.col_datos.insert(0, "Número de identificación")
        self.col_datos.pack(fill="x", pady=(0,15))

        # Sección 2
        tk.Label(container, text="2. Archivo con ORDEN deseado:", fg="white", bg="#1e1e1e", font=("Arial", 10, "bold")).pack(anchor="w")
        f2 = tk.Frame(container, bg="#1e1e1e")
        f2.pack(fill="x", pady=5)
        self.ruta_lista = tk.Entry(f2, bg="#333", fg="white", insertbackground="white")
        self.ruta_lista.pack(side="left", fill="x", expand=True, padx=(0,5))
        tk.Button(f2, text="Buscar", command=lambda: self.seleccionar_generico(self.ruta_lista)).pack(side="right")

        tk.Label(container, text="Nombre de columna ID en Lista:", fg="#bbb", bg="#1e1e1e").pack(anchor="w")
        self.col_lista = tk.Entry(container, bg="#333", fg="white", insertbackground="white")
        self.col_lista.insert(0, "Número de identificación")
        self.col_lista.pack(fill="x", pady=(0,15))

        # Botón Acción
        tk.Button(container, text="🚀 GENERAR EXCEL ORGANIZADO", command=self.procesar_organizador, 
                  bg="#E91E63", fg="white", font=("Consolas", 11, "bold"), pady=10).pack(fill="x", pady=20)

    def seleccionar_generico(self, entry):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def procesar_organizador(self):
        if not self.ruta_datos.get() or not self.ruta_lista.get():
            messagebox.showerror("Error", "Selecciona ambos archivos.")
            return
        try:
            df_datos = pd.read_excel(self.ruta_datos.get())
            df_orden = pd.read_excel(self.ruta_lista.get())
            columnas_originales = list(df_datos.columns)
            
            df_datos.columns = [str(c).strip() for c in df_datos.columns]
            df_orden.columns = [str(c).strip() for c in df_orden.columns]
            columnas_originales_limpias = [str(c).strip() for c in columnas_originales]
            df_datos.columns = columnas_originales_limpias

            c_datos = self.col_datos.get().strip()
            c_orden = self.col_lista.get().strip()

            if c_datos not in df_datos.columns or c_orden not in df_orden.columns:
                messagebox.showerror("Error", "Columnas ID no encontradas.")
                return

            df_datos[c_datos] = df_datos[c_datos].astype(str).str.strip()
            df_orden[c_orden] = df_orden[c_orden].astype(str).str.strip()
            df_datos = df_datos.drop_duplicates(subset=[c_datos])
            
            df_final = pd.merge(df_orden[[c_orden]], df_datos, left_on=c_orden, right_on=c_datos, how='left')
            if c_orden != c_datos: df_final = df_final.drop(columns=[c_orden])
            
            df_final = df_final[columnas_originales_limpias]
            ruta_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Reporte_Ordenado.xlsx")
            if ruta_salida:
                df_final.to_excel(ruta_salida, index=False)
                messagebox.showinfo("¡Éxito!", "Archivo generado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # --- UTILIDADES INTERFAZ ---

    def crear_campo(self, parent, texto, fila):
        tk.Label(parent, text=texto, fg="white", bg="#1e1e1e").grid(row=fila, column=0, sticky="w")
        ent = tk.Entry(parent, width=50, bg="#333", fg="white", insertbackground="white")
        ent.grid(row=fila, column=1, padx=5, pady=5)
        tk.Button(parent, text="...", command=lambda: self.seleccionar_archivo_o_dir(ent, texto)).grid(row=fila, column=2)
        return ent

    def seleccionar_archivo_o_dir(self, entry, texto):
        if "Carpeta" in texto: res = filedialog.askdirectory()
        else: res = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if res:
            entry.delete(0, tk.END)
            entry.insert(0, res)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()