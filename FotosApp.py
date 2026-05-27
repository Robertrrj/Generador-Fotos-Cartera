import streamlit as st
import pandas as pd
import io
import os
import zipfile
import xlsxwriter
from datetime import datetime

st.set_page_config(page_title="Fotos Cartera", page_icon="📊", layout="wide")

# --- ESTILOS CSS ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #dee2e6; }
    .stButton>button { border-radius: 8px; background-color: #1e3a8a; color: white; font-weight: bold; }
    .stDownloadButton>button { background-color: #10b981 !important; color: white !important; border-radius: 8px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- CONFIGURACIÓN ---
CONCEPTOS_FILAS = [
    "Pagos Por Legalizar", "Auditado Para Pago", "En Auditoria ADRES", "Pruebas Covid 19", 
    "Auditado para Pago Pptos Maximos", "Valor Glosas", "Devoluciones", "No Radicadas en EPS", 
    "En Proceso de Auditoria", "Otros Conceptos", "Abogados Externos", "Pagos No Aplicados", 
    "Alistamiento", "Sin Tipificación", "Importe pendiente"
]

ESTRUCTURA_FINAL = [
    "Nº ppal", "Nombre principal", "Referencia", "Amaho", "Fac_open", "Fecha factura", 
    "Año Fact", "Mes Fact", "Importe pendiente", "Pagos Por Legalizar", "Auditado Para Pago", 
    "En Auditoria ADRES", "Pruebas Covid 19", "Auditado para Pago Pptos Maximos", "Valor Glosas", 
    "Devoluciones", "No Radicadas en EPS", "En Proceso de Auditoria", "Otros Conceptos", 
    "Abogados Externos", "Pagos No Aplicados", "Alistamiento", "Sin Tipificación", "Diferencia", 
    "CLASIFICADOR", "Fecha corte ultima conciliacion contable", "Fecha de cierre de conciliacion", "Responsable"
]

PATH_CONCILIADORES = "mapeo_conciliadores.csv"

# --- FUNCIONES DE PROCESAMIENTO ---
def transformar_amaho_logic(ref, clasificador):
    ref = str(ref).strip()
    if not ref or ref.lower() == 'nan' or ref == '' or str(clasificador).upper() == "PAGOS NO APLICADOS": return ref
    if "-" in ref:
        partes = ref.split("-")
        p, s = partes[0], partes[1] if len(partes) > 1 else ""
        c = 10 - (len(p) + len(s))
        return f"{p}{'0' * c}{s}" if c > 0 else (p + s)[:10]
    return ref.replace(" ", "").zfill(10)[:10]

def inicializar_base_datos(archivo_subido=None):
    if archivo_subido is not None:
        try:
            df_new = pd.read_csv(archivo_subido, encoding='latin1')
            df_new.columns = [str(c).strip() for c in df_new.columns]
            cols = [c for c in ["Nº ppal", "Conciliador", "Nombre principal"] if c in df_new.columns]
            df_save = df_new[cols].dropna(subset=["Nº ppal"])
            df_save.to_csv(PATH_CONCILIADORES, index=False, encoding='utf-8-sig')
            return df_save
        except: return pd.DataFrame()
    return pd.read_csv(PATH_CONCILIADORES) if os.path.exists(PATH_CONCILIADORES) else pd.DataFrame()

@st.cache_data(show_spinner=False)
def procesar_toda_la_cartera(uploaded_file, df_db):
    engine = 'pyxlsb' if uploaded_file.name.endswith('.xlsb') else None
    df_detalle = pd.read_excel(uploaded_file, sheet_name='Detalle', engine=engine)
    df_detalle.columns = [str(c).strip() for c in df_detalle.columns]
    
    # Filtros iniciales
    df_proc = df_detalle[(df_detalle['Cía'].isin([23, 25])) & (df_detalle['CLASIFICADOR'] != 'CARTERA CASTIGADA')].copy()
    
    # Asegurar limpieza de importes
    df_proc['Importe pendiente'] = pd.to_numeric(df_proc['Importe pendiente'], errors='coerce').fillna(0.0)
    
    # Tratamiento de fechas
    if pd.api.types.is_numeric_dtype(df_proc['Fecha factura']):
        df_proc['Fecha factura'] = pd.to_datetime(df_proc['Fecha factura'], unit='D', origin='1899-12-30')
    else:
        df_proc['Fecha factura'] = pd.to_datetime(df_proc['Fecha factura'], errors='coerce')
    
    # Corrección de Referencias vacías
    mask_guion = df_proc['Referencia'].astype(str).str.strip().isin(["-", "", "nan", "None"])
    df_proc.loc[mask_guion, 'Referencia'] = df_proc.loc[mask_guion, 'Número documento']
    
    # --- LÓGICA DE AGRUPACIÓN POR REFERENCIA Y Nº PPAL ---
    # Ordenar para que el registro con mayor importe mande en campos descriptivos (CLASIFICADOR)
    df_proc = df_proc.sort_values(by=['Nº ppal', 'Referencia', 'Importe pendiente'], ascending=[True, True, False])
    
    # Definimos como agrupar
    agg_dict = {
        'Nombre principal': 'first',
        'Fecha factura': 'max',
        'Importe pendiente': 'sum',
        'CLASIFICADOR': 'first'
    }
    
    # Columnas de tipificación (Sumar si existen)
    cols_tipif = [c for c in ESTRUCTURA_FINAL[9:23] if c in df_proc.columns]
    for c in cols_tipif:
        df_proc[c] = pd.to_numeric(df_proc[c], errors='coerce').fillna(0.0)
        agg_dict[c] = 'sum'
    
    # AGRUPACIÓN CLAVE: Se agrupa por Cliente y Referencia para evitar mezclar saldos
    df_grouped = df_proc.groupby(['Nº ppal', 'Referencia'], as_index=False).agg(agg_dict)
    
    # Re-aplicar lógicas sobre el set agrupado
    df_grouped['Amaho'] = df_grouped.apply(lambda r: transformar_amaho_logic(r['Referencia'], r['CLASIFICADOR']), axis=1)
    df_grouped['Fac_open'] = df_grouped.apply(lambda r: r['Amaho'] if str(r['Referencia']).strip().upper().startswith("EM") else str(r['Referencia']).replace("-", ""), axis=1)

    # Asegurar todas las columnas de tipificación (llenar con 0 si no existen)
    for col in ESTRUCTURA_FINAL[9:23]:
        if col not in df_grouped.columns:
            df_grouped[col] = 0.0

    # Lógica de clasificadores especiales
    df_grouped.loc[df_grouped['CLASIFICADOR'].str.upper().str.strip() == "PAGOS NO APLICADOS", 'Pagos No Aplicados'] = df_grouped['Importe pendiente']
    df_grouped.loc[df_grouped['CLASIFICADOR'].str.upper().str.strip() == "ALISTAMIENTO", 'Alistamiento'] = df_grouped['Importe pendiente']

    # Cruce con Conciliadores
    if not df_db.empty:
        df_db['key'] = df_db['Nº ppal'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_grouped['key'] = df_grouped['Nº ppal'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_grouped = df_grouped.merge(df_db[['key', 'Conciliador']], on='key', how='left')
        df_grouped['Responsable'] = df_grouped['Conciliador'].fillna("SIN ASIGNAR")
    else:
        df_grouped['Responsable'] = "SIN BASE"

    # Fechas finales
    MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
    df_grouped['Año Fact'] = df_grouped['Fecha factura'].dt.year.fillna(0).astype(int)
    df_grouped['Mes Fact'] = df_grouped['Fecha factura'].dt.month.map(MESES_ES).fillna("")
    df_grouped['Diferencia'] = 0.0
    for c in ["Fecha corte ultima conciliacion contable", "Fecha de cierre de conciliacion"]: df_grouped[c] = ""
    
    return df_grouped[ESTRUCTURA_FINAL].copy()

# --- INTERFAZ ---
modo = st.sidebar.radio("Módulo:", ["📥 Procesamiento", "📊 Dashboard de Auditoría"])

if modo == "📥 Procesamiento":
    st.title("🏦 Procesador de Fotos Cartera")
    with st.sidebar:
        up_db = st.file_uploader("Actualizar Conciliadores", type=["csv"])
        if up_db: inicializar_base_datos(up_db)
        valor_resumen = st.number_input("Valor Cartera (Resumen):", value=0.0)

    uploaded_cartera = st.file_uploader("Cargar Cartera Mensual", type=["xlsx", "xlsb"])

    if uploaded_cartera:
        df_db = inicializar_base_datos()
        df_final = procesar_toda_la_cartera(uploaded_cartera, df_db)
        
        suma_imp = df_final['Importe pendiente'].sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("Suma Importe (Agrupado)", f"${suma_imp:,.0f}")
        c2.metric("Valor Esperado", f"${valor_resumen:,.0f}")
        
        if round(abs(suma_imp - valor_resumen), 2) <= 1: c3.success("✅ Coincide")
        else: c3.error(f"❌ Diferencia: ${suma_imp - valor_resumen:,.0f}")

        if st.button("Generar Pack de Descarga"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                # 1. Reporte General
                det_out = io.BytesIO()
                with pd.ExcelWriter(det_out, engine='xlsxwriter', datetime_format='yyyy/mm/dd') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Detalle_Procesado')
                    workbook = writer.book
                    worksheet = writer.sheets['Detalle_Procesado']
                    fmt_money = workbook.add_format({'num_format': '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'})
                    for ci in range(8, 24): worksheet.set_column(ci, ci, 18, fmt_money)

                zip_file.writestr("00_Detalle_Procesado_General.xlsx", det_out.getvalue())
                
                # 2. Reportes por Responsable
                for resp in df_final['Responsable'].unique():
                    df_r = df_final[df_final['Responsable'] == resp].copy()
                    r_out = io.BytesIO()
                    with pd.ExcelWriter(r_out, engine='xlsxwriter', datetime_format='yyyy/mm/dd') as writer:
                        df_r.to_excel(writer, index=False, sheet_name='Consolidado')
                        wb, ws = writer.book, writer.sheets['Consolidado']
                        
                        # Formato Contabilidad
                        fmt_contab = wb.add_format({'num_format': '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'})
                        
                        # Inyectar fórmula de diferencia en columna X (índice 23)
                        for i in range(len(df_r)):
                            # J=10, W=23 (Letras de Excel), I=9
                            ws.write_formula(i+1, 23, f"=SUM(J{i+2}:W{i+2})-I{i+2}", fmt_contab)
                        
                        # Aplicar formato contabilidad a columnas de dinero (I hasta X)
                        for ci in range(8, 24):
                            ws.set_column(ci, ci, 20, fmt_contab)
                    
                    zip_file.writestr(f"Cartera_{str(resp).replace(' ', '_')}.xlsx", r_out.getvalue())
            
            st.download_button("📥 Descargar ZIP", data=zip_buffer.getvalue(), file_name="Reportes_Cartera_Agrupada.zip")

elif modo == "📊 Dashboard de Auditoría":
    # MÓDULO SIN MODIFICACIONES SEGÚN SOLICITUD
    st.title("📊 Dashboard de Control y Tipificación")
    
    col_path, col_btn = st.columns([4, 1])
    with col_path:
        ruta = st.text_input("Ruta de carpeta con archivos procesados:", placeholder="Ej: C:/Proyectos/Cartera/Salida")
    
    if 'df_master' not in st.session_state:
        st.session_state.df_master = None

    with col_btn:
        st.write("") 
        if st.button("🔄 Actualizar Datos"):
            st.session_state.df_master = None 

    if ruta and os.path.exists(ruta):
        if st.session_state.df_master is None:
            archivos = [os.path.join(ruta, f) for f in os.listdir(ruta) if f.endswith('.xlsx') and not f.startswith('~')]
            if archivos:
                with st.spinner("Cargando archivos..."):
                    lista_dfs = []
                    for a in archivos:
                        temp = pd.read_excel(a)
                        for c in CONCEPTOS_FILAS:
                            if c not in temp.columns: temp[c] = 0.0
                        lista_dfs.append(temp)
                    st.session_state.df_master = pd.concat(lista_dfs, ignore_index=True)
                    st.success("¡Datos cargados con éxito!")
            else:
                st.warning("No hay archivos Excel en la ruta.")

        if st.session_state.df_master is not None:
            df_full = st.session_state.df_master.copy()
            
            with st.expander("🔍 Filtros de Visualización", expanded=True):
                c1, c2, c3 = st.columns(3)
                f_resp = c1.multiselect("Responsable:", options=sorted(df_full['Responsable'].unique().astype(str).tolist()))
                f_npal = c2.multiselect("Nº ppal:", options=sorted(df_full['Nº ppal'].unique().astype(str).tolist()))
                f_nom = c3.multiselect("Nombre principal:", options=sorted(df_full['Nombre principal'].unique().astype(str).tolist()))
            
            df_f = df_full.copy()
            if f_resp: df_f = df_f[df_f['Responsable'].isin(f_resp)]
            if f_npal: df_f = df_f[df_f['Nº ppal'].astype(str).isin(f_npal)]
            if f_nom: df_f = df_f[df_f['Nombre principal'].isin(f_nom)]

            st.subheader("📑 Matriz de Tipificación por Año")
            matriz = df_f.groupby('Año Fact')[CONCEPTOS_FILAS].sum().T
            matriz['Total General'] = matriz.sum(axis=1)
            
            if 'Importe pendiente' in matriz.index:
                imp_pend_val = matriz.loc[['Importe pendiente']]
                conceptos_restantes = matriz.drop('Importe pendiente').sort_values('Total General', ascending=False)
                matriz_final = pd.concat([conceptos_restantes, imp_pend_val])
                matriz_final = matriz_final[matriz_final['Total General'] != 0]
                st.dataframe(matriz_final.style.format("${:,.0f}"), use_container_width=True)

            col_a, col_b = st.columns(2)
            
            with col_a:
                st.subheader("⚠️ Conciliadores con Diferencia ≠ 0")
                cols_tip = [c for c in CONCEPTOS_FILAS if c != "Importe pendiente"]
                df_f['Diferencia_Check'] = df_f[cols_tip].sum(axis=1) - df_f['Importe pendiente']
                dif_df = df_f[df_f['Diferencia_Check'].round(2) != 0].groupby(['Responsable', 'Nombre principal'])['Diferencia_Check'].sum().reset_index()
                st.dataframe(dif_df.style.format({"Diferencia_Check": "${:,.0f}"}), use_container_width=True)

            with col_b:
                st.subheader("🔝 Top 25 Entidades Mayor Saldo")
                top_25 = df_f.groupby('Nombre principal')['Importe pendiente'].sum().sort_values(ascending=False).head(25).reset_index()
                st.dataframe(top_25.style.format({"Importe pendiente": "${:,.0f}"}), use_container_width=True)

            csv = matriz_final.to_csv().encode('utf-8-sig')
            st.download_button("📥 Descargar Matriz (CSV)", data=csv, file_name="Matriz_Dashboard.csv")
    else:
        st.info("Ingrese una ruta válida para cargar los datos en el Dashboard.")