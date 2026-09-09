import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

st.set_page_config(page_title="App Jurídica - Control de Juicios", layout="wide", page_icon="⚖️")

DB_FILE = "juicios_db.csv"

# --- FUNCIONES DE BASE DE DATOS LOCAL ---
def cargar_datos():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE, dtype=str)
    return pd.DataFrame(columns=[
        'Socio / Demandado', 'C.I.', 'N° Socio', 'Juzgado', 'Secretaria', 
        'N° Expte', 'Año', 'Abogado a Cargo', 'Estado Procesal', 
        'Monto Demandado', 'Monto Cobrado', 'Saldo por Cobrar', 'Observaciones'
    ])

def guardar_datos(df):
    df.to_csv(DB_FILE, index=False)

df_db = cargar_datos()

st.title("⚖️ Sistema de Gestión y Control de Juicios Jurídicos")

# --- BARRA LATERAL: MANTENIMIENTO ---
st.sidebar.header("⚙️ Gestión de Datos")

opcion = st.sidebar.radio("Menú:", ["🔍 Consultar Juicios", "➕ Cargar Nuevo Juicio", "📥 Carga Masiva (Excel)"])

# --- MÓDULO 1: CONSULTA Y BÚSQUEDA ---
if opcion == "🔍 Consultar Juicios":
    st.subheader("🔍 Buscador General de Juicios")
    busqueda = st.text_input("Ingresa Nombre, Apellido, C.I. o N° de Socio:", placeholder="Ej: 1234567 o Juan Pérez")

    if not df_db.empty:
        if busqueda:
            mask = df_db.apply(lambda row: row.astype(str).str.contains(busqueda, case=False, na=False)).any(axis=1)
            df_filtrado = df_db[mask]
        else:
            df_filtrado = df_db

        st.markdown("---")
        st.subheader(f"📋 Registros Encontrados ({len(df_filtrado)})")
        
        for idx, row in df_filtrado.iterrows():
            with st.expander(f"⚖️ {row.get('Socio / Demandado', 'S/N')} | C.I.: {row.get('C.I.', 'N/A')} | Estado: {row.get('Estado Procesal', 'N/A')}"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("**📌 DATOS DEL SOCIO**")
                    st.write(f"**Nombre:** {row.get('Socio / Demandado', '-')}")
                    st.write(f"**C.I.:** {row.get('C.I.', '-')}")
                    st.write(f"**N° Socio:** {row.get('N° Socio', '-')}")
                    st.write(f"**Abogado:** {row.get('Abogado a Cargo', '-')}")

                with col2:
                    st.markdown("**🏛️ DATOS JUDICIALES**")
                    st.write(f"**Juzgado:** {row.get('Juzgado', '-')}")
                    st.write(f"**Secretaría:** {row.get('Secretaria', '-')}")
                    st.write(f"**Expediente:** {row.get('N° Expte', '-')}/{row.get('Año', '-')}")
                    st.write(f"**Estado:** `{row.get('Estado Procesal', '-')}`")

                with col3:
                    st.markdown("**💰 FINANZAS Y NOTAS**")
                    st.write(f"**Demandado:** {row.get('Monto Demandado', '-')}")
                    st.write(f"**Cobrado:** {row.get('Monto Cobrado', '-')}")
                    st.write(f"**Saldo:** {row.get('Saldo por Cobrar', '-')}")
                    st.info(f"**Obs:** {row.get('Observaciones', '-')}")

        st.markdown("---")
        st.dataframe(df_filtrado, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, sheet_name='Juicios', index=False)

        st.download_button(
            label="🖨️ Descargar Informe para Imprimir",
            data=buffer.getvalue(),
            file_name=f"Informe_Juicios_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("La base de datos está vacía. Carga un archivo Excel desde el menú lateral para importar los registros.")

# --- MÓDULO 2: CARGAR NUEVO JUICIO MANUALMENTE ---
elif opcion == "➕ Cargar Nuevo Juicio":
    st.subheader("➕ Registro Individual de Nuevo Juicio")
    with st.form("form_nuevo_juicio", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nombre = st.text_input("Nombre y Apellido del Socio / Demandado *")
            ci = st.text_input("Número de C.I. *")
            socio = st.text_input("N° de Socio")
            juzgado = st.text_input("Juzgado")
            secretaria = st.text_input("Secretaría")
            expte = st.text_input("N° de Expediente")
            ano = st.text_input("Año")
        with col2:
            abogado = st.text_input("Abogado a Cargo")
            estado = st.text_input("Estado Procesal")
            monto_dem = st.text_input("Monto Demandado")
            monto_cob = st.text_input("Monto Cobrado")
            saldo = st.text_input("Saldo por Cobrar")
            obs = st.text_area("Próxima Acción / Observaciones")

        submitted = st.form_submit_button("💾 Guardar Juicio en el Sistema")
        if submitted:
            if not nombre or not ci:
                st.error("Los campos Nombre y C.I. son obligatorios.")
            else:
                nuevo_registro = pd.DataFrame([{
                    'Socio / Demandado': nombre, 'C.I.': ci, 'N° Socio': socio,
                    'Juzgado': juzgado, 'Secretaria': secretaria, 'N° Expte': expte,
                    'Año': ano, 'Abogado a Cargo': abogado, 'Estado Procesal': estado,
                    'Monto Demandado': monto_dem, 'Monto Cobrado': monto_cob,
                    'Saldo por Cobrar': saldo, 'Observaciones': obs
                }])
                df_db = pd.concat([df_db, nuevo_registro], ignore_index=True)
                guardar_datos(df_db)
                st.success(f"✅ Juicio para '{nombre}' registrado con éxito.")

# --- MÓDULO 3: CARGA MASIVA INICIAL ---
elif opcion == "📥 Carga Masiva (Excel)":
    st.subheader("📥 Cargar/Reemplazar datos desde Excel")
    archivo = st.file_uploader("Selecciona la planilla de Excel inicial", type=["xlsx", "xls"])
    if archivo:
        if st.button("⚠️ Importar datos e inicializar Sistema"):
            try:
                df_excel = pd.read_excel(archivo, dtype=str)
                df_excel.columns = [str(c).strip() for c in df_excel.columns]
                guardar_datos(df_excel)
                st.success("✅ Base de datos importada correctamente.")
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
