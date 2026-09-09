import streamlit as st
import pandas as pd
import os
import io
import re
from datetime import datetime

st.set_page_config(page_title="App Jurídica - Control de Juicios", layout="wide", page_icon="⚖️")

DB_FILE = "juicios_db.csv"

COLUMNAS_ESTANDAR = [
    'Socio / Demandado', 'C.I.', 'N° Socio', 'Juzgado', 'Secretaria', 
    'N° Expte', 'Año', 'Abogado a Cargo', 'Estado Procesal', 
    'Monto Demandado', 'Monto Cobrado', 'Saldo por Cobrar', 'Observaciones'
]

def limpiar_monto(val):
    """ Convierte texto numérico a float limpio para cálculos """
    if pd.isna(val) or val is None:
        return 0.0
    s = str(val).replace('.', '').replace(',', '.').strip()
    numeros = re.findall(r'[-+]?\d*\.\d+|\d+', s)
    if numeros:
        return float(numeros[0])
    return 0.0

def formato_guarani(val):
    """ Formatea números a estilo de miles paraguayo (ej: 1.500.000) """
    try:
        return f"{int(round(val)):,}".replace(',', '.')
    except:
        return "0"

def normalizar_dataframe(df):
    df.columns = [str(c).strip().upper() for c in df.columns]
    mapping = {
        'NOMBRE Y APELLIDO': 'Socio / Demandado', 'NOMBRE': 'Socio / Demandado', 'SOCIO / DEMANDADO': 'Socio / Demandado',
        'N° DE CI': 'C.I.', 'CI': 'C.I.', 'C.I.': 'C.I.', 'N° DE C.I.': 'C.I.',
        'N° SOCIO': 'N° Socio', 'SOCIO': 'N° Socio',
        'JUZGADO': 'Juzgado', 'SECRETARIA': 'Secretaria',
        'N° EXPTE': 'N° Expte', 'EXPEDIENTE': 'N° Expte', 'AÑO': 'Año',
        'ABOGADO A CARGO': 'Abogado a Cargo', 'ABOGADO': 'Abogado a Cargo',
        'ESTADO PROCESAL': 'Estado Procesal', 'ESTADO': 'Estado Procesal',
        'MONTO DEMANDADO': 'Monto Demandado', 'MONTO COBRADO': 'Monto Cobrado',
        'SALDO POR COBRAR': 'Saldo por Cobrar',
        'OBSERVACION': 'Observaciones', 'OBSERVACIONES': 'Observaciones'
    }
    df = df.rename(columns=mapping)
    for col in COLUMNAS_ESTANDAR:
        if col not in df.columns:
            df[col] = "-"
    return df[COLUMNAS_ESTANDAR].astype(str)

def cargar_datos():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE, dtype=str)
    return pd.DataFrame(columns=COLUMNAS_ESTANDAR)

def guardar_datos(df):
    df.to_csv(DB_FILE, index=False)

df_db = cargar_datos()

st.title("⚖️ Sistema de Gestión y Control de Juicios Jurídicos")

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Gestión de Datos")
opcion = st.sidebar.radio("Menú:", [
    "🔍 Consultar Juicios", 
    "💰 Cargar Pago / Actualizar Saldo", 
    "➕ Cargar Nuevo Juicio", 
    "📥 Carga Masiva (Excel/Word)"
])

# --- MÓDULO 1: CONSULTA Y BÚSQUEDA ---
if opcion == "🔍 Consultar Juicios":
    st.subheader("🔍 Buscador General de Juicios")
    busqueda = st.text_input("Ingresá Nombre, Apellido, C.I. o N° de Socio:", placeholder="Ej: 1340084 o Teodoro Medina")

    if not df_db.empty:
        if busqueda:
            mask = df_db.apply(lambda row: row.astype(str).str.contains(busqueda, case=False, na=False)).any(axis=1)
            df_filtrado = df_db[mask]
        else:
            df_filtrado = df_db

        st.markdown("---")
        st.subheader(f"📋 Registros Encontrados ({len(df_filtrado)})")
        
        for idx, row in df_filtrado.iterrows():
            nombre_socio = row.get('Socio / Demandado', 'S/N')
            ci_socio = row.get('C.I.', 'N/A')
            estado_socio = row.get('Estado Procesal', 'N/A')
            
            with st.expander(f"⚖️ {nombre_socio} | C.I.: {ci_socio} | Estado: {estado_socio}"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("**📌 DATOS DEL SOCIO**")
                    st.write(f"**Nombre:** {nombre_socio}")
                    st.write(f"**C.I.:** {ci_socio}")
                    st.write(f"**N° Socio:** {row.get('N° Socio', '-')}")
                    st.write(f"**Abogado:** {row.get('Abogado a Cargo', '-')}")

                with col2:
                    st.markdown("**🏛️ DATOS JUDICIALES**")
                    st.write(f"**Juzgado:** {row.get('Juzgado', '-')}")
                    st.write(f"**Secretaría:** {row.get('Secretaria', '-')}")
                    st.write(f"**Expediente:** {row.get('N° Expte', '-')}/{row.get('Año', '-')}")
                    st.write(f"**Estado:** `{estado_socio}`")

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

# --- MÓDULO 2: CARGAR PAGO / ACTUALIZAR SALDO ---
elif opcion == "💰 Cargar Pago / Actualizar Saldo":
    st.subheader("💰 Módulo de Carga de Pagos y Actualización de Saldo")
    
    if df_db.empty:
        st.info("La base de datos está vacía. Carga registros primero para procesar pagos.")
    else:
        col_criterio, col_busqueda = st.columns([1, 2])
        with col_criterio:
            criterio = st.selectbox("Criterio de Búsqueda:", ["Número de C.I.", "Número de Socio"])
        with col_busqueda:
            valor_busqueda = st.text_input(f"Ingresá el {criterio} y presioná Enter:", placeholder="Ej: 1340084")

        if valor_busqueda:
            col_target = 'C.I.' if criterio == "Número de C.I." else 'N° Socio'
            # Búsqueda exacta limpia
            coincidencias = df_db[df_db[col_target].astype(str).str.strip() == valor_busqueda.strip()]

            if coincidencias.empty:
                st.error(f"No se encontró ningún juicio registrado con el {criterio}: '{valor_busqueda}'")
            else:
                idx_socio = coincidencias.index[0]
                datos = df_db.loc[idx_socio]

                st.markdown("---")
                st.success(f"👤 **Socio:** {datos['Socio / Demandado']} | **C.I.:** {datos['C.I.']} | **N° Socio:** {datos['N° Socio']}")

                # Obtener valores numéricos
                saldo_actual_num = limpiar_monto(datos['Saldo por Cobrar'])
                cobrado_actual_num = limpiar_monto(datos['Monto Cobrado'])

                col_info1, col_info2, col_info3 = st.columns(3)
                with col_info1:
                    st.metric("Saldo Anterior (Deuda)", f"Gs. {formato_guarani(saldo_actual_num)}")
                with col_info2:
                    st.metric("Total Cobrado a la Fecha", f"Gs. {formato_guarani(cobrado_actual_num)}")
                with col_info3:
                    st.metric("Monto Demandado Inicial", f"Gs. {datos['Monto Demandado']}")

                st.markdown("---")
                with st.form("form_procesar_pago", clear_on_submit=False):
                    st.subheader("💳 Registrar Cobro")
                    
                    col_pago1, col_pago2 = st.columns(2)
                    with col_pago1:
                        monto_pago_input = st.number_input("Monto Abonado Hoy (Gs.) *", min_value=0.0, step=50000.0, format="%.0f")
                        fecha_pago = datetime.now().strftime("%d/%m/%Y")
                    with col_pago2:
                        nuevo_saldo_calculado = max(0.0, saldo_actual_num - monto_pago_input)
                        nuevo_cobrado_calculado = cobrado_actual_num + monto_pago_input
                        st.write(f"**Nuevo Saldo a la Fecha:** Gs. `{formato_guarani(nuevo_saldo_calculado)}`")
                        st.write(f"**Nuevo Total Cobrado:** Gs. `{formato_guarani(nuevo_cobrado_calculado)}`")

                    obs_adicional = st.text_input("Observación del pago:", placeholder=f"Ej: Pago abonado en fecha {fecha_pago}")

                    btn_guardar_pago = st.form_submit_button("💾 Procesar y Registrar Pago")

                    if btn_guardar_pago:
                        if monto_pago_input <= 0:
                            st.error("El monto abonado debe ser mayor a 0.")
                        else:
                            # Actualización en la base de datos
                            df_db.at[idx_socio, 'Saldo por Cobrar'] = formato_guarani(nuevo_saldo_calculado)
                            df_db.at[idx_socio, 'Monto Cobrado'] = formato_guarani(nuevo_cobrado_calculado)
                            
                            obs_anterior = str(df_db.at[idx_socio, 'Observaciones'])
                            nueva_nota = f"Pago Gs. {formato_guarani(monto_pago_input)} ({fecha_pago})"
                            if obs_adicional.strip():
                                nueva_nota += f" - {obs_adicional.strip()}"

                            if obs_anterior and obs_anterior != "-":
                                df_db.at[idx_socio, 'Observaciones'] = f"{obs_anterior} | {nueva_nota}"
                            else:
                                df_db.at[idx_socio, 'Observaciones'] = nueva_nota

                            guardar_datos(df_db)
                            st.balloons()
                            st.success(f"✅ ¡Pago de Gs. {formato_guarani(monto_pago_input)} registrado correctamente!")
                            st.info(f"Nuevo Saldo del Socio: Gs. {formato_guarani(nuevo_saldo_calculado)}")

# --- MÓDULO 3: CARGAR NUEVO JUICIO MANUALMENTE ---
elif opcion == "➕ Cargar Nuevo Juicio":
    st.subheader("➕ Registro Individual de Nuevo Juicio")
    with st.form("form_nuevo_juicio", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nombre = st.text_input("Nombre y Apellido del Socio / Demandado *")
            ci = st.text_input("Número de C.I. *")
            socio = st.text_input("N° de Socio")
            juzgado = st.text_input("Juzgado", value="Juzgado de Paz La Encarnación")
            secretaria = st.text_input("Secretaría", value="Secretaría N° 1")
            expte = st.text_input("N° de Expediente")
            ano = st.text_input("Año")
        with col2:
            abogado = st.text_input("Abogado a Cargo")
            estado = st.text_input("Estado Procesal")
            monto_dem = st.text_input("Monto Demandado")
            monto_cob = st.text_input("Monto Cobrado", value="0")
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
                    'Saldo por Cobrar': saldo if saldo else monto_dem, 'Observaciones': obs
                }])
                df_db = pd.concat([df_db, nuevo_registro], ignore_index=True)
                guardar_datos(df_db)
                st.success(f"✅ Juicio para '{nombre}' registrado con éxito.")

# --- MÓDULO 4: CARGA MASIVA INICIAL ---
elif opcion == "📥 Carga Masiva (Excel/Word)":
    st.subheader("📥 Cargar/Reemplazar datos desde Excel")
    st.write("Si copiaste la tabla de Word a Excel, podés subir la planilla directamente acá.")
    archivo = st.file_uploader("Selecciona la planilla de Excel", type=["xlsx", "xls"])
    if archivo:
        if st.button("⚠️ Importar datos e inicializar Sistema"):
            try:
                df_excel = pd.read_excel(archivo, dtype=str)
                df_normalizado = normalizar_dataframe(df_excel)
                guardar_datos(df_normalizado)
                st.success("✅ Base de datos importada correctamente y campos alineados.")
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
