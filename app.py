import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="App Jurídica - Control de Juicios", layout="wide", page_icon="⚖️")

st.title("⚖️ Sistema de Gestión y Control de Juicios Jurídicos")
st.write("Consulta el estado procesal, montos y detalles de juicios por Nombre, C.I. o N° de Socio.")

# --- CARGA DE DATOS ---
st.sidebar.header("📁 Carga de Datos")
archivo_excel = st.sidebar.file_uploader("Subir planilla de Excel con Juicios", type=["xlsx", "xls"])

if archivo_excel is not None:
    try:
        df = pd.read_excel(archivo_excel)
        # Limpiar nombres de columnas
        df.columns = [str(c).strip() for c in df.columns]
        
        st.sidebar.success("✅ Planilla cargada correctamente")
        
        # --- FILTROS DE BÚSQUEDA ---
        st.subheader("🔍 Buscador de Juicios")
        busqueda = st.text_input("Ingresá Nombre, Apellido, C.I. o N° de Socio:", placeholder="Ej: 1234567 o Juan Pérez")

        if busqueda:
            # Convertir todas las columnas a string para realizar la búsqueda flexible
            mask = df.apply(lambda row: row.astype(str).str.contains(busqueda, case=False, na=False)).any(axis=1)
            df_filtrado = df[mask]
        else:
            df_filtrado = df

        st.markdown("---")

        # --- VISTA GENERAL / MÓVIL ---
        st.subheader(f"📋 Resultados Encontrados ({len(df_filtrado)})")
        
        if not df_filtrado.empty:
            for idx, row in df_filtrado.iterrows():
                with st.expander(f"⚖️ Caso: {row.get('Socio / Demandado', row.get('Nombre', 'Detalle del Juicio'))} | C.I.: {row.get('C.I.', 'N/A')} | Estado: {row.get('Estado Procesal', 'N/A')}"):
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown("**📌 DATOS DEL SOCIO**")
                        st.write(f"**Nombre y Apellido:** {row.get('Socio / Demandado', row.get('Nombre', '-'))}")
                        st.write(f"**C.I. N°:** {row.get('C.I.', '-')}")
                        st.write(f"**N° de Socio:** {row.get('N° Socio', '-')}")
                        st.write(f"**Abogado a Cargo:** {row.get('Abogado a Cargo', '-')}")

                    with col2:
                        st.markdown("**🏛️ DATOS JUDICIALES**")
                        st.write(f"**Juzgado:** {row.get('Juzgado', '-')}")
                        st.write(f"**Secretaría:** {row.get('Secretaria', '-')}")
                        st.write(f"**N° Expediente:** {row.get('N° Expte', row.get('Expediente', '-'))}")
                        st.write(f"**Año:** {row.get('Año', '-')}")
                        st.write(f"**Estado Procesal:** `{row.get('Estado Procesal', '-')}`")

                    with col3:
                        st.markdown("**💰 ESTADO FINANCIERO**")
                        st.write(f"**Monto Demandado:** {row.get('Monto Demandado', '-')}")
                        st.write(f"**Monto Cobrado:** {row.get('Monto Cobrado', '-')}")
                        st.write(f"**Saldo por Cobrar:** {row.get('Saldo por Cobrar', '-')}")
                        st.info(f"**Próxima Acción / Observaciones:** {row.get('Observaciones', '-')}")

            st.markdown("---")
            st.subheader("📊 Tabla General para Impresión / Exportación")
            st.dataframe(df_filtrado, use_container_width=True)

            # --- BOTÓN DE DESCARGA / INFORME PARA IMPRIMIR ---
            fecha_str = datetime.now().strftime("%d_%m_%Y")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, sheet_name='Informe Juicios', index=False)

            st.download_button(
                label="🖨️ Descargar Informe General para Imprimir (Excel)",
                data=buffer.getvalue(),
                file_name=f"Informe_General_Juicios_{fecha_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("No se encontraron registros que coincidan con la búsqueda.")

    except Exception as e:
        st.error(f"Error al procesar el archivo de Excel: {e}")
else:
    st.info("👈 Por favor, sube la planilla de Excel desde la barra lateral para empezar a buscar.")
