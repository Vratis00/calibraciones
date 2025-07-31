import streamlit as st
import sqlite3
import os
from datetime import date
import pandas as pd
import io
from fpdf import FPDF

archivo_db = 'calibraciones.db'

# Inicializar Base de Datos
def inicializar_db():
    conn = sqlite3.connect(archivo_db)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS calibraciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    orden TEXT,
                    fecha TEXT,
                    certificado TEXT,
                    equipo TEXT,
                    cliente TEXT,
                    sede TEXT,
                    conformidad TEXT,
                    ejecutado TEXT,
                    firmado TEXT,
                    avalado TEXT,
                    tipo TEXT
                )''')
    conn.commit()
    conn.close()

# Determinar letra del equipo
def determinar_letra_equipo(equipo):
    equipos_E = ["Tensiometro Digital", "Tensiometro Adulto", "Tensiometro Pediátrico"]
    equipos_B = ["Balanza Adulto", "Pesa Bebé", "Gramera", "Dinamómetro"]
    if equipo in equipos_E:
        return 'E'
    elif equipo in equipos_B:
        return 'B'
    else:
        return 'X'

# Inicialización
inicializar_db()

# Interfaz de usuario en Streamlit
st.title("Asignación de Stickers de Calibración (SQLite)")

# Casillas para ingresar los consecutivos iniciales manualmente
consecutivo_E_manual = st.number_input("Número de Consecutivo Inicial TENSIÓMETROS", min_value=1, step=1, value=1899)
consecutivo_B_manual = st.number_input("Número de Consecutivo Inicial BALANZAS", min_value=1, step=1, value=871)

modo_asignacion = st.checkbox("Modo Asignación de Stickers (simplificado)")

orden = st.text_input("N° Orden Servicio")
fecha = st.date_input("Fecha", value=date.today())
cliente = st.text_input("Cliente")

if not modo_asignacion:
    sede = st.text_input("Sede o Servicio")
    conformidad = st.selectbox("Conformidad", ["Pasa", "No pasa"])
    ejecutado = st.selectbox("Ejecutado por", ["Edwin Garzon", "Juan Melo", "Vratislav Cala"])

st.markdown("---")

st.subheader("Agregar Equipos")
equipo = st.selectbox("Tipo de Equipo", [
    "Tensiometro Digital",
    "Tensiometro Adulto",
    "Tensiometro Pediátrico",
    "Balanza Adulto",
    "Pesa Bebé",
    "Gramera",
    "Dinamómetro"
])
cantidad = st.number_input("Cantidad", min_value=1, step=1)

tipo_equipo = st.selectbox("Tipo de Certificado", ["Esfigmomanómetro", "Balanza"])

if 'equipos' not in st.session_state:
    st.session_state.equipos = []

if st.button("Agregar Equipo"):
    st.session_state.equipos.append((equipo, cantidad, tipo_equipo))

st.write("### Equipos añadidos:")
for eq, cant, tipo in st.session_state.equipos:
    st.write(f"{cant} x {eq} ({tipo})")

if st.button("Guardar Datos en Base de Datos"):
    if not st.session_state.equipos:
        st.error("Debes agregar al menos un equipo")
    elif not orden or not cliente:
        st.error("Orden de Servicio y Cliente son obligatorios")
    else:
        conn = sqlite3.connect(archivo_db)
        c = conn.cursor()

        certificados_asignados = []
        consecutivos = {'E': consecutivo_E_manual, 'B': consecutivo_B_manual}

        for equipo, cantidad, tipo in st.session_state.equipos:
            letra = determinar_letra_equipo(equipo)
            if letra == 'X':
                st.error(f"Equipo inválido: {equipo}")
                st.stop()
            for i in range(cantidad):
                cert_num = f"{letra}-25-{consecutivos[letra]}"
                c.execute("""
                    INSERT INTO calibraciones (orden, fecha, certificado, equipo, cliente, sede, conformidad, ejecutado, firmado, avalado, tipo)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    orden,
                    fecha.strftime('%Y-%m-%d'),
                    cert_num,
                    equipo,
                    cliente,
                    sede if not modo_asignacion else '',
                    conformidad if not modo_asignacion else '',
                    ejecutado if not modo_asignacion else '',
                    "Angie Rodriguez" if not modo_asignacion else '',
                    "Vratislav Cala" if not modo_asignacion else '',
                    tipo
                ))
                certificados_asignados.append(cert_num)
                consecutivos[letra] += 1

        conn.commit()
        conn.close()
        st.success(f"Certificados asignados: {', '.join(certificados_asignados)}")
        st.session_state.equipos = []

# Cargar datos desde un archivo Excel
datos_excel = st.file_uploader("Subir Excel para cargar a la base de datos", type=['xlsx'])

if datos_excel is not None:
    tipo_seleccionado = st.selectbox("Tipo de datos a importar", ["Balanza", "Esfigmomanómetro"])

    if st.button("Cargar Datos desde Excel"):
        df = pd.read_excel(datos_excel, index_col=None)

        st.write("Encabezados leídos:", df.columns.tolist())

        df.columns = df.columns.str.strip().str.upper()

        columnas_alias = {}
        for col in df.columns:
            if 'ORDEN' in col:
                columnas_alias[col] = 'orden'
            elif 'FECHA' in col:
                columnas_alias[col] = 'fecha'
            elif 'CERTIFICADO' in col:
                columnas_alias[col] = 'certificado'
            elif 'EQUIPO' in col:
                columnas_alias[col] = 'equipo'
            elif 'CLIENTE' in col:
                columnas_alias[col] = 'cliente'
            elif 'SEDE' in col:
                columnas_alias[col] = 'sede'
            elif 'CONFORMIDAD' in col:
                columnas_alias[col] = 'conformidad'
            elif 'EJECUTADO' in col:
                columnas_alias[col] = 'ejecutado'
            elif 'FIRMADO' in col:
                columnas_alias[col] = 'firmado'
            elif 'AVALADO' in col:
                columnas_alias[col] = 'avalado'

        df = df.rename(columns=columnas_alias)
        df['tipo'] = tipo_seleccionado

        df['orden'] = df['orden'].astype(str).str.strip()
        df['fecha'] = df['fecha'].astype(str).str.strip()
        df['certificado'] = df['certificado'].astype(str).str.strip()

        conn = sqlite3.connect(archivo_db)
        c = conn.cursor()

        for _, row in df.iterrows():
            c.execute("""
                INSERT INTO calibraciones (orden, fecha, certificado, equipo, cliente, sede, conformidad, ejecutado, firmado, avalado, tipo)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row.get('orden', ''),
                row.get('fecha', ''),
                row.get('certificado', ''),
                row.get('equipo', ''),
                row.get('cliente', ''),
                row.get('sede', ''),
                row.get('conformidad', ''),
                row.get('ejecutado', ''),
                row.get('firmado', ''),
                row.get('avalado', ''),
                row.get('tipo', '')
            ))
        conn.commit()
        conn.close()
        st.success("Datos importados correctamente desde el Excel")

if st.button("🗑️ Borrar Todo"):
    conn = sqlite3.connect(archivo_db)
    c = conn.cursor()
    c.execute("DELETE FROM calibraciones")
    conn.commit()
    conn.close()
    st.session_state.equipos = []
    st.success("Datos eliminados y base de datos reiniciada.")

if os.path.exists(archivo_db):
    with open(archivo_db, "rb") as file:
        st.download_button(
            label="📥 Descargar Base de Datos (.db)",
            data=file,
            file_name=archivo_db,
            mime="application/x-sqlite3"
        )

# VISUALIZAR BASE DE DATOS COMPLETA
st.markdown("---")
st.subheader("📋 Visualizar Base de Datos Completa")
if st.button("Mostrar Todos los Registros"):
    conn = sqlite3.connect(archivo_db)
    df_all = pd.read_sql_query("SELECT id AS ITEM, orden AS 'N° ORDEN DE SERVICIO', fecha AS FECHA, certificado AS 'N° CERTIFICADO', equipo AS EQUIPO, cliente AS CLIENTE, sede AS 'SEDE O SERVICIO', conformidad AS CONFORMIDAD, ejecutado AS 'EJECUTADO POR', firmado AS 'FIRMADO POR', avalado AS 'AVALADO POR', tipo AS TIPO FROM calibraciones", conn)
    conn.close()
    st.dataframe(df_all)

# FILTRAR Y VISUALIZAR REGISTROS
st.markdown("---")
st.subheader("📊 Filtro de Base de Datos")

with st.expander("Filtrar Registros"):
    cliente_filtro = st.text_input("Filtrar por Cliente")
    tipo_filtro = st.selectbox("Filtrar por Tipo", ["Todos", "Balanza", "Esfigmomanómetro"])
    fecha_filtro = st.date_input("Filtrar por Fecha (opcional)", value=None)

if st.button("Aplicar Filtro y Mostrar"):
    conn = sqlite3.connect(archivo_db)
    query = "SELECT id AS ITEM, orden AS 'N° ORDEN DE SERVICIO', fecha AS FECHA, certificado AS 'N° CERTIFICADO', equipo AS EQUIPO, cliente AS CLIENTE, sede AS 'SEDE O SERVICIO', conformidad AS CONFORMIDAD, ejecutado AS 'EJECUTADO POR', firmado AS 'FIRMADO POR', avalado AS 'AVALADO POR', tipo AS TIPO FROM calibraciones WHERE 1=1"

    if cliente_filtro:
        query += f" AND cliente LIKE '%{cliente_filtro}%'"
    if tipo_filtro != "Todos":
        query += f" AND tipo = '{tipo_filtro}'"
    if fecha_filtro:
        query += f" AND fecha = '{fecha_filtro}'"

    df = pd.read_sql_query(query, conn)
    conn.close()

    st.dataframe(df)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Filtrado')
        writer.save()
        st.download_button(
            label="📥 Descargar Excel Filtrado",
            data=buffer,
            file_name="reporte_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Exportar PDF Simple
    if st.button("📄 Exportar PDF Resumen"):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Reporte de Calibraciones", ln=True, align='C')

        for index, row in df.iterrows():
            pdf.cell(0, 10, txt=f"Certificado: {row['N° CERTIFICADO']} - Equipo: {row['EQUIPO']} - Cliente: {row['CLIENTE']}", ln=True)

        pdf_buffer = io.BytesIO()
        pdf.output(pdf_buffer)
        st.download_button(
            label="📥 Descargar PDF Resumen",
            data=pdf_buffer.getvalue(),
            file_name="reporte.pdf",
            mime="application/pdf"
        )
