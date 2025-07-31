import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
import io

# -----------------------------------------
# INICIALIZAR BASE DE DATOS (si no existe)
# -----------------------------------------
def inicializar_db():
    conn = sqlite3.connect('calibraciones.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS calibraciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            orden TEXT,
            fecha TEXT,
            equipo TEXT,
            certificado TEXT,
            cliente TEXT,
            sede TEXT,
            conformidad TEXT,
            ejecutado TEXT,
            firmado TEXT,
            avalado TEXT
        )
    ''')
    conn.commit()
    conn.close()

# -----------------------------------------
# DETERMINAR LETRA DEL EQUIPO
# -----------------------------------------
def determinar_letra_equipo(equipo):
    equipos_E = ["Tensiometro Digital", "Tensiometro Adulto", "Tensiometro Pediátrico"]
    equipos_B = ["Balanza Adulto", "Pesa Bebé", "Gramera", "Dinamómetro"]
    if equipo in equipos_E:
        return 'E'
    elif equipo in equipos_B:
        return 'B'
    else:
        return 'X'

# -----------------------------------------
# OBTENER CONSECUTIVO ACTUAL DESDE LA BD
# -----------------------------------------
def obtener_consecutivo_actual(letra, consecutivo_manual):
    conn = sqlite3.connect('calibraciones.db')
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(CAST(substr(certificado, 7) AS INTEGER)) FROM calibraciones WHERE certificado LIKE ? ", (f"{letra}-25-%",))
    resultado = cursor.fetchone()[0]
    conn.close()
    return resultado + 1 if resultado else consecutivo_manual

# -----------------------------------------
# GUARDAR EQUIPOS EN LA BASE DE DATOS
# -----------------------------------------
def guardar_calibraciones(orden, fecha, cliente, sede, conformidad, ejecutado, equipos, modo_asignacion, consecutivo_E_manual, consecutivo_B_manual):
    conn = sqlite3.connect('calibraciones.db')
    cursor = conn.cursor()
    consecutivos = {
        'E': obtener_consecutivo_actual('E', consecutivo_E_manual),
        'B': obtener_consecutivo_actual('B', consecutivo_B_manual)
    }

    certificados_asignados = []

    for equipo, cantidad in equipos:
        letra = determinar_letra_equipo(equipo)
        if letra == 'X':
            st.error(f"Equipo inválido: {equipo}")
            return
        for i in range(cantidad):
            cert_num = f"{letra}-25-{consecutivos[letra]}"
            cursor.execute('''
                INSERT INTO calibraciones (orden, fecha, equipo, certificado, cliente, sede, conformidad, ejecutado, firmado, avalado)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                orden,
                fecha.strftime('%Y-%m-%d'),
                equipo,
                cert_num,
                cliente,
                sede if not modo_asignacion else '',
                conformidad if not modo_asignacion else '',
                ejecutado if not modo_asignacion else '',
                "Angie Rodriguez" if not modo_asignacion else '',
                "Vratislav Cala" if not modo_asignacion else ''
            ))
            certificados_asignados.append(cert_num)
            consecutivos[letra] += 1

    conn.commit()
    conn.close()
    return certificados_asignados

# -----------------------------------------
# BORRAR TODOS LOS DATOS DE LA BD
# -----------------------------------------
def borrar_todo():
    conn = sqlite3.connect('calibraciones.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM calibraciones")
    conn.commit()
    conn.close()

# -----------------------------------------
# DESCARGAR EXCEL DESDE LA BASE DE DATOS
# -----------------------------------------
def descargar_excel():
    conn = sqlite3.connect('calibraciones.db')
    df = pd.read_sql_query("SELECT orden AS 'Orden de Servicio', fecha AS 'Fecha', equipo AS 'Equipo', certificado AS 'Certificado', cliente AS 'Cliente', sede AS 'Sede/Servicio', conformidad AS 'Conformidad', ejecutado AS 'Ejecutado por', firmado AS 'Firmado por', avalado AS 'Avalado por' FROM calibraciones", conn)
    conn.close()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Calibraciones')
    st.download_button(label="📥 Descargar Excel", data=output.getvalue(), file_name="calibraciones.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------------------
# APP PRINCIPAL STREAMLIT
# -----------------------------------------
inicializar_db()

st.title("Asignación de Stickers de Calibración (SQLITE)")

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
else:
    sede = conformidad = ejecutado = ''

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

if 'equipos' not in st.session_state:
    st.session_state.equipos = []

if st.button("Agregar Equipo"):
    st.session_state.equipos.append((equipo, cantidad))

st.write("### Equipos añadidos:")
for eq, cant in st.session_state.equipos:
    st.write(f"{cant} x {eq}")

if st.button("Guardar Datos"):
    if not st.session_state.equipos:
        st.error("Debes agregar al menos un equipo")
    elif not orden or not cliente:
        st.error("Orden de Servicio y Cliente son obligatorios")
    else:
        certificados_asignados = guardar_calibraciones(orden, fecha, cliente, sede, conformidad, ejecutado, st.session_state.equipos, modo_asignacion, consecutivo_E_manual, consecutivo_B_manual)
        st.success(f"Certificados asignados: {', '.join(certificados_asignados)}")
        st.session_state.equipos = []

if st.button("🗑️ Borrar Todo"):
    borrar_todo()
    st.session_state.equipos = []
    st.success("Datos eliminados y base de datos reiniciada.")

descargar_excel()
