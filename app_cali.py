import streamlit as st
import sqlite3
import os
from datetime import date
import pandas as pd

archivo_db = 'calibraciones.db'

# Inicializar Base de Datos
def inicializar_db():
    conn = sqlite3.connect(archivo_db)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS calibraciones (
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

if 'equipos' not in st.session_state:
    st.session_state.equipos = []

if st.button("Agregar Equipo"):
    st.session_state.equipos.append((equipo, cantidad))

st.write("### Equipos añadidos:")
for eq, cant in st.session_state.equipos:
    st.write(f"{cant} x {eq}")

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

        for equipo, cantidad in st.session_state.equipos:
            letra = determinar_letra_equipo(equipo)
            if letra == 'X':
                st.error(f"Equipo inválido: {equipo}")
                st.stop()
            for i in range(cantidad):
                cert_num = f"{letra}-25-{consecutivos[letra]}"
                c.execute("""
                    INSERT INTO calibraciones (orden, fecha, equipo, certificado, cliente, sede, conformidad, ejecutado, firmado, avalado)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
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
        st.success(f"Certificados asignados: {', '.join(certificados_asignados)}")
        st.session_state.equipos = []

if st.button("🗑️ Borrar Todo"):
    conn = sqlite3.connect(archivo_db)
    c = conn.cursor()
    c.execute("DELETE FROM calibraciones")
    conn.commit()
    conn.close()
    st.session_state.equipos = []
    st.success("Datos eliminados y base de datos reiniciada.")

# BOTÓN DE DESCARGA DE BASE DE DATOS
if os.path.exists(archivo_db):
    with open(archivo_db, "rb") as file:
        st.download_button(
            label="📥 Descargar Base de Datos (.db)",
            data=file,
            file_name=archivo_db,
            mime="application/x-sqlite3"
        )

# MOSTRAR REGISTROS DE LA BASE DE DATOS EN TABLA
if st.button("📊 Visualizar Base de Datos"):
    conn = sqlite3.connect(archivo_db)
    df = pd.read_sql_query("SELECT * FROM calibraciones", conn)
    conn.close()

    st.dataframe(df)
    st.success("Registros cargados exitosamente")

