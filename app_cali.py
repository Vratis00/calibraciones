import streamlit as st
import openpyxl
import os
from datetime import date
from google.oauth2 import service_account
from googleapiclient.discovery import build

archivo_excel = 'calibraciones.xlsx'

# Inicializar Excel
def inicializar_excel():
    if not os.path.exists(archivo_excel):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Orden de Servicio', 'Fecha', 'Equipo', 'Certificado', 'Cliente', 'Sede/Servicio',
                   'Conformidad', 'Ejecutado por', 'Firmado por', 'Avalado por'])
        wb.save(archivo_excel)
    else:
        wb = openpyxl.load_workbook(archivo_excel)
        ws = wb.active
        if ws.max_row < 2:
            ws.append(['Orden de Servicio', 'Fecha', 'Equipo', 'Certificado', 'Cliente', 'Sede/Servicio',
                       'Conformidad', 'Ejecutado por', 'Firmado por', 'Avalado por'])
            wb.save(archivo_excel)

# Obtener último certificado
def obtener_ultimo_certificado(letra):
    wb = openpyxl.load_workbook(archivo_excel)
    ws = wb.active
    numeros = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        cert = row[3]
        if cert and cert.startswith(letra):
            num = int(cert.split('-')[-1])
            numeros.append(num)
    wb.close()
    if not numeros:
        return 1588 if letra == 'E' else 870
    return max(numeros)

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

# Subir archivo a Google Drive usando Service Account
def subir_a_drive_service_account(nombre_archivo):
    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    SERVICE_ACCOUNT_FILE = 'service_account.json'

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    service = build('drive', 'v3', credentials=credentials)

    file_metadata = {'name': nombre_archivo}
    media = MediaFileUpload(nombre_archivo, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    st.success(f"Archivo subido a Google Drive. ID: {file.get('id')}")

from googleapiclient.http import MediaFileUpload

# Inicialización
inicializar_excel()

# Interfaz de usuario en Streamlit
st.title("Asignación de Stickers de Calibración")

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

if st.button("Guardar y Subir a Google Drive"):
    if not st.session_state.equipos:
        st.error("Debes agregar al menos un equipo")
    elif not orden or not cliente:
        st.error("Orden de Servicio y Cliente son obligatorios")
    else:
        wb = openpyxl.load_workbook(archivo_excel)
        ws = wb.active

        # Limpiar filas
        ws.delete_rows(2, ws.max_row)

        certificados_asignados = []
        letras_equipos = {'E': [], 'B': []}

        for equipo, cantidad in st.session_state.equipos:
            letra = determinar_letra_equipo(equipo)
            if letra == 'X':
                st.error(f"Equipo inválido: {equipo}")
                st.stop()
            letras_equipos[letra].append((equipo, cantidad))

        for letra, lista in letras_equipos.items():
            if not lista:
                continue
            ultimo = obtener_ultimo_certificado(letra)
            for equipo, cantidad in lista:
                for i in range(cantidad):
                    nuevo_num = ultimo + 1
                    cert_num = f"{letra}-25-{nuevo_num}"
                    ws.append([
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
                    ])
                    certificados_asignados.append(cert_num)
                    ultimo += 1

        wb.save(archivo_excel)
        wb.close()
        st.success(f"Certificados asignados: {', '.join(certificados_asignados)}")
        subir_a_drive_service_account(archivo_excel)
        st.session_state.equipos = []
