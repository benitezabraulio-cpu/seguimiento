import streamlit as st
from datetime import datetime
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import io
import base64
from PIL import Image

# Configuración de la página
st.set_page_config(
    page_title="Seguimiento de Obra",
    page_icon="🏗️",
    layout="wide"
)

# Lista de tareas predefinidas
TAREAS = [
    "Trazado y marcado de cajas, tubos y cuadros",
    "Ejecución rozas en paredes y techos",
    "Montaje de soportes",
    "Colocación tubos y conductos",
    "Tendido de cables",
    "Identificación y etiquetado",
    "Conexionado de cables en bornes o regletas",
    "Instalación y conexionado de mecanismos",
    "Fijación de carril DIN y mecanismos en cuadro eléctrico",
    "Cableado interno del cuadro eléctrico",
    "Configuración de equipos domóticos y/o automáticos",
    "Conexionado de sensores/actuadores de equipos domóticos/automáticos",
    "Pruebas de continuidad",
    "Pruebas de aislamiento",
    "Verificación de tierras",
    "Programación del automatismo",
    "Pruebas de funcionamiento"
]

# Estados de avance
ESTADOS = [
    "Avance de la tarea en torno al 25% aprox.",
    "Avance de la tarea en torno al 50% aprox.",
    "Avance de la tarea en torno al 75% aprox.",
    "OK, finalizado sin errores",
    "Finalizado, pero con errores pendientes de corregir",
    "Finalizado y corregidos los errores"
]

# Inicializar session state
if 'registros' not in st.session_state:
    st.session_state.registros = []
if 'ultima_actualizacion' not in st.session_state:
    st.session_state.ultima_actualizacion = datetime.now()

def limpiar_registros_antiguos():
    """Elimina registros de más de 2 horas"""
    if st.session_state.registros:
        ahora = datetime.now()
        registros_validos = []
        for registro in st.session_state.registros:
            hora_registro = registro.get('hora_registro')
            if hora_registro:
                try:
                    diferencia = (ahora - hora_registro).total_seconds() / 3600
                    if diferencia <= 2:
                        registros_validos.append(registro)
                except:
                    registros_validos.append(registro)
        st.session_state.registros = registros_validos

def guardar_registro(fecha, trabajador, tarea, estado, comentarios):
    """Guarda un nuevo registro"""
    registro = {
        'fecha': fecha,
        'trabajador': trabajador,
        'tarea': tarea,
        'estado': estado,
        'comentarios': comentarios,
        'hora_registro': datetime.now()
    }
    st.session_state.registros.append(registro)
    return True

def generar_excel():
    """Genera archivo Excel con los registros"""
    if not st.session_state.registros:
        return None
    
    df = pd.DataFrame(st.session_state.registros)
    # Convertir hora_registro a string para Excel
    if 'hora_registro' in df.columns:
        df['hora_registro'] = df['hora_registro'].astype(str)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Seguimiento Obra')
    
    return output.getvalue()

def enviar_email(excel_data, email_destino):
    """Envía el Excel por email"""
    try:
        # Configuración SMTP (para Gmail)
        email_origen = "tu_email@gmail.com"  # Cambiar
        password = "tu_contraseña_app"  # Cambiar (usar contraseña de aplicación)
        
        msg = MIMEMultipart()
        msg['From'] = email_origen
        msg['To'] = email_destino
        msg['Subject'] = f"Informe Seguimiento Obra - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Adjuntar Excel
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(excel_data)
        encoders.encode_base64(attachment)
        attachment.add_header(
            'Content-Disposition',
            f'attachment; filename=seguimiento_obra_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        msg.attach(attachment)
        
        # Enviar
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_origen, password)
        server.send_message(msg)
        server.quit()
        
        return True, "Email enviado correctamente"
    except Exception as e:
        return False, f"Error al enviar: {str(e)}"

def descargar_excel():
    """Permite descargar el Excel desde el navegador"""
    excel_data = generar_excel()
    if excel_data:
        b64 = base64.b64encode(excel_data).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="seguimiento_obra.xlsx">📥 Descargar Excel</a>'
        return href
    return None

# Limpiar registros antiguos al inicio
limpiar_registros_antiguos()

# Interfaz principal
st.title("🏗️ Sistema de Seguimiento de Obra")

# Columnas para logo y título
col1, col2 = st.columns([1, 4])
with col1:
    try:
        # Si tienes logo, colócalo en la misma carpeta como 'logo.png'
        logo = Image.open("logo.png")
        st.image(logo, width=100)
    except:
        st.markdown("### 🏢")
        st.caption("Logo Empresa")

with col2:
    st.markdown("## Control de Avance de Obra")
    st.markdown("---")

# Crear dos columnas principales
col_form, col_registros = st.columns([1, 1])

with col_form:
    st.markdown("### 📝 Nuevo Registro")
    
    with st.form("registro_form"):
        fecha = st.date_input("📅 Fecha del informe", datetime.now())
        trabajador = st.text_input("👷 Nombre del trabajador")
        tarea = st.selectbox("📋 Seleccionar tarea", TAREAS)
        estado = st.selectbox("📊 Estado de la tarea", ESTADOS)
        comentarios = st.text_area("💬 Comentarios adicionales", height=100)
        
        submitted = st.form_submit_button("✅ GUARDAR REGISTRO", use_container_width=True)
        
        if submitted:
            if not trabajador:
                st.error("❌ Por favor ingrese el nombre del trabajador")
            else:
                if guardar_registro(fecha.strftime("%d/%m/%Y"), trabajador, tarea, estado, comentarios):
                    st.success("✅ Registro guardado correctamente")
                    st.rerun()

with col_registros:
    st.markdown("### 📋 Registros Recientes")
    
    if st.session_state.registros:
        # Mostrar últimos 10 registros
        df_mostrar = pd.DataFrame(st.session_state.registros[-10:])
        df_mostrar = df_mostrar[['fecha', 'trabajador', 'tarea', 'estado']]
        df_mostrar.columns = ['Fecha', 'Trabajador', 'Tarea', 'Estado']
        st.dataframe(df_mostrar, use_container_width=True, height=300)
        
        # Estadísticas rápidas
        total_registros = len(st.session_state.registros)
        tareas_completadas = sum(1 for r in st.session_state.registros if "finalizado" in r['estado'].lower())
        
        col1_stat, col2_stat, col3_stat = st.columns(3)
        with col1_stat:
            st.metric("Total Registros", total_registros)
        with col2_stat:
            st.metric("Tareas Completadas", tareas_completadas)
        with col3_stat:
            horas_restantes = 2 - (datetime.now() - st.session_state.ultima_actualizacion).total_seconds() / 3600
            st.metric("Horas hasta limpieza", f"{max(0, horas_restantes):.1f}h")
    else:
        st.info("ℹ️ No hay registros aún. Complete el formulario para comenzar.")

# Sección de exportación
st.markdown("---")
st.markdown("### 📤 Exportar Datos")

col_export1, col_export2, col_export3 = st.columns(3)

with col_export1:
    if st.button("📊 GENERAR EXCEL", use_container_width=True):
        excel_data = generar_excel()
        if excel_data:
            # Crear enlace de descarga
            b64 = base64.b64encode(excel_data).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="seguimiento_obra_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx" style="text-decoration: none;">✅ Click aquí para descargar Excel</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("✅ Excel generado correctamente")
        else:
            st.warning("⚠️ No hay datos para generar el Excel")

with col_export2:
    with st.expander("📧 Enviar por Email"):
        email_destino = st.text_input("Email destino", placeholder="ejemplo@empresa.com")
        if st.button("Enviar", use_container_width=True):
            if not email_destino:
                st.error("❌ Ingrese un email destino")
            else:
                excel_data = generar_excel()
                if excel_data:
                    with st.spinner("Enviando email..."):
                        success, message = enviar_email(excel_data, email_destino)
                        if success:
                            st.success(message)
                        else:
                            st.error(message)
                else:
                    st.warning("⚠️ No hay datos para enviar")

with col_export3:
    if st.button("🗑️ LIMPIAR FORMULARIO", use_container_width=True):
        st.session_state.form_limpiado = True
        st.rerun()

# Aviso de tiempo
st.markdown("---")
st.warning("⚠️ **NOTA IMPORTANTE:** Los datos solo se guardan durante 2 horas. "
          "Descargue el Excel o envíelo por email para conservar el registro.")

# Sidebar con información
with st.sidebar:
    st.markdown("### 📊 Resumen General")
    
    if st.session_state.registros:
        # Resumen por trabajador
        df_resumen = pd.DataFrame(st.session_state.registros)
        st.markdown("#### 👷 Actividad por trabajador:")
        trabajador_counts = df_resumen['trabajador'].value_counts()
        for trabajador, count in trabajador_counts.items():
            st.write(f"- {trabajador}: {count} tareas")
        
        st.markdown("#### 📈 Estado de tareas:")
        estado_counts = df_resumen['estado'].value_counts()
        for estado, count in estado_counts.items():
            st.write(f"- {estado[:30]}: {count}")
    else:
        st.info("No hay registros aún")
    
    st.markdown("---")
    st.markdown("### 📱 Instrucciones")
    st.markdown("""
    1. Complete el formulario con los datos de la tarea
    2. Presione 'Guardar Registro'
    3. Genere el Excel cuando necesite exportar
    4. Envíe por email para archivo permanente
    
    **Tiempo de retención:** 2 horas
    """)

# Actualizar timestamp periódicamente
if st.button("🔄 Actualizar"):
    limpiar_registros_antiguos()
    st.rerun()
