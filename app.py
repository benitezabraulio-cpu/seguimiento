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
if 'email_config' not in st.session_state:
    st.session_state.email_config = {
        'origen': '',
        'password': '',
        'destino': ''
    }

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

def enviar_email(excel_data):
    """Envía el Excel usando la configuración guardada"""
    try:
        config = st.session_state.email_config
        
        # Validar que la configuración esté completa
        if not config['origen']:
            return False, "❌ Primero configura el EMAIL ORIGEN en el panel lateral"
        if not config['password']:
            return False, "❌ Primero configura la CONTRASEÑA en el panel lateral"
        if not config['destino']:
            return False, "❌ Primero configura el EMAIL DESTINO en el panel lateral"
        
        # Crear mensaje
        msg = MIMEMultipart()
        msg['From'] = config['origen']
        msg['To'] = config['destino']
        msg['Subject'] = f"Informe Seguimiento Obra - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Cuerpo del mensaje
        cuerpo = f"""Informe de Seguimiento de Obra

Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
Total de registros: {len(st.session_state.registros)}
Trabajadores: {len(set(r['trabajador'] for r in st.session_state.registros))}

Se adjunta archivo Excel con el detalle completo.
        """
        msg.attach(MIMEBase('text', 'plain', _charset='utf-8'))
        
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
        server.login(config['origen'], config['password'])
        server.send_message(msg)
        server.quit()
        
        return True, f"✅ Email enviado correctamente a {config['destino']}"
    except Exception as e:
        return False, f"❌ Error al enviar: {str(e)}"

# Limpiar registros antiguos al inicio
limpiar_registros_antiguos()

# ==================== INTERFAZ PRINCIPAL ====================

st.title("🏗️ Sistema de Seguimiento de Obra")

# Barra lateral con configuración de email
with st.sidebar:
    st.markdown("### 📧 CONFIGURACIÓN DE EMAIL")
    st.markdown("---")
    
    st.markdown("**Configuración del correo electrónico:**")
    
    email_origen = st.text_input(
        "📤 Email ORIGEN (el que envía)",
        value=st.session_state.email_config['origen'],
        placeholder="tuempresa@gmail.com",
        help="Correo desde donde se enviarán los informes"
    )
    
    email_password = st.text_input(
        "🔑 Contraseña de aplicación",
        value=st.session_state.email_config['password'],
        type="password",
        placeholder="xxxx xxxx xxxx xxxx",
        help="Usar contraseña de aplicación de Gmail (no tu contraseña normal)"
    )
    
    st.markdown("---")
    st.markdown("**📥 Email DESTINO (donde quieres recibir los informes):**")
    
    email_destino = st.text_input(
        "📬 Email DESTINO",
        value=st.session_state.email_config['destino'],
        placeholder="informes@tuempresa.com",
        help="Correo donde quieres recibir los informes"
    )
    
    col1_btn, col2_btn = st.columns(2)
    with col1_btn:
        if st.button("💾 Guardar configuración", use_container_width=True):
            st.session_state.email_config['origen'] = email_origen
            st.session_state.email_config['password'] = email_password
            st.session_state.email_config['destino'] = email_destino
            st.success("✅ Configuración guardada")
    
    with col2_btn:
        if st.button("🗑️ Limpiar", use_container_width=True):
            st.session_state.email_config = {'origen': '', 'password': '', 'destino': ''}
            st.rerun()
    
    # Mostrar estado de la configuración
    st.markdown("---")
    st.markdown("**Estado de configuración:**")
    
    if st.session_state.email_config['origen']:
        st.success("✅ Email origen configurado")
    else:
        st.error("❌ Email origen no configurado")
    
    if st.session_state.email_config['destino']:
        st.success(f"✅ Email destino: {st.session_state.email_config['destino']}")
    else:
        st.error("❌ Email destino no configurado")
    
    if st.session_state.email_config['password']:
        st.success("✅ Contraseña configurada")
    else:
        st.error("❌ Contraseña no configurada")
    
    st.markdown("---")
    st.markdown("### ℹ️ Ayuda")
    st.markdown("""
    **¿Cómo obtener contraseña de aplicación?**
    1. Ve a tu cuenta de Google
    2. Activa verificación en 2 pasos
    3. Genera contraseña de aplicación
    """)

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
            if st.session_state.registros:
                primer_registro = min(r['hora_registro'] for r in st.session_state.registros if 'hora_registro' in r)
                horas_transcurridas = (datetime.now() - primer_registro).total_seconds() / 3600
                horas_restantes = max(0, 2 - horas_transcurridas)
                st.metric("Horas hasta limpieza", f"{horas_restantes:.1f}h")
            else:
                st.metric("Horas hasta limpieza", "0h")
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
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="seguimiento_obra_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx" style="text-decoration: none; background-color: #4CAF50; color: white; padding: 10px 20px; border-radius: 5px; display: inline-block;">✅ HAZ CLIC PARA DESCARGAR EXCEL</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("✅ Excel generado correctamente")
        else:
            st.warning("⚠️ No hay datos para generar el Excel")

with col_export2:
    if st.button("📧 ENVIAR POR EMAIL", use_container_width=True):
        excel_data = generar_excel()
        if excel_data:
            # Verificar configuración primero
            if not st.session_state.email_config['origen'] or not st.session_state.email_config['password'] or not st.session_state.email_config['destino']:
                st.error("❌ Primero configura el email en el panel lateral")
            else:
                with st.spinner("📧 Enviando email..."):
                    success, message = enviar_email(excel_data)
                    if success:
                        st.success(message)
                    else:
                        st.error(message)
        else:
            st.warning("⚠️ No hay datos para enviar")

with col_export3:
    if st.button("🔄 ACTUALIZAR", use_container_width=True):
        limpiar_registros_antiguos()
        st.rerun()

# Aviso de tiempo
st.markdown("---")
st.warning("⚠️ **NOTA IMPORTANTE:** Los datos solo se guardan durante 2 horas. "
          "Descargue el Excel o envíelo por email para conservar el registro.")

# Mostrar email destino actual si está configurado
if st.session_state.email_config['destino']:
    st.info(f"📬 Los informes se enviarán a: **{st.session_state.email_config['destino']}**")
