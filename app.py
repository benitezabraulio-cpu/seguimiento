import streamlit as st
from datetime import datetime
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import io
import base64
from PIL import Image

# Verificar que openpyxl esté instalado
try:
    from openpyxl import Workbook
except ImportError:
    st.error("❌ Falta instalar openpyxl. Ejecuta: pip install openpyxl")
    st.stop()

# Configuración de la página
st.set_page_config(
    page_title="Seguimiento de Obra - Fundación Masaveu",
    page_icon="🏗️",
    layout="wide"
)

# ===== CONFIGURACIÓN DE EMAIL =====
EMAIL_ORIGEN = "beniteza.braulio@alumnos25.fundacionmasaveu.com"
EMAIL_DESTINO = "ana@fundacionmasaveu.com"
# ===================================

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
if 'password_config' not in st.session_state:
    st.session_state.password_config = ''

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
    
    try:
        df = pd.DataFrame(st.session_state.registros)
        # Convertir hora_registro a string para Excel
        if 'hora_registro' in df.columns:
            df['hora_registro'] = df['hora_registro'].astype(str)
        
        # Eliminar columnas que no queremos mostrar
        if 'comentarios' in df.columns:
            # Mantener comentarios pero en otra posición
            pass
        
        output = io.BytesIO()
        # Usar xlsxwriter como alternativa si openpyxl falla
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Seguimiento Obra')
        except:
            # Fallback a xlsxwriter
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Seguimiento Obra')
        
        return output.getvalue()
    except Exception as e:
        st.error(f"Error al generar Excel: {str(e)}")
        return None

def enviar_email(excel_data, password):
    """Envía el Excel usando la configuración fija"""
    try:
        # Validar que la contraseña esté configurada
        if not password:
            return False, "❌ Primero introduce la contraseña de aplicación en el panel lateral"
        
        # Crear mensaje
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ORIGEN
        msg['To'] = EMAIL_DESTINO
        msg['Subject'] = f"Informe Seguimiento Obra - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Cuerpo del mensaje
        cuerpo = f"""Informe de Seguimiento de Obra

Fecha del informe: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
Total de registros: {len(st.session_state.registros)}
Trabajadores que han reportado: {len(set(r['trabajador'] for r in st.session_state.registros))}

Resumen de tareas:
{generar_resumen_tareas()}

Se adjunta archivo Excel con el detalle completo de todos los registros.

--
Sistema de Seguimiento de Obra
Fundación Masaveu
        """
        
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
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
        server.login(EMAIL_ORIGEN, password)
        server.send_message(msg)
        server.quit()
        
        return True, f"✅ Email enviado correctamente a {EMAIL_DESTINO}"
    except Exception as e:
        return False, f"❌ Error al enviar: {str(e)}"

def generar_resumen_tareas():
    """Genera un resumen de las tareas para el cuerpo del email"""
    if not st.session_state.registros:
        return "No hay registros."
    
    resumen = ""
    estados_count = {}
    tareas_por_trabajador = {}
    
    for r in st.session_state.registros:
        # Contar por estado
        estado = r['estado']
        estados_count[estado] = estados_count.get(estado, 0) + 1
        
        # Contar por trabajador
        trabajador = r['trabajador']
        if trabajador not in tareas_por_trabajador:
            tareas_por_trabajador[trabajador] = 0
        tareas_por_trabajador[trabajador] += 1
    
    resumen += "\n📊 Por estado:\n"
    for estado, count in estados_count.items():
        resumen += f"   - {estado[:40]}: {count} tareas\n"
    
    resumen += "\n👷 Por trabajador:\n"
    for trabajador, count in tareas_por_trabajador.items():
        resumen += f"   - {trabajador}: {count} tareas\n"
    
    return resumen

# Limpiar registros antiguos al inicio
limpiar_registros_antiguos()

# ==================== INTERFAZ PRINCIPAL ====================

st.title("🏗️ Sistema de Seguimiento de Obra - Fundación Masaveu")

# Barra lateral con configuración
with st.sidebar:
    st.markdown("### 📧 CONFIGURACIÓN DE EMAIL")
    st.markdown("---")
    
    st.markdown(f"**📤 Email ORIGEN:**")
    st.code(EMAIL_ORIGEN, language="")
    
    st.markdown(f"**📬 Email DESTINO:**")
    st.code(EMAIL_DESTINO, language="")
    
    st.markdown("---")
    st.markdown("**🔑 Contraseña de aplicación (Gmail):**")
    
    password = st.text_input(
        "Contraseña de aplicación",
        type="password",
        placeholder="xxxx xxxx xxxx xxxx",
        help="Necesitas generar una contraseña de aplicación en tu cuenta de Gmail",
        value=st.session_state.password_config
    )
    
    if st.button("💾 Guardar contraseña", use_container_width=True):
        st.session_state.password_config = password
        st.success("✅ Contraseña guardada")
    
    # Mostrar estado
    st.markdown("---")
    st.markdown("**Estado de configuración:**")
    
    if st.session_state.password_config:
        st.success("✅ Contraseña configurada")
        st.success(f"✅ Envíos a: {EMAIL_DESTINO}")
    else:
        st.error("❌ Contraseña no configurada")
    
    st.markdown("---")
    st.markdown("### ℹ️ Ayuda")
    st.markdown("""
    **¿Cómo obtener contraseña de aplicación?**
    1. Ve a tu cuenta de Google
    2. Activa verificación en 2 pasos
    3. Ve a "Contraseñas de aplicación"
    4. Selecciona "Otra" y pon nombre
    5. Copia la contraseña de 16 dígitos
    """)
    
    st.markdown("---")
    st.markdown(f"**📬 Los informes se envían a:**")
    st.info(f"**{EMAIL_DESTINO}**")

# Columnas para logo y título
col1, col2 = st.columns([1, 4])
with col1:
    try:
        # Si tienes logo, colócalo en la misma carpeta como 'logo.png'
        logo = Image.open("logo.png")
        st.image(logo, width=100)
    except:
        st.markdown("### 🏢")
        st.caption("Fundación Masaveu")

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
            # Verificar contraseña
            if not st.session_state.password_config:
                st.error("❌ Primero introduce la contraseña de aplicación en el panel lateral")
            else:
                with st.spinner("📧 Enviando email..."):
                    success, message = enviar_email(excel_data, st.session_state.password_config)
                    if success:
                        st.success(message)
                        st.balloons()
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

# Mostrar información de envío
st.info(f"📬 **Los informes se enviarán automáticamente a:** {EMAIL_DESTINO}")
