import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageTk
import json
import threading

class SeguimientoObraApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Seguimiento de Obra - Sistema de Control")
        self.root.geometry("1100x750")
        
        # Configurar colores
        self.colores = {
            'fondo': '#f0f0f0',
            'boton': '#4CAF50',
            'boton_texto': 'white',
            'frame': '#ffffff'
        }
        
        self.root.configure(bg=self.colores['fondo'])
        
        # Archivo de datos
        self.archivo_datos = "registros_obra.xlsx"
        self.crear_archivo_si_no_existe()
        
        # Lista de tareas predefinidas
        self.tareas = [
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
        self.estados = [
            "Avance de la tarea en torno al 25% aprox.",
            "Avance de la tarea en torno al 50% aprox.",
            "Avance de la tarea en torno al 75% aprox.",
            "OK, finalizado sin errores",
            "Finalizado, pero con errores pendientes de corregir",
            "Finalizado y corregidos los errores"
        ]
        
        self.crear_interfaz()
        self.cargar_registros_actuales()
    
    def crear_archivo_si_no_existe(self):
        """Crea el archivo Excel si no existe con las columnas necesarias"""
        if not os.path.exists(self.archivo_datos):
            df = pd.DataFrame(columns=[
                "Fecha",
                "Trabajador",
                "Tarea",
                "Estado",
                "Comentarios",
                "Hora_Registro"
            ])
            df.to_excel(self.archivo_datos, index=False)
    
    def crear_interfaz(self):
        # Frame superior con logo
        frame_superior = tk.Frame(self.root, bg=self.colores['fondo'], height=120)
        frame_superior.pack(fill=tk.X, padx=10, pady=5)
        
        # Cargar y mostrar logo
        try:
            img = Image.open("logo_empresa.png")
            img = img.resize((100, 100), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(img)
            lbl_logo = tk.Label(frame_superior, image=self.logo, bg=self.colores['fondo'])
            lbl_logo.pack(side=tk.LEFT, padx=10)
        except:
            # Si no hay logo, mostrar texto
            lbl_logo_texto = tk.Label(frame_superior, text="LOGO EMPRESA", 
                                     font=("Arial", 20, "bold"),
                                     bg=self.colores['fondo'])
            lbl_logo_texto.pack(side=tk.LEFT, padx=10)
        
        titulo = tk.Label(frame_superior, text="SISTEMA DE SEGUIMIENTO DE OBRA",
                         font=("Arial", 18, "bold"),
                         bg=self.colores['fondo'])
        titulo.pack(side=tk.LEFT, padx=20)
        
        # Frame principal de entrada de datos
        frame_principal = tk.Frame(self.root, bg=self.colores['frame'],
                                  relief=tk.RAISED, borderwidth=2)
        frame_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Título del formulario
        tk.Label(frame_principal, text="REGISTRO DE AVANCE DE OBRA",
                font=("Arial", 14, "bold"),
                bg=self.colores['frame']).grid(row=0, column=0, columnspan=2, pady=10)
        
        # Campo: Fecha
        tk.Label(frame_principal, text="Fecha del informe:", 
                font=("Arial", 10), bg=self.colores['frame']).grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        self.fecha_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.fecha_entry = tk.Entry(frame_principal, textvariable=self.fecha_var, width=30)
        self.fecha_entry.grid(row=1, column=1, padx=10, pady=5)
        
        # Campo: Trabajador
        tk.Label(frame_principal, text="Nombre del trabajador:", 
                font=("Arial", 10), bg=self.colores['frame']).grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        self.trabajador_var = tk.StringVar()
        self.trabajador_entry = tk.Entry(frame_principal, textvariable=self.trabajador_var, width=40)
        self.trabajador_entry.grid(row=2, column=1, padx=10, pady=5)
        
        # Desplegable: Tareas
        tk.Label(frame_principal, text="Seleccionar tarea:", 
                font=("Arial", 10), bg=self.colores['frame']).grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        self.tarea_var = tk.StringVar()
        self.tarea_combobox = ttk.Combobox(frame_principal, textvariable=self.tarea_var,
                                          values=self.tareas, width=50)
        self.tarea_combobox.grid(row=3, column=1, padx=10, pady=5)
        
        # Desplegable: Estados
        tk.Label(frame_principal, text="Estado de la tarea:", 
                font=("Arial", 10), bg=self.colores['frame']).grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        self.estado_var = tk.StringVar()
        self.estado_combobox = ttk.Combobox(frame_principal, textvariable=self.estado_var,
                                          values=self.estados, width=50)
        self.estado_combobox.grid(row=4, column=1, padx=10, pady=5)
        
        # Campo: Comentarios
        tk.Label(frame_principal, text="Comentarios adicionales:", 
                font=("Arial", 10), bg=self.colores['frame']).grid(row=5, column=0, sticky=tk.NW, padx=10, pady=5)
        self.comentarios_text = tk.Text(frame_principal, height=5, width=50)
        self.comentarios_text.grid(row=5, column=1, padx=10, pady=5)
        
        # Frame de botones
        frame_botones = tk.Frame(frame_principal, bg=self.colores['frame'])
        frame_botones.grid(row=6, column=0, columnspan=2, pady=20)
        
        # Botón Guardar Registro
        btn_guardar = tk.Button(frame_botones, text="GUARDAR REGISTRO",
                               bg=self.colores['boton'], fg=self.colores['boton_texto'],
                               font=("Arial", 10, "bold"), padx=20, pady=10,
                               command=self.guardar_registro)
        btn_guardar.pack(side=tk.LEFT, padx=10)
        
        # Botón Generar Excel
        btn_excel = tk.Button(frame_botones, text="GENERAR EXCEL",
                             bg="#2196F3", fg="white",
                             font=("Arial", 10, "bold"), padx=20, pady=10,
                             command=self.generar_excel)
        btn_excel.pack(side=tk.LEFT, padx=10)
        
        # Botón Enviar Email
        btn_email = tk.Button(frame_botones, text="ENVIAR POR EMAIL",
                             bg="#FF9800", fg="white",
                             font=("Arial", 10, "bold"), padx=20, pady=10,
                             command=self.enviar_email)
        btn_email.pack(side=tk.LEFT, padx=10)
        
        # Botón Limpiar formulario
        btn_limpiar = tk.Button(frame_botones, text="LIMPIAR FORMULARIO",
                               bg="#9E9E9E", fg="white",
                               font=("Arial", 10, "bold"), padx=20, pady=10,
                               command=self.limpiar_formulario)
        btn_limpiar.pack(side=tk.LEFT, padx=10)
        
        # Frame para mostrar registros
        frame_registros = tk.Frame(self.root, bg=self.colores['frame'],
                                  relief=tk.RAISED, borderwidth=2)
        frame_registros.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tk.Label(frame_registros, text="REGISTROS RECIENTES",
                font=("Arial", 12, "bold"),
                bg=self.colores['frame']).pack(pady=5)
        
        # Treeview para mostrar registros
        scrollbar = tk.Scrollbar(frame_registros)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = ttk.Treeview(frame_registros, yscrollcommand=scrollbar.set,
                                columns=("Fecha", "Trabajador", "Tarea", "Estado"),
                                show="headings", height=8)
        
        self.tree.heading("Fecha", text="Fecha")
        self.tree.heading("Trabajador", text="Trabajador")
        self.tree.heading("Tarea", text="Tarea")
        self.tree.heading("Estado", text="Estado")
        
        self.tree.column("Fecha", width=100)
        self.tree.column("Trabajador", width=150)
        self.tree.column("Tarea", width=400)
        self.tree.column("Estado", width=300)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.tree.yview)
        
        # Aviso de tiempo
        frame_aviso = tk.Frame(self.root, bg="#FFEB3B", height=30)
        frame_aviso.pack(fill=tk.X, padx=10, pady=5)
        
        aviso_texto = "⚠️ NOTA: Los datos solo se guardan por 2 horas. Envía el Excel por email para conservar el registro."
        tk.Label(frame_aviso, text=aviso_texto, font=("Arial", 9),
                bg="#FFEB3B", fg="#333").pack(pady=5)
    
    def guardar_registro(self):
        """Guarda el registro en el archivo Excel"""
        # Validar campos
        if not self.trabajador_var.get():
            messagebox.showwarning("Campo requerido", "Por favor ingrese el nombre del trabajador")
            return
        if not self.tarea_var.get():
            messagebox.showwarning("Campo requerido", "Por favor seleccione una tarea")
            return
        if not self.estado_var.get():
            messagebox.showwarning("Campo requerido", "Por favor seleccione un estado")
            return
        
        # Crear registro
        registro = {
            "Fecha": self.fecha_var.get(),
            "Trabajador": self.trabajador_var.get(),
            "Tarea": self.tarea_var.get(),
            "Estado": self.estado_var.get(),
            "Comentarios": self.comentarios_text.get("1.0", tk.END).strip(),
            "Hora_Registro": datetime.now().strftime("%H:%M:%S")
        }
        
        # Guardar en Excel
        try:
            df = pd.read_excel(self.archivo_datos)
            df_nuevo = pd.DataFrame([registro])
            df = pd.concat([df, df_nuevo], ignore_index=True)
            df.to_excel(self.archivo_datos, index=False)
            
            # Actualizar visualización
            self.cargar_registros_actuales()
            self.limpiar_formulario()
            
            messagebox.showinfo("Éxito", "Registro guardado correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el registro: {str(e)}")
    
    def cargar_registros_actuales(self):
        """Carga y muestra los registros actuales en el Treeview"""
        # Limpiar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            if os.path.exists(self.archivo_datos):
                df = pd.read_excel(self.archivo_datos)
                # Mostrar últimos 20 registros
                for idx, row in df.tail(20).iterrows():
                    self.tree.insert("", "end", values=(
                        row.get("Fecha", ""),
                        row.get("Trabajador", ""),
                        row.get("Tarea", ""),
                        row.get("Estado", "")
                    ))
        except Exception as e:
            print(f"Error al cargar registros: {e}")
    
    def generar_excel(self):
        """Genera y guarda el archivo Excel actualizado"""
        try:
            fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_archivo = f"seguimiento_obra_{fecha_actual}.xlsx"
            
            # Crear directorio exports si no existe
            if not os.path.exists("exports"):
                os.makedirs("exports")
            
            ruta_archivo = os.path.join("exports", nombre_archivo)
            
            # Copiar el archivo actual
            if os.path.exists(self.archivo_datos):
                import shutil
                shutil.copy(self.archivo_datos, ruta_archivo)
                messagebox.showinfo("Éxito", f"Excel generado correctamente en:\n{ruta_archivo}")
            else:
                messagebox.showwarning("Advertencia", "No hay datos para generar el Excel")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el Excel: {str(e)}")
    
    def enviar_email(self):
        """Envía el archivo Excel por email"""
        # Configuración de email (DEBES CONFIGURAR TUS DATOS)
        email_origen = "tu_email@gmail.com"  # Cambiar por email de la empresa
        password = "tu_contraseña"  # Cambiar por contraseña
        email_destino = "empresa@dominio.com"  # Cambiar por email destino
        
        if not os.path.exists(self.archivo_datos):
            messagebox.showwarning("Advertencia", "No hay datos para enviar")
            return
        
        try:
            # Crear mensaje
            msg = MIMEMultipart()
            msg['From'] = email_origen
            msg['To'] = email_destino
            msg['Subject'] = f"Informe de Seguimiento de Obra - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            
            # Adjuntar archivo
            with open(self.archivo_datos, "rb") as adjunto:
                parte = MIMEBase("application", "octet-stream")
                parte.set_payload(adjunto.read())
            
            encoders.encode_base64(parte)
            parte.add_header("Content-Disposition", f"attachment; filename=seguimiento_obra.xlsx")
            msg.attach(parte)
            
            # Enviar email (en un hilo separado para no bloquear la interfaz)
            def enviar():
                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(email_origen, password)
                    server.send_message(msg)
                    server.quit()
                    self.root.after(0, lambda: messagebox.showinfo("Éxito", "Email enviado correctamente"))
                except Exception as e:
                    self.root.after(0, lambda: messagebox.showerror("Error", f"No se pudo enviar el email: {str(e)}"))
            
            threading.Thread(target=enviar, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo preparar el email: {str(e)}")
    
    def limpiar_formulario(self):
        """Limpia todos los campos del formulario"""
        self.trabajador_var.set("")
        self.tarea_var.set("")
        self.estado_var.set("")
        self.comentarios_text.delete("1.0", tk.END)
        self.fecha_var.set(datetime.now().strftime("%d/%m/%Y"))

if __name__ == "__main__":
    # Instalar dependencias necesarias
    # pip install pandas openpyxl pillow
    
    root = tk.Tk()
    app = SeguimientoObraApp(root)
    
    # Configurar temporizador para limpiar datos después de 2 horas
    def limpiar_datos_antiguos():
        """Elimina registros de más de 2 horas"""
        try:
            if os.path.exists(app.archivo_datos):
                df = pd.read_excel(app.archivo_datos)
                hora_actual = datetime.now()
                
                # Filtrar registros de las últimas 2 horas
                registros_validos = []
                for idx, row in df.iterrows():
                    hora_registro = row.get("Hora_Registro", "00:00:00")
                    fecha_registro = row.get("Fecha", datetime.now().strftime("%d/%m/%Y"))
                    
                    try:
                        fecha_hora_str = f"{fecha_registro} {hora_registro}"
                        fecha_hora = datetime.strptime(fecha_hora_str, "%d/%m/%Y %H:%M:%S")
                        diferencia = (hora_actual - fecha_hora).total_seconds() / 3600
                        
                        if diferencia <= 2:
                            registros_validos.append(row)
                    except:
                        registros_validos.append(row)
                
                if len(registros_validos) > 0:
                    df_nuevo = pd.DataFrame(registros_validos)
                    df_nuevo.to_excel(app.archivo_datos, index=False)
                    app.cargar_registros_actuales()
        except Exception as e:
            print(f"Error al limpiar datos: {e}")
        
        # Programar próxima limpieza (cada 30 minutos)
        root.after(1800000, limpiar_datos_antiguos)
    
    # Iniciar limpieza automática
    root.after(1800000, limpiar_datos_antiguos)
    
    root.mainloop()
