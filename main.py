import os
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from tkinter import (
    Tk, filedialog, Button, Label, Text, Scrollbar, END,
    StringVar, Entry, Frame
)

# === Lógica de envío ===
def enviar_correos(excel_path, docs_folder, remitente, clave, smtp_server, smtp_port, log_widget):
    try:
        df = pd.read_excel(excel_path)
        clientes = df[['DNI', 'Email']].dropna()

        for archivo in os.listdir(docs_folder):
            for _, row in clientes.iterrows():
                dni = str(row['DNI'])
                email = row['Email']
                if archivo.startswith(dni):
                    ruta_doc = os.path.join(docs_folder, archivo)

                    # Crear email
                    mensaje = EmailMessage()
                    mensaje['Subject'] = 'Documento para usted'
                    mensaje['From'] = remitente
                    mensaje['To'] = email
                    mensaje.set_content(f"Hola,\nAdjunto encontrarás el documento correspondiente a tu DNI ({dni}).\n\nSaludos.")

                    tipo_mime, _ = mimetypes.guess_type(ruta_doc)
                    if tipo_mime is None:
                        tipo_mime = 'application/octet-stream'

                    with open(ruta_doc, 'rb') as f:
                        contenido = f.read()
                        mensaje.add_attachment(
                            contenido,
                            maintype=tipo_mime.split('/')[0],
                            subtype=tipo_mime.split('/')[1],
                            filename=archivo
                        )

                    # Enviar
                    try:
                        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
                            smtp.starttls()
                            smtp.login(remitente, clave)
                            smtp.send_message(mensaje)
                            log_widget.insert(END, f"✅ Enviado: {archivo} a {email}\n")
                    except Exception as e:
                        log_widget.insert(END, f"❌ Error al enviar '{archivo}' a {email}: {e}\n")

        log_widget.insert(END, "\n🚀 ¡Proceso terminado!\n")
    except Exception as e:
        log_widget.insert(END, f"⚠️ Error general: {e}\n")


# === Interfaz gráfica ===
def main():
    def seleccionar_excel():
        ruta = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta:
            excel_label.config(text=f"📄 Excel: {ruta}")
            excel_path.set(ruta)

    def seleccionar_carpeta():
        ruta = filedialog.askdirectory()
        if ruta:
            carpeta_label.config(text=f"📁 Carpeta: {ruta}")
            docs_folder.set(ruta)

    def lanzar_envio():
        remitente = remitente_var.get().strip()
        clave = clave_var.get().strip()
        smtp_server = servidor_var.get().strip()
        smtp_port = int(puerto_var.get().strip())

        if not all([remitente, clave, excel_path.get(), docs_folder.get(), smtp_server, smtp_port]):
            log_text.insert(END, "⚠️ Debes completar todos los campos y seleccionar Excel y carpeta.\n")
            return

        enviar_correos(
            excel_path.get(),
            docs_folder.get(),
            remitente,
            clave,
            smtp_server,
            smtp_port,
            log_text
        )

    root = Tk()
    root.title("📬 Envío automático de documentos")
    root.geometry("750x600")

    # Variables
    remitente_var = StringVar()
    clave_var = StringVar()
    servidor_var = StringVar(value="smtp.gmail.com")
    puerto_var = StringVar(value="587")
    excel_path = StringVar()
    docs_folder = StringVar()

    # === Inputs de configuración ===
    config_frame = Frame(root)
    config_frame.pack(pady=5)

    Label(config_frame, text="📧 Email remitente:").grid(row=0, column=0, sticky="w")
    Entry(config_frame, textvariable=remitente_var, width=40).grid(row=0, column=1, padx=5)

    Label(config_frame, text="🔑 Clave de aplicación:").grid(row=1, column=0, sticky="w")
    Entry(config_frame, textvariable=clave_var, show='*', width=40).grid(row=1, column=1, padx=5)

    Label(config_frame, text="🌐 Servidor SMTP:").grid(row=2, column=0, sticky="w")
    Entry(config_frame, textvariable=servidor_var, width=40).grid(row=2, column=1, padx=5)

    Label(config_frame, text="📡 Puerto:").grid(row=3, column=0, sticky="w")
    Entry(config_frame, textvariable=puerto_var, width=10).grid(row=3, column=1, sticky="w", padx=5)

    # === Selección de Excel y carpeta ===
    Button(root, text="Seleccionar Excel", command=seleccionar_excel).pack(pady=5)
    excel_label = Label(root, text="📄 Excel: No seleccionado")
    excel_label.pack()

    Button(root, text="Seleccionar Carpeta", command=seleccionar_carpeta).pack(pady=5)
    carpeta_label = Label(root, text="📁 Carpeta: No seleccionada")
    carpeta_label.pack()

    # === Botón de envío ===
    Button(root, text="🚀 Enviar Correos", command=lanzar_envio, bg='green', fg='white').pack(pady=10)

    # === Área de logs ===
    scrollbar = Scrollbar(root)
    scrollbar.pack(side='right', fill='y')

    log_text = Text(root, height=20, yscrollcommand=scrollbar.set)
    log_text.pack(fill='both', expand=True, padx=10)
    scrollbar.config(command=log_text.yview)

    root.mainloop()

if __name__ == "__main__":
    main()
