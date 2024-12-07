import tkinter as tk
from tkinter import *
from tkinter import messagebox, filedialog
import pandas as pd
from fpdf import FPDF
from datetime import datetime
from PIL import Image, ImageTk
import os


def cambia_color(event):
    event.widget.configure(bg="gray30", fg="dodger blue")

def regresa_color(event):
    event.widget.configure(bg="SystemButtonFace", fg="black")

# Seleccionar archivo Excel
def ligacontactos():
    contactos = filedialog.askopenfilename(title="Abrir archivo", filetypes=[("Archivo Excel", "*.xlsx *.xls")])
    txt9.delete(0, 'end')
    txt9.insert(0, contactos)

# Seleccionar carpeta para guardar
def elegir_directorio():
    directorio = filedialog.askdirectory(title="Selecciona una carpeta para guardar los archivos PDF")
    txt10.delete(0, 'end')
    txt10.insert(0, directorio)

# Generar PDFs
def grabar():
    try:
        basecontactos = txt9.get()
        carpeta_salida = txt10.get()

        if not basecontactos:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo Excel.")
            return
        if not carpeta_salida:
            messagebox.showwarning("Advertencia", "Por favor, selecciona una carpeta para guardar los PDFs.")
            return

        cartas = pd.read_excel(basecontactos)  # Leer archivo Excel

        # Crear PDFs para cada fila en el archivo Excel
        for index, row in cartas.iterrows():
            folio = row['Folio']
            fecha_carta = row['Fecha de carta']
            contrato = row['Contrato']
            fecha_otorgamiento = row['Fecha de activacion']
            nombre_alumno = row['Nombre Alumno']
            nombre_archivo = row['Nombre del archivo']

            try:
                # Intentar convertir la fecha automáticamente
                if isinstance(row['Fecha de carta'], str):  # Si la fecha es una cadena
                    fecha_carta_obj = datetime.strptime(row['Fecha de carta'], '%Y-%m-%d %H:%M:%S')  # Formato común
                else:  # Si es un objeto de fecha
                    fecha_carta_obj = pd.to_datetime(row['Fecha de carta']).to_pydatetime()

                # Formatear la fecha a DD-MM-AAAA
                fecha_carta = fecha_carta_obj.strftime('%d-%m-%Y')
            except Exception as e:
                print(f"Error al convertir la fecha: {e}")
                fecha_carta = "Fecha inválida"

            # Crear PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font('Arial', size=12)

            # Agregar logo
            logo_path = "/Users/richardmacairm1/Documents/Laudex/Cartas Finiquito/logo.png"
            pdf.image(logo_path, x=10, y=8, w=50)  # Ajustar tamaño y posición del logo
            pdf.set_font('Arial', size=12)

            # Agregar folio y fecha
            pdf.cell(0, 10, txt=f"Folio: {folio}", ln=True, align='R')
            pdf.ln(5)
            pdf.cell(0, 10, txt=f"Ciudad de México a {fecha_carta}", ln=True, align='R')
            pdf.ln(10)

            # Título de la carta
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, txt="CARTA FINIQUITO", ln=True, align='C')
            pdf.ln(10)

            # Contenido de la carta
            pdf.set_font('Arial', size=12)
            texto = (f"Suscriptor: {nombre_alumno}\n\n"
                     f"Por medio del presente, CORPORATIVO LAUDEX S.A.P.I DE C.V SOFOM E.N.R., hace de su conocimiento la liquidación total del contrato {contrato} que fue otorgado el {fecha_otorgamiento} al ciudadano {nombre_alumno} no ejerciendo alguna responsabilidad, acción y/o derecho entre ambas partes sea de carácter civil, mercantil u otro medio legal, para todos los efectos legales a que haya lugar.\n"
                     f"Así mismo se enviará constancia de su comportamiento a las sociedades de información crediticia que corresponda.\n"
                     f"Se extiende la presente a solicitud del acreditado señalado, con fines informativos y sin responsabilidad alguna para CORPORATIVO LAUDEX S.A.P.I DE C.V SOFOM E.N.R.\n")
            
            pdf.multi_cell(0, 10, txt=texto)
            pdf.ln(20)

                        # Firma y nombre
            firma_path = "/Users/richardmacairm1/Documents/Laudex/Cartas Finiquito/firma.png"  # Ruta de la firma
            pdf.image(firma_path, x=80, y=pdf.get_y(), w=50)  # Agregar firma
            pdf.ln(25)  # Espacio para la firma

            # Firma
            pdf.cell(0, 10, txt="___________________________________", ln=True, align='C')
            pdf.cell(0, 10, txt="Ana Lucia Carbajal", ln=True, align='C')
            pdf.cell(0, 10, txt="Gerente de Atención a Clientes", ln=True, align='C')

            # Guardar el archivo PDF
            pdf_output = os.path.join(carpeta_salida, f"{nombre_archivo}.pdf")
            pdf.output(pdf_output)
            print(f"PDF generado: {pdf_output}")

        messagebox.showinfo("Éxito", "Los PDFs han sido generados correctamente.")
    
    except Exception as e:
        print(f"Error: {e}")
        messagebox.showerror("Error", f"Ocurrió un error: {e}")


# Interfaz gráfica
vent = Tk()
vent.title("Generador de cartas")
vent.geometry("700x210") 
# Cargar y redimensionar la imagen
try:
    img = Image.open("/Users/richardmacairm1/Documents/Laudex/Cartas Finiquito/Laudex.png")
    img = img.resize((200, 200))  # Cambia las dimensiones a 200x200 píxeles
    imag = ImageTk.PhotoImage(img)
    
    # Crear el label con la imagen redimensionada
    label = Label(vent, image=imag)
    label.image = imag  # Mantener la referencia de la imagen
    label.place(relx=0.01, rely=0.01)
except Exception as e:
    print(f"No se pudo cargar la imagen: {e}")

# Botón para seleccionar archivo Excel
botonexcel = Button(vent, text="Seleccionar archivo Excel", command=ligacontactos, font=("Helvetica Neue", 14))
botonexcel.place(relx=0.35, rely=0.03, relwidth=.60, relheight=0.15)
botonexcel.bind("<Enter>", cambia_color)
botonexcel.bind("<Leave>", regresa_color)

# Campo de texto para mostrar la ruta del archivo
txt9 = Entry(vent, bg="ivory3", fg="blue4", font=("Helvetica Neue", 12))
txt9.place(relx=0.35, rely=0.18, relwidth=.60, relheight=0.15)

# Botón para seleccionar archivo Excel
botondownload = Button(vent, text="Seleccionar donde guardar cartas", command=elegir_directorio, font=("Helvetica Neue", 14))
botondownload.place(relx=0.35, rely=0.48, relwidth=.60, relheight=0.15)
botondownload.bind("<Enter>", cambia_color)
botondownload.bind("<Leave>", regresa_color)

# Campo de texto para mostrar la ruta del archivo
txt10 = Entry(vent, bg="ivory3", fg="blue4", font=("Helvetica Neue", 12))
txt10.place(relx=0.35, rely=0.63, relwidth=.60, relheight=0.15)

# Botón para generar cartas
botonenviar = Button(vent, text="Generar Cartas", command=grabar, font=("Helvetica Neue", 14))
botonenviar.place(relx=0.55, rely=0.80)
botonenviar.bind("<Enter>", cambia_color)
botonenviar.bind("<Leave>", regresa_color)

vent.mainloop()