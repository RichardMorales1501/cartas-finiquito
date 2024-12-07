import tkinter as tk
from tkinter import *
from tkinter import messagebox, filedialog
import pandas as pd
from fpdf import FPDF
from datetime import datetime

def cambia_color(event):
    event.widget.configure(bg="gray30", fg="dodger blue")

def regresa_color(event):
    event.widget.configure(bg="SystemButtonFace", fg="black")

# Seleccionar archivo Excel
def ligacontactos():
    contactos = filedialog.askopenfilename(title="Abrir archivo", filetypes=[("Archivo Excel", "*.xlsx *.xls")])
    txt9.delete(0, 'end')
    txt9.insert(0, contactos)

# Generar PDFs
def grabar():
    try:
        basecontactos = txt9.get()
        cartas = pd.read_excel(basecontactos)  # Leer archivo Excel
        
        # Crear PDFs para cada fila en el archivo Excel
        for index, row in cartas.iterrows():
            folio = row['Folio']
            suscriptor = row['Suscriptor']
            fecha_carta = row['Fecha de carta']
            contrato = row['Contrato']
            fecha_otorgamiento = row['Fecha de otorgamiento']
            nombre_alumno = row['Nombre Alumno']
            nombre_archivo = row['Nombre del archivo']
            
            # Crear PDF
            pdf = FPDF()
            pdf.add_page()
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
            texto = (f"Suscriptor: {suscriptor}\n\n"
                     f"Por medio del presente, CORPORATIVO LAUDEX S.A.P.I DE C.V SOFOM E.N.R.,\n"
                     f"hace de su conocimiento la liquidación total del contrato {contrato} que fue otorgado\n"
                     f"el {fecha_otorgamiento} al ciudadano {nombre_alumno} no ejerciendo\n"
                     f"alguna responsabilidad, acción y/o derecho entre ambas partes sea de carácter civil,\n"
                     f"mercantil u otro medio legal, para todos los efectos legales a que haya lugar.\n\n"
                     f"Así mismo se enviará constancia de su comportamiento a las sociedades de información\n"
                     f"crediticia que corresponda.\n\n"
                     f"Se extiende la presente a solicitud del acreditado señalado, con fines informativos y sin\n"
                     f"responsabilidad alguna para CORPORATIVO LAUDEX S.A.P.I DE C.V SOFOM E.N.R.")
            
            pdf.multi_cell(0, 10, txt=texto)
            pdf.ln(20)

            # Firma
            pdf.cell(0, 10, txt="___________________________________", ln=True, align='C')
            pdf.cell(0, 10, txt="Ana Lucia Carbajal", ln=True, align='C')
            pdf.cell(0, 10, txt="Gerente de Atención a Clientes", ln=True, align='C')

            # Guardar el archivo PDF
            pdf_output = f"{nombre_archivo}.pdf"
            pdf.output(pdf_output)
            print(f"PDF generado: {pdf_output}")
        
        messagebox.showinfo("Éxito", "Los PDFs han sido generados correctamente.")
    
    except Exception as e:
        print(f"Error: {e}")
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# Interfaz gráfica
vent = Tk()
vent.title("Generador de Cartas Finiquito")
vent.geometry("1000x550")

# Botón para seleccionar archivo Excel
botonexcel = Button(vent, text="Seleccionar archivo Excel", command=ligacontactos, font=("Helvetica Neue", 12))
botonexcel.place(relx=0.35, rely=0.03)
botonexcel.bind("<Enter>", cambia_color)
botonexcel.bind("<Leave>", regresa_color)

# Campo de texto para mostrar la ruta del archivo
txt9 = Entry(vent, bg="ivory3", fg="blue4", font=("Helvetica Neue", 12))
txt9.place(relx=0.50, rely=0.03, relwidth=0.45, relheight=0.05)

# Botón para generar cartas
botonenviar = Button(vent, text="Generar Cartas", command=grabar, font=("Helvetica Neue", 12))
botonenviar.place(relx=0.55, rely=0.93)
botonenviar.bind("<Enter>", cambia_color)
botonenviar.bind("<Leave>", regresa_color)

vent.mainloop()