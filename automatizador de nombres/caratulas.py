import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

# Cargar el archivo Excel
archivo = r'G:\Otros ordenadores\Mi portátil\UDESA\ECO I\Otoño 2024\Alumnos E010 Economía I Grupo 13 - Teórica 1.xlsx'
df = pd.read_excel(archivo)

# Asegurarse de que no haya nulos en Apellido o Nombre
df = df.dropna(subset=['Apellido', 'Nombre'])

# Crear columna con el nombre completo
df['Nombre completo'] = df['Apellido'].str.strip() + ', ' + df['Nombre'].str.strip()

# Revisión rápida de columnas disponibles
print(df.columns)

# Crear un único documento para todas las carátulas
doc = Document()


for i, nombre in enumerate(df['Nombre completo'], 1):
    # Título
    titulo = doc.add_heading('Examen Parcial - Economía I', level=1)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Subtítulo
    subtitulo2 = doc.add_paragraph('Universidad de San Andrés')
    subtitulo2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph('')

    # Nombre del alumno
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'Estudiante: {nombre}')
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph('')
    doc.add_paragraph('Grupo: 13')
    doc.add_paragraph(f'Página: {i}')

    # Salto de página (excepto al final)
    if i < len(df):
        doc.add_page_break()

# Guardar el documento completo
doc.save('Caratulas_Grupo13.docx')

print("Documento único con todas las carátulas generado correctamente.")