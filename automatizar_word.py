
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

# Cargar la plantilla
doc = DocxTemplate("plantilla.docx")

# Datos para la plantilla
nombre = "Christian Burns"
telefono = "(011) 42422600"
correo = "christianburns03@gmail.com"
fecha = datetime.today().strftime("%d/%m/%Y")

constantes = {'nombre': nombre, 'telefono': telefono, 'correo': correo, 'fecha': fecha}

df = pd.read_excel('Alumnos.xlsx')

for indice, fila in df.iterrows():
    contenido = {
        'nombre_alumno':fila["Nombre del Alumno"],
        'nota_mat': fila['mat'],
        'nota_fis': fila['fis'],
        'nota_qui': fila['qui']
    }
    contenido.update(constantes)
    
    doc.render(contenido)
    doc.save(f"Notas_de_{fila['Nombre del Alumno']}.docx")
    print(contenido)



# Renderizar y guardar el documento
doc.render(constantes)
doc.save(f"prueba.docx")