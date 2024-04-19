import os
from docx import Document

def eliminar_lineas_especificas(doc_path, lineas_para_eliminar):
    doc = Document(doc_path)

    for p in doc.paragraphs:
        if any(linea in p.text for linea in lineas_para_eliminar):
            p.clear()

    doc.save(doc_path)

def procesar_todos_los_documentos(carpeta):
    lineas_para_eliminar = [
        ]
    
    for filename in os.listdir(carpeta):
        if filename.endswith(".docx"):
            doc_path = os.path.join(carpeta, filename)
            eliminar_lineas_especificas(doc_path, lineas_para_eliminar)
            print(f"Procesado: {filename}")

carpeta_de_documentos = "docx_directory"
procesar_todos_los_documentos(carpeta_de_documentos)
