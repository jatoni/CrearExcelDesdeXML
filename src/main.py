import pandas as pd
import xml.etree.ElementTree as ET
import os
import pprint
import xmltodict
from openpyxl import load_workbook
from pathlib import Path
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows

responsables_por_clave = {
    "150901": "Mtra. Beatriz Adriana Barron Linares",
    "010162": "Mtra. Lilibeth Hernandez Alva",
    "010157": "Mtra. Lilibeth Hernandez Alva"  
}


def readXMLAndBuildData():
    dataBuilded = {}
    base_path = './ArchivosXML'
    carpetas = os.listdir(base_path)
    for carpeta in carpetas:
        name_file = os.path.basename(carpeta) + " Hoja 1"
        ruta_carpeta = os.path.join(base_path, carpeta)
        if os.path.isdir(ruta_carpeta):
            archivos = sorted(os.listdir(ruta_carpeta))
            for archivo in archivos:
                if archivo.endswith('.xml'):
                    ruta_archivo = os.path.join(ruta_carpeta, archivo)
                    name_clave = os.path.basename(archivo).split('_')[2]
                    folio_Control = os.path.splitext(os.path.basename(archivo))[0]
                    with open(ruta_archivo, 'r', encoding='utf-8') as f:
                        contenido_xml = f.read()
                    diccionarioData = xmltodict.parse(contenido_xml)["TituloElectronico"]
                    nombreAlumno = diccionarioData["Profesionista"]["@nombre"] + " " +diccionarioData["Profesionista"]["@primerApellido"] + " " +diccionarioData["Profesionista"]["@segundoApellido"]
                    curp = diccionarioData["Profesionista"]["@curp"]
                    programa = diccionarioData["Carrera"]["@nombreCarrera"]
                    lugar_expedicion = diccionarioData["Expedicion"]["@entidadFederativa"]
                    fecha_expedicion = diccionarioData["Expedicion"]["@fechaExpedicion"]
                    rvoe = diccionarioData["Carrera"]["@numeroRvoe"]
                    clave_insitucion = diccionarioData["Institucion"]["@cveInstitucion"]
                    registro = {
                        "NUM_PROG": len(dataBuilded.get(name_file, [])) + 1,
                        "ALUMNO": nombreAlumno,
                        "CURP": curp,
                        "PROGRAMA": programa,
                        "CLAVE_DE_CARRERA": name_clave,
                        "FOLIO_DE_CONTROL": folio_Control,
                        "LUGAR_DE_EXPEDICION": lugar_expedicion,
                        "FECHA_DE_EXPEDICION": fecha_expedicion,
                        "RVOE": rvoe,
                        "CLAVE_DE_INSTITUCION": clave_insitucion
                    }
                    dataBuilded.setdefault(name_file, []).append(registro)
                    
    return dataBuilded


def agregar_hoja_nueva_excel(ruta_excel, dic):
    df = pd.DataFrame(registro_alumnos["PAQ. T-222042 Hoja 1"])
    try:
        libro = load_workbook(ruta_excel)
    except FileNotFoundError:
        libro = Workbook()

    for name_file, info in dic.items():
        if name_file in libro.sheetnames:
            print(f"La hoja '{name_file}' ya existe. No se agregará hoja nueva.")
            continue

        ws = libro.create_sheet(name_file)

        # Encabezados principales
        ws.merge_cells("A1:L1")
        ws["A1"] = "UNIVERSIDAD ETAC ON ALIAT"
        ws["A1"].alignment = Alignment(horizontal="center")
        ws["A1"].font = Font(size=14, bold=True)

        ws.merge_cells("A2:L2")
        ws["A2"] = "CAMPUS ÚNICO"
        ws["A2"].alignment = Alignment(horizontal="center")
        ws["A2"].font = Font(size=12, bold=True)

        ws.merge_cells("A4:L4")
        ws["A4"] = "SERVICIOS ESCOLARES"
        ws["A4"].alignment = Alignment(horizontal="center")
        ws["A4"].font = Font(size=12, bold=True)

        ws.merge_cells("A6:L6")
        ws["A6"] = "LIBRO DE CONTROL DE FOLIOS DE TÍTULOS ELECTRÓNICOS"
        ws["A6"].alignment = Alignment(horizontal="center")
        ws["A6"].font = Font(size=12, bold=True)

        # Escribir datos
        df = pd.DataFrame(info)
        start_row = 14
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=start_row):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Aplicar bordes
        thin = Side(border_style="thin", color="000000")
        for row in ws.iter_rows(min_row=start_row, max_row=start_row + len(df),
                                min_col=1, max_col=len(df.columns)):
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # Determinar responsable según clave de carrera
        responsable = ""
        if "CLAVE DE CARRERA" in df.columns:
            for clave, nombre in responsables_por_clave.items():
                if clave in df["CLAVE DE CARRERA"].astype(str).values:
                    responsable = nombre
                    break

        fila_firmas = start_row + len(df) + 10
        ws.merge_cells(start_row=fila_firmas, start_column=2, end_row=fila_firmas, end_column=5)
        ws.cell(row=fila_firmas, column=2, value=responsable).alignment = Alignment(horizontal="center")
        
        ws.merge_cells(start_row=fila_firmas+1, start_column=2, end_row=fila_firmas+1, end_column=5)
        ws.cell(row=fila_firmas+1, column=2, value="RESPONSABLE DE SERVICIOS ESCOLARES").alignment = Alignment(horizontal="center")

        ws.merge_cells(start_row=fila_firmas, start_column=9, end_row=fila_firmas, end_column=11)
        ws.cell(row=fila_firmas, column=9, value="Mtra. Dely Karolina Urbano Sanchez").alignment = Alignment(horizontal="center")

        ws.merge_cells(start_row=fila_firmas+1, start_column=9, end_row=fila_firmas+1, end_column=11)
        ws.cell(row=fila_firmas+1, column=9, value="RECTOR").alignment = Alignment(horizontal="center")

        print(f"Hoja '{name_file}' agregada con formato y firmas.")

    libro.save(ruta_excel)

ruta = "/Users/juanantoniotorres/Documents/ProyectoSalvandoVidaJacqueline/DocuemntosTitulos/Libro de Control de Folios de Titulos y Grados Electrónicos 2025.xlsx"
registro_alumnos = readXMLAndBuildData()
dic = registro_alumnos
agregar_hoja_nueva_excel(ruta, dic)           