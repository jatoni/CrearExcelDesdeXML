import pandas as pd
import xml.etree.ElementTree as ET
import os
import pprint
import xmltodict
from openpyxl import load_workbook
from pathlib import Path
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import re

responsables_por_clave = {
    "150901": "Mtra. Beatriz Adriana Barron Linares",
    "010162": "Mtra. Lilibeth Hernandez Alva",
    "010157": "Mtra. Lilibeth Hernandez Alva"  
}

def readXMLAndBuildData():
    dataBuilded = {}
    base_path = './ArchivosXML'
    carpetas = os.listdir(base_path)

    max_alumnos_por_hoja = 26

    for carpeta in carpetas:
        ruta_carpeta = os.path.join(base_path, carpeta)
        if os.path.isdir(ruta_carpeta):
            archivos = sorted(os.listdir(ruta_carpeta))
            hoja_num = 1  # empezamos en hoja 1
            for archivo in archivos:
                if archivo.endswith('.xml'):
                    # nombre base sin hoja
                    name_file_base = os.path.basename(carpeta)
                    # nombre con hoja
                    name_file = f"{name_file_base} Hoja {hoja_num}"

                    ruta_archivo = os.path.join(ruta_carpeta, archivo)
                    name_clave = os.path.basename(archivo).split('_')[2]
                    folio_Control = os.path.splitext(os.path.basename(archivo))[0]

                    with open(ruta_archivo, 'r', encoding='utf-8') as f:
                        contenido_xml = f.read()

                    diccionarioData = xmltodict.parse(contenido_xml)["TituloElectronico"]

                    nombreAlumno = (
                        diccionarioData["Profesionista"]["@nombre"] + " " +
                        diccionarioData["Profesionista"]["@primerApellido"] + " " +
                        diccionarioData["Profesionista"]["@segundoApellido"]
                    )
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

                    # si ya está llena la hoja, pasamos a la siguiente
                    if len(dataBuilded.get(name_file, [])) >= max_alumnos_por_hoja:
                        hoja_num += 1
                        name_file = f"{name_file_base} Hoja {hoja_num}"

                    dataBuilded.setdefault(name_file, []).append(registro)

    return dataBuilded



def agregar_hoja_nueva_excel(ruta_excel, dic):
    try:
        libro = load_workbook(ruta_excel)
    except FileNotFoundError:
        libro = Workbook()

    for name_file, info in dic.items():
        if name_file in libro.sheetnames:
            print(f"La hoja '{name_file}' ya existe. No se agregará hoja nueva.")
            continue

        ws = libro.create_sheet(name_file)

        # ----------------- ENCABEZADOS -----------------
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

        # Reducimos el merge para dejar espacio a FECHA y HOJA
        ws.merge_cells("A6:J6")
        ws["A6"] = "LIBRO DE CONTROL DE FOLIOS DE TÍTULOS ELECTRÓNICOS"
        ws["A6"].alignment = Alignment(horizontal="center")
        ws["A6"].font = Font(size=12, bold=True)

        # ----------------- FECHA Y HOJA -----------------
        df_info = pd.DataFrame(info)
        fecha_valor = ""
        if "FECHA_DE_EXPEDICION" in df_info.columns and not df_info["FECHA_DE_EXPEDICION"].empty:
            fecha_valor = str(df_info["FECHA_DE_EXPEDICION"].iloc[0])
        
        # Extraer número de hoja desde name_file
        numero_hoja = int(re.search(r'\d+$', name_file).group())

        # Calcular cuántas hojas existen en el paquete actual
        prefijo_paq = re.match(r'PAQ\. T-\d+', name_file).group()  # Ej: 'PAQ. T-214627'
        total_hojas = sum(1 for nombre in dic.keys() if nombre.startswith(prefijo_paq))

        ws.cell(row=6, column=11, value="FECHA").alignment = Alignment(horizontal="center")
        ws.cell(row=6, column=12, value=fecha_valor).alignment = Alignment(horizontal="center")
        ws.cell(row=7, column=11, value="HOJA").alignment = Alignment(horizontal="center")
        ws.cell(row=7, column=12, value=f"{numero_hoja}/{total_hojas}").alignment = Alignment(horizontal="center")

        # ----------------- TABLA DE DATOS -----------------
        start_row = 14
        for r_idx, row in enumerate(dataframe_to_rows(df_info, index=False, header=True), start=start_row):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Bordes
        thin = Side(border_style="thin", color="000000")
        for row in ws.iter_rows(min_row=start_row, max_row=start_row + len(df_info),
                                min_col=1, max_col=len(df_info.columns)):
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # ----------------- RESPONSABLE -----------------
        responsable = " " 
        if "CLAVE_DE_CARRERA" in df_info.columns:
            for clave, nombre in responsables_por_clave.items():
                if clave in df_info["CLAVE_DE_CARRERA"].astype(str).values:
                    responsable = nombre
                    break

        # ----------------- FIRMAS -----------------
        fila_firmas = start_row + len(df_info) + 10
        ws.merge_cells(start_row=fila_firmas, start_column=2, end_row=fila_firmas, end_column=5)
        ws.cell(row=fila_firmas, column=2, value=responsable).alignment = Alignment(horizontal="center")

        ws.merge_cells(start_row=fila_firmas + 1, start_column=2, end_row=fila_firmas + 1, end_column=5)
        ws.cell(row=fila_firmas + 1, column=2, value="RESPONSABLE DE SERVICIOS ESCOLARES").alignment = Alignment(horizontal="center")

        ws.merge_cells(start_row=fila_firmas, start_column=9, end_row=fila_firmas, end_column=11)
        ws.cell(row=fila_firmas, column=9, value="Mtra. Dely Karolina Urbano Sanchez").alignment = Alignment(horizontal="center")

        ws.merge_cells(start_row=fila_firmas + 1, start_column=9, end_row=fila_firmas + 1, end_column=11)
        ws.cell(row=fila_firmas + 1, column=9, value="RECTOR").alignment = Alignment(horizontal="center")

        print(f"Hoja '{name_file}' agregada con formato, fecha y firmas.")

    libro.save(ruta_excel)
    
ruta = "/Users/juanantoniotorres/Documents/ProyectoSalvandoVidaJacqueline/DocuemntosTitulos/Libro de Control de Folios de Titulos y Grados Electrónicos 2025.xlsx"
registro_alumnos = readXMLAndBuildData()
dic = registro_alumnos
agregar_hoja_nueva_excel(ruta, dic)           