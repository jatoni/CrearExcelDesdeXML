import pandas as pd
import xml.etree.ElementTree as ET
import os
import pprint
import xmltodict
from openpyxl import load_workbook
from pathlib import Path
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
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
            name_file_base = os.path.basename(carpeta)

            for archivo in archivos:
                if not archivo.endswith('.xml'):
                    continue

                # comprobar antes si la hoja actual ya está llena
                name_file = f"{name_file_base} Hoja {hoja_num}"
                if len(dataBuilded.get(name_file, [])) >= max_alumnos_por_hoja:
                    hoja_num += 1
                    name_file = f"{name_file_base} Hoja {hoja_num}"

                ruta_archivo = os.path.join(ruta_carpeta, archivo)
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
                name_clave = diccionarioData["Carrera"]["@cveCarrera"]

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
                    "CLAVE_DE_INSTITUCION": clave_insitucion,
                    "FOLIO_DIGITAL": diccionarioData["Autenticacion"]["@folioDigital"],
                }

                dataBuilded.setdefault(name_file, []).append(registro)

    return dataBuilded

def agregar_hoja_nueva_excel(ruta_excel, dic):
    coutn = 0
    libro = load_workbook(ruta_excel)
    hoja_base = libro[libro.sheetnames[0]]  # Primera hoja como plantilla

    for name_file, info in dic.items():
        if name_file in libro.sheetnames:
            print(f"La hoja '{name_file}' ya existe. No se agregará hoja nueva.")
            continue

        # Copiar hoja base
        nueva_hoja = libro.copy_worksheet(hoja_base)
        nueva_hoja.title = name_file

        df_info = pd.DataFrame(info)

        # ---------- TABLA DE DATOS ----------
        start_row = 15  # fila donde empieza tu tabla en la plantilla
        for r_idx, row in enumerate(dataframe_to_rows(df_info, index=False, header=False), start=start_row):
            for c_idx, value in enumerate(row, start=1):
                nueva_hoja.cell(row=r_idx, column=c_idx, value=value)

        # ---------- FECHA ----------
        fecha_valor = ""
        if "FECHA_DE_EXPEDICION" in df_info.columns and not df_info["FECHA_DE_EXPEDICION"].empty:
            fecha_valor = str(df_info["FECHA_DE_EXPEDICION"].iloc[0])

        # ---------- HOJA ----------
        numero_hoja = int(re.search(r'\d+$', name_file).group())
        prefijo_paq = re.match(r'PAQ\. T-\d+', name_file).group()
        total_hojas = sum(1 for nombre in dic.keys() if nombre.startswith(prefijo_paq))

        # Colocar fecha y hoja en celdas fijas
        nueva_hoja["J8"] = fecha_valor
        nueva_hoja["J9"] = f"{numero_hoja}/{total_hojas}"

        # ---------- FIRMAS ----------
        responsable = " "
        if "CLAVE_DE_INSTITUCION" in df_info.columns:
            for clave, nombre in responsables_por_clave.items():
                if clave in df_info["CLAVE_DE_INSTITUCION"].astype(str).values:
                    responsable = nombre
                    break
        
        
        nueva_hoja.merge_cells("C51:D51")
        # Firma izquierda
        nueva_hoja["C51"] = responsable

        # Firma derecha
        nueva_hoja["I51"] = "Mtra. Dely Karolina Urbano Sanchez"
        nueva_hoja["I52"] = "RECTOR"

        print(f"Hoja '{name_file}' creada con fecha, hoja y firmas en posiciones exactas.")
        coutn += 1
    print(f"Total de hojas agregadas: {coutn}")
    libro.save(ruta_excel)

    
ruta = "/Users/juanantoniotorres/Documents/ProyectoSalvandoVidaJacqueline/DocuemntosTitulos/Libro de Control de Folios de Titulos y Grados Electrónicos 2025.xlsx"
registro_alumnos = readXMLAndBuildData()
dic = registro_alumnos
agregar_hoja_nueva_excel(ruta, dic)