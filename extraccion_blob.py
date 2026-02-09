import base64
import os
import re
import sys
import tempfile
from datetime import datetime, timedelta

import numpy as np
import oracledb
import pandas as pd
import pypandoc
import requests
import win32com.client
from docx import Document
from io import BytesIO
from json import loads
from sqlalchemy import create_engine, text

# Función para extraer texto de un documento .doc utilizando pywin32
def extract_text_from_doc(blob_data):
    try:
        if blob_data is None:
            return None  # En caso de valores nulos

        # Crear un archivo temporal con extensión .doc
        with tempfile.NamedTemporaryFile(delete=False, suffix=".doc") as temp_file:
            # Escribir los datos binarios del BLOB en el archivo temporal
            temp_file.write(blob_data)
            temp_file_path = temp_file.name  # Obtener la ruta del archivo temporal

        # Usar pywin32 para abrir el archivo .doc en Word y extraer el texto
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # No mostrar la interfaz de Word
        doc = word.Documents.Open(temp_file_path)
        text_content = doc.Content.Text  # Extraer todo el texto del documento
        doc.Close()
        word.Quit()

        # Eliminar el archivo temporal después de la conversión
        os.remove(temp_file_path)

        return text_content.strip()  # Limpiar espacios innecesarios
    except Exception as e:
        return f"Error: {str(e)}"


oracledb.version = "8.3.0"
sys.modules["cx_Oracle"] = oracledb

SERVER = "172.24.124.30"
USER = "****"
PWD = "****"
SERVICE_NAME = "*****"
ENGINE = create_engine(
    f"oracle://{USER}:{PWD}@(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)"
    f"(HOST={SERVER})(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME={SERVICE_NAME})))"
)


QUERY = text(
    """
 SELECT
    sw.PATIENT_PERSON_KEY,
    pil.PATIENT_ID AS CEDULA,
    r.REPORT_KEY AS ID_ESTUDIO_RIS,
    r.REPORT_CREATED_DATE AS FECHA_ESTUDIO,
    TRUNC(MONTHS_BETWEEN(SYSDATE, per.BIRTH_DATE) / 12) AS EDAD,
    rc.DESCRIPTION AS MODALIDAD,
    s.DESCRIPTION,
    r.DOCUMENT,
    r.DOCUMENT_PLAIN_TEXT
FROM SITE_WORKLIST sw
INNER JOIN SPS s
    ON sw.SPS_ID = s.SPS_ID
INNER JOIN PATIENT pat
    ON sw.PATIENT_PERSON_KEY = pat.PATIENT_PERSON_KEY
INNER JOIN PERSON per
    ON per.PERSON_KEY = pat.PATIENT_PERSON_KEY
INNER JOIN PATIENT_ID_LIST pil
    ON pil.PATIENT_PERSON_KEY = per.PERSON_KEY
INNER JOIN REPORT r
    ON sw.REPORT_KEY = r.REPORT_KEY
INNER JOIN RP_CODE rc
    ON sw.RP_CODE_KEY = rc.RP_CODE_KEY
WHERE
    r.REPORT_CREATED_DATE BETWEEN :start_date AND :end_date
    AND TRUNC(MONTHS_BETWEEN(SYSDATE, per.BIRTH_DATE) / 12) >= 18
    AND sw.RP_CODE_KEY IN ('12150', '12156', '12149', '12151', '12154')
    AND s.DESCRIPTION IN ('US - Renal')
"""
)


start_date = datetime(2014, 7, 1, 0, 0, 0)
final_end_date = datetime(2020, 12, 31, 23, 59, 59)
frames = []

current_start = start_date
while current_start <= final_end_date:
    next_start = current_start + pd.DateOffset(months=6)
    current_end = min(
        next_start.to_pydatetime() - timedelta(seconds=1),
        final_end_date,
    )

    print(
        f"Procesando bloque: {current_start:%Y-%m-%d}"
        f" -> {current_end:%Y-%m-%d}"
    )
    chunk_df = pd.read_sql_query(
        QUERY,
        ENGINE,
        params={
            "start_date": current_start,
            "end_date": current_end,
        },
    )
    frames.append(chunk_df)

    current_start = next_start.to_pydatetime()


df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
if not df.empty:
    df.drop_duplicates(subset=["CEDULA", "ID_ESTUDIO_RIS"], keep="last", inplace=True)

print(df.head(20))
