import streamlit as st
import tempfile
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
import openpyxl
import os
from io import BytesIO

st.title("Represent")

uploaded_files = st.file_uploader(
    "Por favor, carregue os ficheiros pdf.",
    type=["pdf"],
    accept_multiple_files=True,
    key="uploader_pdf"
)

if uploaded_files:

    placeholder = st.empty()
    placeholder.info("⏳ A processar ficheiros...")

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

        for uploaded_file in uploaded_files:

            base_name = os.path.splitext(uploaded_file.name)[0]

            # Criar PDF temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.read())
                temp_pdf_path = temp_pdf.name

            temp_dir = os.path.dirname(temp_pdf_path)
            excel_entrada = os.path.join(temp_dir, base_name + ".xlsx")
            excel_saida = os.path.join(temp_dir, base_name + "_processado.xlsx")

            # Processamento
            ref_text, name_text = pdf_to_excel(temp_pdf_path, excel_entrada)
            inf_texto = [f"Ref: {ref_text}", name_text]

            trim_excel_before_marker(excel_entrada, excel_saida)
            add_images(temp_pdf_path, excel_saida, inf_texto)

            # Adicionar ao ZIP
            zip_file.write(excel_saida, os.path.basename(excel_saida))

            # Limpeza
            os.remove(excel_entrada)
            os.remove(excel_saida)
            os.remove(temp_pdf_path)

    placeholder.empty()
    st.success("Todos os ficheiros foram processados!")

    zip_buffer.seek(0)

    st.download_button(
        label="Descarregar todos os ficheiros",
        data=zip_buffer,
        file_name="ficheiros_processados.zip",
        mime="application/zip"
    )
