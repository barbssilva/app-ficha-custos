import streamlit as st
import tempfile
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
import openpyxl
import os

st.title("Represent")

uploaded_file = st.file_uploader(
    "Por favor, carregue o ficheiro pdf.",
    type=["pdf"],
    key="uploader_pdf"
)

from fichas_custos_excel_represent import trim_excel_before_marker, pdf_to_excel
        
if uploaded_file is not None:
    base_name = os.path.splitext(uploaded_file.name)[0]
        
    # Criar ficheiro PDF temporário
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.read())
        temp_pdf_path = temp_pdf.name

    # Agora cria o excel_entrada e excel_saida no mesmo diretório do ficheiro temporário,
    # mas com nomes baseados no ficheiro original:
    temp_dir = os.path.dirname(temp_pdf_path)
    excel_entrada = os.path.join(temp_dir, base_name + ".xlsx")
    excel_saida = os.path.join(temp_dir,"tabela_custos.xlsx")


    placeholder = st.empty()
    placeholder.info("⏳ Por favor aguarde...")

    ref_text, name_text = pdf_to_excel(temp_pdf_path,excel_entrada)
    inf_texto = [f"Ref: {ref_text}",name_text]

    trim_excel_before_marker(excel_entrada,excel_saida)
    add_images(temp_pdf_path,excel_saida,inf_texto)

    placeholder.empty()
    st.success("Processo terminado!")
        
    # Abrir o ficheiro Excel processado para download
    with open(excel_saida, "rb") as f:
        st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(excel_saida))
        
    #Remover o primeiro ficheiro excel criado uma vez que já não será utilizado
    os.remove(excel_entrada)
