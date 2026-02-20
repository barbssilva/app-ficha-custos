import streamlit as st
import tempfile
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
import openpyxl
import os
from io import BytesIO
import zipfile

st.title("Represent")

from fichas_custos_excel_represent import trim_excel_before_marker, pdf_to_excel, add_images

st.write("Pode carregar vários ficheiros pdf de uma só vez, no fim fará download de um ficheiro .zip com todos os ficheiros excel")

uploaded_files = st.file_uploader(
    "Por favor, carregue os ficheiros pdf.",
    type=["pdf"],
    accept_multiple_files=True,
    key="uploader_pdf"
)

if uploaded_files:
    if len(uploaded_files) == 1:
        for uploaded_file in uploaded_files:
            base_name = os.path.splitext(uploaded_file.name)[0]
                
            # Criar ficheiro PDF temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.read())
                temp_pdf_path = temp_pdf.name
        
            # Agora cria o excel_entrada e excel_saida no mesmo diretório do ficheiro temporário,
            # mas com nomes baseados no ficheiro original:
            temp_dir = os.path.dirname(temp_pdf_path)
            excel_entrada = os.path.join(temp_dir, base_name + ".xlsx")
            #excel_saida = os.path.join(temp_dir,"tabela_custos.xlsx")
        
        
            placeholder = st.empty()
            placeholder.info("⏳ Por favor aguarde...")
        
            ref_text, name_text = pdf_to_excel(temp_pdf_path,excel_entrada)
            inf_texto = [f"Ref: {ref_text}",name_text]
            excel_saida = os.path.join(temp_dir,f"tabela_custos_{ref_text}.xlsx")
        
            trim_excel_before_marker(excel_entrada,excel_saida)
            add_images(temp_pdf_path,excel_saida,inf_texto)
        
            placeholder.empty()
            st.success("Processo terminado!")
                
            # Abrir o ficheiro Excel processado para download
            with open(excel_saida, "rb") as f:
                st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(excel_saida))
                
            #Remover o primeiro ficheiro excel criado uma vez que já não será utilizado
            os.remove(excel_entrada)

    else:
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
                #excel_saida = os.path.join(temp_dir, base_name + "_processado.xlsx")
    
                # Processamento
                ref_text, name_text = pdf_to_excel(temp_pdf_path, excel_entrada)
                inf_texto = [f"Ref: {ref_text}", name_text]
                excel_saida = os.path.join(temp_dir,f"tabela_custos_{ref_text}.xlsx")
    
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
