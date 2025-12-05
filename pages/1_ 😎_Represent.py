import streamlit as st

st.title("Represent")

ficha_custo = st.file_uploader(
    "Por favor, carregue o ficheiro pdf.",
    type=["pdf"],
    key="uploader_pdf"
)


