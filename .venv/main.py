import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import matplotlib as plt
import io


def create_report(template_path,data,chart_data=None):
    st.write("creando informe..")
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if f'{{{{{key}}}}}' in paragraph.text:
                st.write("Reemplazando ",key," con ", value ,"en el informe.")
            paragraph.text = paragraph.text.replace(f'{{{{{key}}}}}',str(value)) 
    
    output =io.BytesIO()
    doc.save(output)
    output.seek(0)
    st.write("reporte creado con exito.")
    return output
        
    
        

def main():
    st.title("generador de reportes")
    template_file = st.file_uploader("cargar plantilla word", type="docx")
    data_file = st.file_uploader("cargar datos", type=["xlsx","csv"])
    if template_file and data_file:
        st. success("archivos cargados correctamente")
        df = pd.read_csv(data_file) if data_file.name.endswith('.csv') else pd.read_excel(data_file)
        st.subheader("Datos cargados")
        st.dataframe(df)


        row_index = st.selectbox("seleccionar archivo para el informe", options=range(len(df)))
        selected_data = df.iloc[row_index].to_dict()
        st.write(selected_data)
        
        
    if st.button("Generar Informe"):
        output = create_report(template_file,selected_data)
        st.download_button("descargar informe",output,"nforme_generado.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
       

    

main()