import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import matplotlib as plt
import io


def create_report(template_path,data,rep1,chart_data=None):
    st.write("creando informe..")
    st.write(data)
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if f'{{{{{'Contenido'}}}}}' in paragraph.text:
                paragraph.text = paragraph.text.replace(f'{{{{{key}}}}}',str(rep1)) 
            if f'{{{{{key}}}}}' in paragraph.text:
                st.write("Reemplazando ",key," con ", value ,"en el informe.")
            paragraph.text = paragraph.text.replace(f'{{{{{key}}}}}',str(value)) 


    
    output =io.BytesIO()
    doc.save(output)
    output.seek(0)
    st.write("Reporte creado con exito.")
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
        result=df.groupby('Folio').agg(
            Contenido=('Contenido','sum')
        )
        
        st.write('informacion del reporte')
        selected_data = df.iloc[0].to_dict()
        folio =df.iloc[0,0]
 

    if st.button("Generar Informe"):
        rep1=result.iloc[0,0]
        output = create_report(template_file,selected_data,rep1)
        st.download_button("descargar informe",output,"Reporte-"+folio.astype(str)+"-.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
       

main()