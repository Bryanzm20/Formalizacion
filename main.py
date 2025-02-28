import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from datetime import time, datetime
import os

# Cargar el DataFrame desde el archivo Excel con manejo de errores
try:
    df_General = pd.read_excel(r"Controldepesos.xlsx")
    area_deseada = "Formalizacion"
    df = df_General.loc[df_General['Area'] == area_deseada]
    # Asegurar que la columna 'Fecha' sea de tipo datetime
    df['Fecha'] = pd.to_datetime(df['Fecha']).dt.date
except FileNotFoundError:
    st.error("El archivo Controldepesos.xlsx no se encontró.")
    st.stop()
except Exception as e:
    st.error(f"Ocurrió un error al leer el archivo Excel: {e}")
    st.stop()

st.title('Análisis de Datos y Generación de PDF')

# Definición de la función generar_pdf con manejo de errores
def generar_pdf(df_pdf, fecha_form, material_form, encargado_form):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5 * inch)
    styles = getSampleStyleSheet()
    elements = []

    try:
        # Encabezado con manejo de errores para la imagen
        imagen_path = "Img/logozcnl.png"
        if os.path.exists(imagen_path):
            imagen = Image(imagen_path, width=2.5 * inch, height=0.8 * inch)
        else:
            st.error(f"No se encontró la imagen en la ruta: {imagen_path}")
            st.stop()

        table_header = Table([[imagen,
                                Paragraph("REGISTRO DE PESAJE DE MATERIAL MINERALIZADO<br/>RECIBIDO DE LAS FORMALIZACIONES", styles['Heading3'])]],
                                colWidths=[3 * inch, 5 * inch])
        table_header.setStyle(TableStyle([('ALIGN', (0, 0), (0, 0), 'LEFT'),
                                            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                            ('LEFTPADDING', (1, 0), (1, 0), 10)]))
        elements.append(table_header)
        elements.append(Spacer(1, 0.2 * inch))

        # Datos del formulario y calculados
        hora_inicio = df_pdf['Hora'].min().strftime('%H:%M')
        hora_fin = df_pdf['Hora'].max().strftime('%H:%M')

        datos_tabla = [
            ["Fecha (dd/mm/aa):", fecha_form.strftime('%d/%m/%Y'), "Material:", material_form],
            ["Encargado Pesaje FM:", encargado_form, "Frente (s) de explotación:", "N/A"],
            ["Hora inicio:", hora_inicio, "Lugar del pesaje:", "Higabra"],
            ["Hora fin:", hora_fin, Paragraph("Lugar de recepción y<br/> muestreo:"), "Platanal"]
        ]

        # Calcular el ancho de las columnas para la tabla de datos
        num_cols_datos = len(datos_tabla[0])
        ancho_columna_datos = 7.5 * inch / num_cols_datos

        tabla_datos = Table(datos_tabla, colWidths=[ancho_columna_datos] * num_cols_datos)
        tabla_datos.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                                            ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
        elements.append(tabla_datos)
        elements.append(Spacer(1, 0.2 * inch))

        # Nueva tabla con datos de peso de camiones
        total_tara = df_pdf['Peso Tara (Kg)'].sum() / 1000
        total_bruto = df_pdf['Peso Bruto (Kg)'].sum() / 1000
        total_neto = df_pdf['Peso Neto (Kg)'].sum() / 1000

        pesos_tabla = [
            [Paragraph("Total peso camiones vacíos (Ton):", styles['Normal']), f"{total_tara:.2f}",
             Paragraph("Total peso camiones cargados (Ton):", styles['Normal']), f"{total_bruto:.2f}",
             Paragraph("Peso total del material pesado (Ton):", styles['Normal']), f"{total_neto:.2f}"]
        ]

        # Calcular el ancho de las columnas para la tabla de pesos
        num_cols_pesos = len(pesos_tabla[0])
        ancho_columna_pesos = 7.5 * inch / num_cols_pesos

        tabla_pesos = Table(pesos_tabla, colWidths=[ancho_columna_pesos] * num_cols_pesos)
        tabla_pesos.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                                            ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                            ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                                            ('ALIGN', (1, 0), (-1, 0), 'CENTER')]))
        elements.append(tabla_pesos)
        elements.append(Spacer(1, 0.2 * inch))

        # inicio tablas firmas
        num_cols_firma = 4
        ancho_columna_firma = 7.5 * inch / num_cols_firma

        firma_tabla = Table([
            [Paragraph("Nombre de quien<br/> entrega<br/>(Empresa FM)", styles['Normal']), "", "Firma de quien entrega", ""]
        ], colWidths=[ancho_columna_firma] * num_cols_firma)
        firma_tabla.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (0, 1), (-1, 1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        elements.append(firma_tabla)
        elements.append(Spacer(1, 0.2 * inch))

        firma_tabla_recibe_formalizacion = Table([
            [Paragraph("Nombre de quien recibe<br/>(Formalizacion)", styles['Normal']), "", "Firma de quien recibe", ""]
        ], colWidths=[ancho_columna_firma] * num_cols_firma)
        firma_tabla_recibe_formalizacion.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (0, 1), (-1, 1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        elements.append(firma_tabla_recibe_formalizacion)
        elements.append(Spacer(1, 0.2 * inch))

        firma_tabla_recibe_operaciones = Table([
            [Paragraph("Nombre de quien recibe<br/>(Operaciones)", styles['Normal']), "", "Firma de quien recibe", ""]
        ], colWidths=[ancho_columna_firma] * num_cols_firma)
        firma_tabla_recibe_operaciones.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (0, 1), (-1, 1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        elements.append(firma_tabla_recibe_operaciones)
        elements.append(Spacer(1, 0.2 * inch))

        data_adicional = {
            'Hora pesaje': ['9:34', '9:20', '', '', '', '', 'TOTAL'],
            'Placa': ['JYN047', 'LKL226', '', '', '', '', ''],
            'Peso volqueta vacia': [11520, 11500, '', '', '', '', 23020],'Peso volqueta llena': [38420, 37980, '', '', '', '', 76400],
            'Total': [26900, 26480, 0, 0, 0, 0, 53380]
        }
        df_adicional = pd.DataFrame(data_adicional)
        data_adicional_pdf = [df_adicional.columns.tolist()] + df_adicional.values.tolist()
        num_cols_df_adicional = len(df_adicional.columns)
        ancho_columna_df_adicional = 7.5 * inch / num_cols_df_adicional

        table_adicional = Table(data_adicional_pdf, colWidths=[ancho_columna_df_adicional] * num_cols_df_adicional)
        table_adicional.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table_adicional)

        # Sección de observaciones
        observaciones_text = """Observaciones:

Para recepción de mineral los días sábados, domingos y festivos:

• Se tomará el precio del Au según la Bolsa de Metales de Londres en su versión p.m. correspondiente al ultimo día hábil previo a la entrega de mineral, tomando así para las entregas los días sábados y domingos el precio del Au correspondiente al del día viernes.

• El TRM se tomará del día del envió. """

        observaciones_paragraph = Paragraph(observaciones_text, styles['Normal'])
        elements.append(observaciones_paragraph)

        doc.build(elements)
        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"Ocurrió un error al generar el PDF: {e}")
        return None

# Formulario
with st.form("datos_formulario"):
    fecha = st.date_input("Fecha")
    material = st.selectbox("Material", df['Material'].unique())
    encargado = st.selectbox("Encargado Pesaje FM", ['Camilo Gonzalez', 'Melissa'])
    submit_button = st.form_submit_button("Generar PDF")

# Filtrar el DataFrame según el material seleccionado
df_filtrado = df[df['Material'] == material]

# Mostrar el DataFrame filtrado antes de generar el PDF
st.write("DataFrame Filtrado:")
st.dataframe(df_filtrado)

if submit_button:
    # Mostrar valores del formulario para depuración
    st.write(f"Fecha seleccionada: {fecha}")
    st.write(f"Material seleccionado: {material}")
    st.write(f"Encargado seleccionado: {encargado}")

    pdf_buffer = generar_pdf(df_filtrado, fecha, material, encargado)
    if pdf_buffer:
        st.download_button(label='Descargar PDF', data=pdf_buffer, file_name='datos_filtrados.pdf', mime='application/pdf')
    else:
        st.write("Ocurrió un error al generar el PDF.")