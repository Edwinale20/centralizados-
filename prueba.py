import pandas as pd
import openpyxl
import streamlit as st
from io import BytesIO
import plotly.figure_factory as ff
import plotly.graph_objects as go

def cargar_procesar_archivo():
    # Paso 1: Título de la app
    st.title("Carga y proceso de centralizado BAT")

    # Paso 2: Subir el archivo semanal "centralizado BAT"
    archivo_subido = st.file_uploader("Sube el archivo", type=["xlsx"])

    if archivo_subido is None:
        st.info("Sube el archivo de centralizado")
        st.stop()

    # Opción para elegir el tipo de pedido
    tipo_pedido = st.selectbox("Selecciona el tipo de pedido:", ["stock"])

    if archivo_subido:
        try:
            # Verificar que el archivo contenga la hoja 'DETALLE PEDIDO'
            with pd.ExcelFile(archivo_subido) as xls:
                if 'DETALLE PEDIDO' in xls.sheet_names:
                    dataframe_bat = pd.read_excel(xls, sheet_name='DETALLE PEDIDO')
                    st.write("Archivo leído correctamente.")
                else:
                    st.error("La hoja 'DETALLE PEDIDO' no existe en el archivo subido.")
                    st.stop()

            # Mantener solo las columnas de interés
            columnas_interes = ['PLAZA BAT', 'N TIENDA', 'UPC', 'SKU 7 ELEVEN', 'ARTICULO 7 ELEVEN', 
                                'CAJETILLAS X PQT', 'CAJETILLAS', 'PAQUETES', 'FECHA DE PEDIDO']
            dataframe_bat = dataframe_bat[[col for col in columnas_interes if col in dataframe_bat.columns]]

            # Asegurarse de que la columna PAQUETES sea numérica
            dataframe_bat['PAQUETES'] = pd.to_numeric(dataframe_bat['PAQUETES'], errors='coerce').fillna(0)

            # Verificar que las columnas necesarias existan en el DataFrame
            if 'PLAZA BAT' in dataframe_bat.columns and 'FECHA DE PEDIDO' in dataframe_bat.columns and 'PAQUETES' in dataframe_bat.columns:
                # Eliminar filas con valores nulos
                dataframe_bat = dataframe_bat.dropna(subset=['PLAZA BAT', 'FECHA DE PEDIDO', 'PAQUETES'])

                # Asegurarse de que las columnas tienen el tipo de dato correcto
                dataframe_bat['PLAZA BAT'] = dataframe_bat['PLAZA BAT'].astype(str)
                dataframe_bat['FECHA DE PEDIDO'] = pd.to_datetime(dataframe_bat['FECHA DE PEDIDO'], errors='coerce')
                dataframe_bat['PAQUETES'] = pd.to_numeric(dataframe_bat['PAQUETES'], errors='coerce')
                
                # Verificar que no haya valores nulos después de la conversión
                dataframe_bat = dataframe_bat.dropna(subset=['PLAZA BAT', 'FECHA DE PEDIDO', 'PAQUETES'])

                # Calcular la suma de paquetes para cada PLAZA BAT
                suma_paquetes = dataframe_bat.groupby(['PLAZA BAT', 'FECHA DE PEDIDO'])['PAQUETES'].sum().reset_index()
                suma_paquetes.columns = ['PLAZA', 'FECHA DE PEDIDO', 'PAQUETES']

                # Formatear las fechas
                suma_paquetes['FECHA DE PEDIDO'] = suma_paquetes['FECHA DE PEDIDO'].dt.strftime('%Y-%m-%d')
                suma_paquetes['FECHA DE ENTREGA'] = (pd.to_datetime(suma_paquetes['FECHA DE PEDIDO']) + pd.to_timedelta(1, unit='d')).dt.strftime('%Y-%m-%d')

                # Crear tabla con columnas adicionales vacías
                plazas = {
                    'REYNOSA': '100 110', 'MÉXICO': '200', 'JALISCO': '300', 'SALTILLO': '400 410',
                    'MONTERREY': '500', 'BAJA CALIFORNIA': '600 610 620', 'HERMOSILLO': '650',
                    'PUEBLA': '700', 'CUERNAVACA': '720', 'YUCATAN': '800', 'QUINTANA ROO': '890'
                }
                suma_paquetes['ID PLAZA'] = suma_paquetes['PLAZA'].map(lambda x: plazas.get(x, ''))
                suma_paquetes['FOLIOS'] = ''
                suma_paquetes['TIPO DE PEDIDO'] = tipo_pedido.capitalize()

                # Ordenar plazas
                orden_plazas = ['REYNOSA', 'MÉXICO', 'JALISCO', 'SALTILLO', 'MONTERREY', 'BAJA CALIFORNIA', 
                                'HERMOSILLO', 'PUEBLA', 'CUERNAVACA', 'YUCATAN', 'QUINTANA ROO']
                suma_paquetes['PLAZA'] = pd.Categorical(suma_paquetes['PLAZA'], categories=orden_plazas, ordered=True)
                suma_paquetes = suma_paquetes.sort_values(['PLAZA', 'FECHA DE PEDIDO'])

                # Mostrar tabla
                st.write(suma_paquetes)

                # Opción para descargar la tabla
                csv = suma_paquetes.to_csv(index=False)
                st.download_button(
                    label="Descargar tabla",
                    data=csv,
                    file_name='Tabla para correo.csv',
                    mime='text/csv',
                )

                return suma_paquetes  # Devolvemos suma_paquetes para los siguientes cálculos

            else:
                st.error("Las columnas necesarias ('PLAZA BAT', 'FECHA DE PEDIDO', 'PAQUETES') no están presentes en el archivo subido.")
                st.stop()

        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")
            st.stop()

def crear_grafico(suma_paquetes):
    # Crear gráfico comparativo
    st.title("Gráfica Comparativa de Paquetes por Plaza BAT")

    # Definir límites de paquetes por plaza
    limites_paquetes = {
        'Noreste': 22000, 'MÉXICO': 8000, 'PENÍNSULA': 2000, 'HERMOSILLO': 2000,
        'JALISCO': 4000, 'BAJA CALIFORNIA': 4000
    }

    # Calcular la suma de paquetes por agrupaciones
    paquetes_noreste = suma_paquetes[suma_paquetes['PLAZA'].isin(['REYNOSA', 'MONTERREY', 'SALTILLO'])]['PAQUETES'].sum()
    paquetes_peninsula = suma_paquetes[suma_paquetes['PLAZA'].isin(['YUCATAN', 'QUINTANA ROO'])]['PAQUETES'].sum()
    paquetes_mexico = suma_paquetes[suma_paquetes['PLAZA'].isin(['MÉXICO', 'PUEBLA', 'CUERNAVACA'])]['PAQUETES'].sum()

    # Crear un nuevo DataFrame con las agrupaciones
    data = {
        'Plaza': ['Noreste', 'MÉXICO', 'PENÍNSULA', 'HERMOSILLO', 'JALISCO', 'BAJA CALIFORNIA'],
        'Paquetes': [paquetes_noreste, paquetes_mexico, paquetes_peninsula,
                     suma_paquetes[suma_paquetes['PLAZA'] == 'HERMOSILLO']['PAQUETES'].sum(),
                     suma_paquetes[suma_paquetes['PLAZA'] == 'JALISCO']['PAQUETES'].sum(),
                     suma_paquetes[suma_paquetes['PLAZA'] == 'BAJA CALIFORNIA']['PAQUETES'].sum()],
        'Límite': [22000, 8000, 2000, 2000, 4000, 4000]
    }

    df_comparativa = pd.DataFrame(data)

    # Crear la tabla para la comparación
    table_data = [['Plaza', 'Paquetes', 'Límite']] + df_comparativa.values.tolist()

    # Inicializar la figura con la tabla
    fig = ff.create_table(table_data, height_constant=60)

    # Crear trazos para la gráfica de barras
    trace1 = go.Bar(x=df_comparativa['Plaza'], y=df_comparativa['Paquetes'], name='Paquetes')
    trace2 = go.Bar(x=df_comparativa['Plaza'], y=df_comparativa['Límite'], name='Límite')

    # Añadir trazos a la figura
    fig.add_traces([trace1, trace2])

    # Mostrar la gráfica
    st.plotly_chart(fig)

# Llamar a las funciones
suma_paquetes = cargar_procesar_archivo()
if suma_paquetes is not None:
    crear_grafico(suma_paquetes)
