# Paso 1: Importar las librerías necesarias
import pandas as pd
import openpyxl
import csv
import streamlit as st
import os
import plotly.graph_objects as go

# Ruta de la carpeta de OneDrive del usuario
ruta_onedrive = r"C:\Users\omen0\OneDrive\Documentos"

# Paso 2: Subir el archivo semanal "centralizado BAT" desde la interfaz de Streamlit
st.title("Carga y proceso de 'centralizado BAT'")
  
# Opción para cargar el archivo
archivo_subido = st.file_uploader("Sube el archivo", type=["xlsx"])

# Opción para elegir el tipo de pedido
tipo_pedido = st.selectbox("Selecciona el tipo de pedido:", ["stock", "complementario"])

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
        columnas_interes = ['PLAZA BAT', 'N TIENDA', 'UPC', 'SKU 7 ELEVEN', 'ARTICULO 7 ELEVEN', 'CAJETILLAS X PQT', 'CAJETILLAS', 'PAQUETES', 'FECHA DE PEDIDO']
        dataframe_bat = dataframe_bat[[col for col in columnas_interes if col in dataframe_bat.columns]]

        # Asegurarse de que la columna PAQUETES sea numérica
        dataframe_bat['PAQUETES'] = pd.to_numeric(dataframe_bat['PAQUETES'], errors='coerce').fillna(0)

    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")

    # Definir columnas sin PAQUETES
    columnas_sin_paquetes = ['UPC', 'SKU 7 ELEVEN', 'ARTICULO 7 ELEVEN', 'CAJETILLAS X PQT', 'CAJETILLAS']

    # Paso 3: Filtrar por PLAZA BAT o N TIENDA usando un botón y mostrar todas las columnas (sin la columna PAQUETES)
    st.title("Filtrar por PLAZA BAT o N TIENDA")

    # Determinar la columna para filtrar
    columna_filtrar = 'N TIENDA' if 'N TIENDA' in dataframe_bat.columns else 'PLAZA BAT'

    # Input de usuario para seleccionar la PLAZA BAT o N TIENDA
    if columna_filtrar in dataframe_bat.columns:
        seleccion_filtrar = st.selectbox(f'Selecciona la {columna_filtrar}:', dataframe_bat[columna_filtrar].unique())
    else:
        st.error(f"La columna '{columna_filtrar}' no existe en el archivo.")

    # Botón para aplicar el filtro
    if st.button('Filtrar'):
        dataframe_filtrado = dataframe_bat[dataframe_bat[columna_filtrar] == seleccion_filtrar][columnas_sin_paquetes]
        st.write(f"Filtrado por {columna_filtrar}: {seleccion_filtrar}")
        st.write(dataframe_filtrado)

# Paso 4: Filtrar por cada plaza y guardar en archivos separados por fecha de pedido (sin la columna PAQUETES)
plazas = {
    'REYNOSA': ('100 110', '9271'),
    'MÉXICO': ('200', '9211'),
    'JALISCO': ('300', '9221'),
    'SALTILLO': ('400 410', '9261'),
    'MONTERREY': ('500', '9201'),
    'BAJA CALIFORNIA': ('600 610 620', '9231'),
    'HERMOSILLO': ('650', '9251'),
    'PUEBLA': ('700', '9291'),
    'CUERNAVACA': ('720', '9281'),
    'YUCATAN': ('800', '9241'),
    'QUINTANA ROO': ('890', '9289')
}

# Crear carpeta si no existe
carpeta_destino = os.path.join(ruta_onedrive, "centralizados semanal BAT")
if not os.path.exists(carpeta_destino):
    os.makedirs(carpeta_destino)
    st.write(f"Carpeta '{carpeta_destino}' creada en OneDrive.")

# Botón para guardar archivos
if st.button('Guardar Archivos'):
    fechas_pedido = dataframe_bat['FECHA DE PEDIDO'].unique()
    for fecha in fechas_pedido:
        fecha_str = pd.to_datetime(fecha).strftime("%d%m%Y")
        for plaza, (codigo, id_tienda) in plazas.items():
            # Mostrar un mensaje conciso indicando la plaza y la fecha
            st.write(f"Procesando plaza: {plaza}, fecha: {fecha_str}")
            
            # Verificar la lógica de filtrado y mostrar mensajes de depuración concisos
            if tipo_pedido == "complementario" and 'N TIENDA' in dataframe_bat.columns:
                df_plaza = dataframe_bat[(dataframe_bat['N TIENDA'] == plaza) & (dataframe_bat['FECHA DE PEDIDO'] == fecha)][columnas_sin_paquetes]
                st.write(f"Filtrando por N TIENDA: {plaza}")
            else:
                df_plaza = dataframe_bat[(dataframe_bat['PLAZA BAT'] == plaza) & (dataframe_bat['FECHA DE PEDIDO'] == fecha)][columnas_sin_paquetes]
                st.write(f"Filtrando por PLAZA BAT: {plaza}")

            # Mostrar solo la cantidad de filas filtradas
            st.write(f"Datos filtrados: {df_plaza.shape[0]} filas")

            if not df_plaza.empty:
                df_plaza.insert(0, 'ID Tienda', id_tienda)  # Insertar la columna ID Tienda como la primera columna
                # Cambiar nombres de columnas
                df_plaza.columns = ['id Tienda', 'Codigo de Barras', 'Id Articulo', 'Descripcion', 'Unidad Empaque', 'Cantidad (Pza)']
                nombre_archivo = f"{codigo} {fecha_str}.csv"
                ruta_archivo = os.path.join(carpeta_destino, nombre_archivo)
                df_plaza.to_csv(ruta_archivo, index=False)
                st.write(f"Archivo guardado: {ruta_archivo}")
            else:
                st.warning(f"No se encontraron datos para la {columna_filtrar} {plaza} en la fecha {fecha_str}")
    st.write("Proceso completado.")
  
    # Paso 6: Crear tabla con la suma de paquetes para cada PLAZA BAT
    st.title("Tabla de Suma de Paquetes por PLAZA BAT")

    # Calcular la suma de paquetes para cada PLAZA BAT
    suma_paquetes = dataframe_bat.groupby(['PLAZA BAT', 'FECHA DE PEDIDO'])['PAQUETES'].sum().reset_index()
    suma_paquetes.columns = ['PLAZA', 'FECHA DE PEDIDO', 'PAQUETES']

    # Formatear las fechas para que no incluyan la hora
    suma_paquetes['FECHA DE PEDIDO'] = pd.to_datetime(suma_paquetes['FECHA DE PEDIDO']).dt.strftime('%Y-%m-%d')
    suma_paquetes['FECHA DE ENTREGA'] = (pd.to_datetime(suma_paquetes['FECHA DE PEDIDO']) + pd.to_timedelta(1, unit='d')).dt.strftime('%Y-%m-%d')

    # Crear tabla con columnas adicionales vacías
    suma_paquetes['ID PLAZA'] = suma_paquetes['PLAZA'].map(lambda x: plazas[x][0])
    suma_paquetes['FOLIOS'] = ''
    suma_paquetes['TIPO DE PEDIDO'] = tipo_pedido.capitalize()

    # Ordenar las plazas de menor a mayor
    orden_plazas = ['REYNOSA', 'MÉXICO', 'JALISCO', 'SALTILLO', 'MONTERREY', 'BAJA CALIFORNIA', 'HERMOSILLO', 'PUEBLA', 'CUERNAVACA', 'YUCATAN', 'QUINTANA ROO']
    suma_paquetes['PLAZA'] = pd.Categorical(suma_paquetes['PLAZA'], categories=orden_plazas, ordered=True)
    suma_paquetes = suma_paquetes.sort_values('PLAZA')

    # Reorganizar las columnas
    suma_paquetes = suma_paquetes[['PLAZA', 'ID PLAZA', 'PAQUETES', 'FOLIOS', 'FECHA DE PEDIDO', 'FECHA DE ENTREGA', 'TIPO DE PEDIDO']]
    st.write(suma_paquetes)

    # Opción para copiar el DataFrame
    st.title("Copiar DataFrame")
    csv = suma_paquetes.to_csv(index=False)
    st.download_button(
        label="Copiar Tabla",
        data=csv,
        file_name='suma_paquetes.csv',
        mime='text/csv',
    )

    # Paso 6: Crear tabla con la suma de paquetes para cada PLAZA BAT
    st.title("Tabla de Suma de Paquetes por PLAZA BAT")

    # Calcular la suma de paquetes para cada PLAZA BAT
    suma_paquetes = dataframe_bat.groupby(['PLAZA BAT', 'FECHA DE PEDIDO'])['PAQUETES'].sum().reset_index()
    suma_paquetes.columns = ['PLAZA', 'FECHA DE PEDIDO', 'PAQUETES']

    # Formatear las fechas para que no incluyan la hora
    suma_paquetes['FECHA DE PEDIDO'] = pd.to_datetime(suma_paquetes['FECHA DE PEDIDO']).dt.strftime('%Y-%m-%d')
    suma_paquetes['FECHA DE ENTREGA'] = (pd.to_datetime(suma_paquetes['FECHA DE PEDIDO']) + pd.to_timedelta(1, unit='d')).dt.strftime('%Y-%m-%d')

    # Crear tabla con columnas adicionales vacías
    suma_paquetes['ID PLAZA'] = suma_paquetes['PLAZA'].map(lambda x: plazas[x][0])
    suma_paquetes['FOLIOS'] = ''
    suma_paquetes['TIPO DE PEDIDO'] = tipo_pedido.capitalize()

    # Ordenar las plazas de menor a mayor
    orden_plazas = ['REYNOSA', 'MÉXICO', 'JALISCO', 'SALTILLO', 'MONTERREY', 'BAJA CALIFORNIA', 'HERMOSILLO', 'PUEBLA', 'CUERNAVACA', 'YUCATAN', 'QUINTANA ROO']
    suma_paquetes['PLAZA'] = pd.Categorical(suma_paquetes['PLAZA'], categories=orden_plazas, ordered=True)
    suma_paquetes = suma_paquetes.sort_values('PLAZA')

    # Reorganizar las columnas
    suma_paquetes = suma_paquetes[['PLAZA', 'ID PLAZA', 'PAQUETES', 'FOLIOS', 'FECHA DE PEDIDO', 'FECHA DE ENTREGA', 'TIPO DE PEDIDO']]
    st.write(suma_paquetes)

    # Paso 7: Crear gráficos de barras comparativos de paquetes por plaza BAT y sus límites
    st.title("Gráfica Comparativa de Paquetes por Plaza BAT")

    # Definir límites de paquetes por plaza
    limites_paquetes = {
        'Noreste': 22000,
        'MÉXICO': 8000,
        'PENÍNSULA': 2000,
        'HERMOSILLO': 2000,
        'JALISCO': 4000,
        'BAJA CALIFORNIA': 4000
    }

    # Calcular la suma de paquetes por agrupaciones específicas
    paquetes_noreste = suma_paquetes[suma_paquetes['PLAZA'].isin(['REYNOSA', 'MONTERREY', 'SALTILLO'])]['PAQUETES'].sum()
    paquetes_peninsula = suma_paquetes[suma_paquetes['PLAZA'].isin(['YUCATAN', 'QUINTANA ROO'])]['PAQUETES'].sum()

    # Crear un nuevo DataFrame con las agrupaciones
    data = {
        'Plaza': ['Noreste', 'MÉXICO', 'PENÍNSULA', 'HERMOSILLO', 'JALISCO', 'BAJA CALIFORNIA'],
        'Paquetes': [
            paquetes_noreste,
            suma_paquetes[suma_paquetes['PLAZA'] == 'MÉXICO']['PAQUETES'].sum(),
            paquetes_peninsula,
            suma_paquetes[suma_paquetes['PLAZA'] == 'HERMOSILLO']['PAQUETES'].sum(),
            suma_paquetes[suma_paquetes['PLAZA'] == 'JALISCO']['PAQUETES'].sum(),
            suma_paquetes[suma_paquetes['PLAZA'] == 'BAJA CALIFORNIA']['PAQUETES'].sum()
        ],
        'Límite': [22000, 8000, 2000, 2000, 4000, 4000]
    }

    df_comparativa = pd.DataFrame(data)

    # Crear una tabla para la comparación
    table_data = [
        ['Plaza', 'Paquetes', 'Límite'],
        *df_comparativa.values.tolist()
    ]

    # Inicializar la figura con la tabla
    fig = ff.create_table(table_data, height_constant=60)

    # Crear trazos para la gráfica de barras
    trace1 = go.Bar(x=df_comparativa['Plaza'], y=df_comparativa['Paquetes'], xaxis='x2', yaxis='y2',
                    marker=dict(color='orange'),
                    name='Paquetes')
    trace2 = go.Bar(x=df_comparativa['Plaza'], y=df_comparativa['Límite'], xaxis='x2', yaxis='y2',
                    marker=dict(color='green'),
                    name='Límite')

    # Añadir trazos a la figura
    fig.add_traces([trace1, trace2])

    # Inicializar ejes x2 y y2
    fig['layout']['xaxis2'] = {}
    fig['layout']['yaxis2'] = {}

    # Editar el diseño para subplots
    fig.layout.yaxis.update({'domain': [0, .45]})
    fig.layout.yaxis2.update({'domain': [.6, 1]})

    # Anclar los ejes x2 y y2
    fig.layout.yaxis2.update({'anchor': 'x2'})
    fig.layout.xaxis2.update({'anchor': 'y2'})
    fig.layout.yaxis2.update({'title': 'Cantidad de Paquetes'})

    # Actualizar los márgenes para añadir título y ver las etiquetas
    fig.layout.margin.update({'t':75, 'l':50})
    fig.layout.update({'title': 'Comparativa de Paquetes por Plaza BAT'})

    # Actualizar la altura debido a la interacción con la tabla
    fig.layout.update({'height':800})

    # Mostrar la gráfica en Streamlit
    st.plotly_chart(fig)
