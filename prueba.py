import pandas as pd
import openpyxl
import csv
import streamlit as st
import os
import plotly.graph_objects as go
from io import BytesIO

# Paso 1: Importar las librerías necesarias
st.title("Carga y proceso de centralizado BAT")

# Paso 2: Subir el archivo semanal "centralizado BAT" desde la interfaz de Streamlit
archivo_subido = st.file_uploader("Sube el archivo", type=["xlsx"])

if archivo_subido is None:
    st.info("Sube el archivo de centralizado")
    st.stop()
