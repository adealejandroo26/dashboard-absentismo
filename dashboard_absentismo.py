
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px

st.set_page_config(page_title="Dashboard Absentismo", layout="wide")
st.title("✅ Dashboard de Absentismo Laboral")

# Cargar archivo Excel
uploaded_file = st.file_uploader("Sube el archivo Excel con los datos de ausencias", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Inicio'] = pd.to_datetime(df['Inicio'])
    df['Fin'] = pd.to_datetime(df['Fin'])
    df['Horas de ausencia'] = (df['Fin'] - df['Inicio']).dt.total_seconds() / 3600
    df['Año'] = df['Inicio'].dt.year
    df['Mes'] = df['Inicio'].dt.month
    df['Mes_nombre'] = df['Inicio'].dt.strftime('%B')

    geografias = df['Geografía'].dropna().unique()
    geografias_seleccionadas = st.multiselect("Selecciona geografía(s):", sorted(geografias), default=geografias.tolist())

    anios = df['Año'].dropna().unique()
    anio_seleccionado = st.selectbox("Selecciona año:", sorted(anios))

    codigos_disponibles = df['Codigo'].dropna().unique()
    codigos_seleccionados = st.multiselect("Selecciona códigos de ausencia a incluir:", sorted(codigos_disponibles), default=sorted(codigos_disponibles))

    df_filtrado = df[
        (df['Geografía'].isin(geografias_seleccionadas)) &
        (df['Año'] == anio_seleccionado) &
        (df['Codigo'].isin(codigos_seleccionados))
    ]

    st.sidebar.header("📅 Configuración por mes")
    empleados_mes = {}
    jornada_mes = {}
    for mes in range(1, 13):
        nombre_mes = datetime(2023, mes, 1).strftime('%B')
        empleados_mes[mes] = st.sidebar.number_input(f"{nombre_mes} - Empleados", min_value=0, value=100, step=1)
        jornada_mes[mes] = st.sidebar.number_input(f"{nombre_mes} - Jornada mensual (h)", min_value=0, value=140, step=1)

    # Selector de rango de fechas personalizado
    st.subheader("📆 Selecciona el rango de fechas para el análisis")
    fecha_min = df_filtrado['Inicio'].min()
    fecha_max = df_filtrado['Fin'].max()
    rango = st.date_input("Rango de fechas", [fecha_min, fecha_max], min_value=fecha_min, max_value=fecha_max)

    if len(rango) == 2:
        inicio_rango, fin_rango = rango
        df_periodo = df_filtrado[
            (df_filtrado['Inicio'] >= pd.to_datetime(inicio_rango)) &
            (df_filtrado['Inicio'] <= pd.to_datetime(fin_rango))
        ]

        resumen = df_periodo.groupby('Mes')['Horas de ausencia'].sum().reset_index()
        resumen['Horas teóricas'] = resumen['Mes'].apply(lambda m: empleados_mes.get(m, 0) * jornada_mes.get(m, 0))
        resumen['Absentismo (%)'] = (resumen['Horas de ausencia'] / resumen['Horas teóricas']) * 100
        resumen['Mes_nombre'] = resumen['Mes'].apply(lambda m: datetime(2023, m, 1).strftime('%B'))

        st.subheader("📊 Gráfico de Absentismo por Mes")
        fig = px.bar(
            resumen,
            x='Mes_nombre',
            y='Absentismo (%)',
            text='Absentismo (%)',
            labels={'Mes_nombre': 'Mes'},
            title='Absentismo mensual (%)'
        )
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Detalle de cálculos")
        st.dataframe(resumen[['Mes_nombre', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']])

        total_horas_ausencia = resumen['Horas de ausencia'].sum()
        total_horas_teoricas = resumen['Horas teóricas'].sum()
        absentismo_total = (total_horas_ausencia / total_horas_teoricas) * 100 if total_horas_teoricas > 0 else 0

        st.metric("📈 Absentismo total en el periodo seleccionado", f"{absentismo_total:.2f}%")

        # Exportar a Excel y permitir descarga
        from io import BytesIO

        if st.button("📥 Exportar datos a Excel"):
            export_df = resumen[['Mes_nombre', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']]
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Resumen')
            st.download_button(
                label="📂 Descargar archivo Excel",
                data=buffer.getvalue(),
                file_name="resumen_absentismo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
