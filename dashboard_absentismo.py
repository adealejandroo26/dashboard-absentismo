
import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard Absentismo", layout="wide")
st.title("✅ Dashboard de Absentismo Laboral")

# Subida de archivo
uploaded_file = st.file_uploader("Sube el archivo Excel con los datos de ausencias", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Inicio'] = pd.to_datetime(df['Inicio'])
    df['Fin'] = pd.to_datetime(df['Fin'])
    df['Horas de ausencia'] = (df['Fin'] - df['Inicio']).dt.total_seconds() / 3600
    df['Año'] = df['Inicio'].dt.year
    df['Mes'] = df['Inicio'].dt.month
    df['Mes_nombre'] = df['Inicio'].dt.strftime('%B')

    geografias = sorted(df['Geografía'].dropna().unique())
    geografias_seleccionadas = st.multiselect("Selecciona geografía(s):", geografias, default=geografias)

    anios = sorted(df['Año'].dropna().unique())
    anio_seleccionado = st.selectbox("Selecciona año:", anios)

    codigos_disponibles = sorted(df['Codigo'].dropna().unique())
    codigos_seleccionados = st.multiselect("Selecciona códigos de ausencia a incluir:", codigos_disponibles, default=codigos_disponibles)

    configuracion = {}

    st.sidebar.header("⚙️ Configuración por geografía")
    for geo in geografias_seleccionadas:
        st.sidebar.subheader(f"🌍 {geo}")
        jornada_fija = st.sidebar.number_input(f"Jornada mensual para {geo} (h)", min_value=0, value=140, step=1, key=f"jornada_{geo}")
        empleados_por_mes = {}
        for mes in range(1, 13):
            nombre_mes = datetime(2023, mes, 1).strftime('%B')
            empleados = st.sidebar.number_input(f"{geo} - {nombre_mes} - Empleados", min_value=0, value=100, step=1, key=f"{geo}_{mes}")
            empleados_por_mes[mes] = empleados
        configuracion[geo] = {
            "jornada_mensual": jornada_fija,
            "empleados_mes": empleados_por_mes
        }

    df_filtrado = df[
        (df['Geografía'].isin(geografias_seleccionadas)) &
        (df['Año'] == anio_seleccionado) &
        (df['Codigo'].isin(codigos_seleccionados))
    ]

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

        resumen_total = pd.DataFrame()

        for geo in geografias_seleccionadas:
            df_geo = df_periodo[df_periodo['Geografía'] == geo]
            resumen = df_geo.groupby('Mes')['Horas de ausencia'].sum().reset_index()
            resumen['Geografía'] = geo
            resumen['Horas teóricas'] = resumen['Mes'].apply(
                lambda m: configuracion[geo]["empleados_mes"].get(m, 0) * configuracion[geo]["jornada_mensual"]
            )
            resumen_total = pd.concat([resumen_total, resumen], ignore_index=True)

        resumen_total['Mes_nombre'] = resumen_total['Mes'].apply(lambda m: datetime(2023, m, 1).strftime('%B'))
        resumen_total['Absentismo (%)'] = (resumen_total['Horas de ausencia'] / resumen_total['Horas teóricas']) * 100

        st.subheader("📊 Gráfico de Absentismo por Mes y Geografía")
        fig = px.bar(
            resumen_total,
            x='Mes_nombre',
            y='Absentismo (%)',
            color='Geografía',
            barmode='group',
            text='Absentismo (%)',
            labels={'Mes_nombre': 'Mes'},
            title='Absentismo mensual (%) por geografía'
        )
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Detalle de cálculos")
        st.dataframe(resumen_total[['Geografía', 'Mes_nombre', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']])

        total_horas_ausencia = resumen_total['Horas de ausencia'].sum()
        total_horas_teoricas = resumen_total['Horas teóricas'].sum()
        absentismo_total = (total_horas_ausencia / total_horas_teoricas) * 100 if total_horas_teoricas > 0 else 0

        st.metric("📈 Absentismo total en el periodo seleccionado", f"{absentismo_total:.2f}%")

        if st.button("📥 Exportar datos a Excel"):
            export_df = resumen_total[['Geografía', 'Mes_nombre', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']]
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Resumen')
            st.download_button(
                label="📂 Descargar archivo Excel",
                data=buffer.getvalue(),
                file_name="resumen_absentismo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

