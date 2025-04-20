import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard Comparativo", layout="wide")
st.title("📊 Comparativo de Absentismo entre Rango de Fechas")

uploaded_file = st.file_uploader("Sube el archivo Excel con los datos de ausencias", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Inicio'] = pd.to_datetime(df['Inicio'])
    df['Fin'] = pd.to_datetime(df['Fin'])
    df['Horas de ausencia'] = (df['Fin'] - df['Inicio']).dt.total_seconds() / 3600
    df['Año'] = df['Inicio'].dt.year
    df['Mes'] = df['Inicio'].dt.month
    df['Mes_nombre'] = df['Inicio'].dt.strftime('%b')

    geografias = sorted(df['Geografía'].dropna().unique())
    geografias_seleccionadas = st.multiselect("Selecciona geografía(s):", geografias, default=geografias)

    codigos_disponibles = sorted(df['Codigo'].dropna().unique())
    codigos_seleccionados = st.multiselect("Selecciona códigos de ausencia:", codigos_disponibles, default=codigos_disponibles)

    funciones_disponibles = sorted(df['Función'].dropna().unique())
    funciones_seleccionadas = st.multiselect("Selecciona función(es):", funciones_disponibles, default=funciones_disponibles)

    st.sidebar.header("⚙️ Configuración por geografía")
    configuracion = {}
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

    st.subheader("📆 Añade y nombra los rangos comparativos")
    num_rangos = st.number_input("¿Cuántos rangos deseas comparar?", min_value=1, max_value=10, value=2, step=1)
    rangos = []

    for i in range(num_rangos):
        col1, col2 = st.columns([1, 2])
        with col1:
            nombre = st.text_input(f"Nombre para el rango #{i+1}", f"Rango {i+1}", key=f"nombre_rango_{i}")
        with col2:
            fechas = st.date_input(f"Fechas para {nombre}", key=f"fecha_rango_{i}")
        if len(fechas) == 2:
            rangos.append((nombre, pd.to_datetime(fechas[0]), pd.to_datetime(fechas[1])))

    umbral = st.number_input("Índice de absentismo objetivo (%)", min_value=0.0, max_value=100.0, value=4.0, step=0.1)

    if rangos:
        df = df[
            (df['Geografía'].isin(geografias_seleccionadas)) &
            (df['Codigo'].isin(codigos_seleccionados)) &
            (df['Función'].isin(funciones_seleccionadas))
        ]

        resumen_final = pd.DataFrame()

        for nombre_rango, inicio, fin in rangos:
            df_rango = df[(df['Inicio'] >= inicio) & (df['Inicio'] <= fin)].copy()
            df_rango['Rango'] = nombre_rango
            for geo in geografias_seleccionadas:
                df_geo = df_rango[df_rango['Geografía'] == geo]
                resumen = df_geo.groupby('Mes')['Horas de ausencia'].sum().reset_index()
                resumen['Geografía'] = geo
                resumen['Rango'] = nombre_rango
                resumen['Horas teóricas'] = resumen['Mes'].apply(
                    lambda m: configuracion[geo]['empleados_mes'][m] * configuracion[geo]['jornada_mensual']
                )
                resumen_final = pd.concat([resumen_final, resumen], ignore_index=True)

        resumen_final['Mes_nombre'] = resumen_final['Mes'].apply(lambda m: datetime(2023, m, 1).strftime('%b'))
        resumen_final['Absentismo (%)'] = (resumen_final['Horas de ausencia'] / resumen_final['Horas teóricas']) * 100
        resumen_final['Absentismo (%)'] = resumen_final['Absentismo (%)'].round(2)

        st.subheader("📊 Gráfico comparativo de barras por Rango")
        fig_bar = px.bar(
            resumen_final,
            x='Mes_nombre',
            y='Absentismo (%)',
            color='Rango',
            barmode='group',
            facet_col='Geografía',
            title='Absentismo por Mes, Geografía y Rango'
        )
        fig_bar.update_traces(texttemplate='%{y}%', textposition='outside')
        st.plotly_chart(fig_bar, use_container_width=True)

        st.subheader("📈 Comparación con Índice Objetivo")
        fig_line = px.line(
            resumen_final,
            x='Mes_nombre',
            y='Absentismo (%)',
            color='Rango',
            line_group='Rango',
            facet_col='Geografía',
            title='Absentismo vs Índice Objetivo por Rango'
        )
        for geo in resumen_final['Geografía'].unique():
            fig_line.add_scatter(
                x=resumen_final['Mes_nombre'].unique(),
                y=[umbral] * 12,
                mode='lines',
                name=f'Objetivo {geo}',
                line=dict(dash='dash', color='gray'),
                row=1,
                col=1
            )
        st.plotly_chart(fig_line, use_container_width=True)

        st.subheader("📋 Detalle de datos comparativos")
        st.dataframe(resumen_final[['Rango', 'Geografía', 'Mes_nombre', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']])

        if st.button("📥 Exportar comparativo a Excel"):
            export_df = resumen_final[['Rango', 'Geografía', 'Mes_nombre', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']]
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Comparativo')
            st.download_button(
                label="📂 Descargar comparativo Excel",
                data=buffer.getvalue(),
                file_name="comparativo_absentismo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

