import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard Comparativo", layout="wide")
st.title("📊 Comparativo de Absentismo por Rango")

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

    st.subheader("📆 Define los rangos a comparar")
    num_rangos = st.number_input("¿Cuántos rangos quieres comparar?", min_value=1, max_value=10, value=2, step=1)
    rangos = []

    for i in range(num_rangos):
        col1, col2 = st.columns([1, 2])
        with col1:
            nombre = st.text_input(f"Nombre para el rango #{i+1}", f"Rango {i+1}", key=f"nombre_rango_{i}")
        with col2:
            fechas = st.date_input(f"Fechas para {nombre}", value=[datetime(2023, 1, 1), datetime(2023, 3, 31)], key=f"fecha_rango_{i}")
        if isinstance(fechas, (list, tuple)) and len(fechas) == 2:
            rangos.append((nombre, pd.to_datetime(fechas[0]), pd.to_datetime(fechas[1])))

    umbral = st.number_input("Índice de absentismo objetivo (%)", min_value=0.0, max_value=100.0, value=4.0, step=0.1)

    if rangos:
        df_filtrado = df[
            (df['Geografía'].isin(geografias_seleccionadas)) &
            (df['Codigo'].isin(codigos_seleccionados)) &
            (df['Función'].isin(funciones_seleccionadas))
        ]

        resumen_completo = []

        for nombre_rango, inicio, fin in rangos:
            df_rango = df_filtrado[(df_filtrado['Inicio'] >= inicio) & (df_filtrado['Inicio'] <= fin)].copy()
            resumen_total = pd.DataFrame()

            for geo in geografias_seleccionadas:
                df_geo = df_rango[df_rango['Geografía'] == geo]
                resumen = df_geo.groupby('Mes')['Horas de ausencia'].sum().reset_index()
                resumen['Geografía'] = geo
                resumen['Mes_nombre'] = resumen['Mes'].apply(lambda m: datetime(2023, m, 1).strftime('%b'))
                resumen['Horas teóricas'] = resumen['Mes'].apply(
                    lambda m: configuracion[geo]["empleados_mes"].get(m, 0) * configuracion[geo]["jornada_mensual"]
                )
                resumen_total = pd.concat([resumen_total, resumen], ignore_index=True)

            resumen_total['Absentismo (%)'] = (resumen_total['Horas de ausencia'] / resumen_total['Horas teóricas']) * 100
            resumen_total['Absentismo (%)'] = resumen_total['Absentismo (%)'].round(2)
            resumen_total['Rango'] = nombre_rango

            resumen_completo.append(resumen_total)

            st.subheader(f"📊 {nombre_rango}: Gráfico de Barras")
            fig_bar = px.bar(
                resumen_total,
                x='Mes_nombre',
                y='Absentismo (%)',
                color='Geografía',
                barmode='group',
                text=resumen_total['Absentismo (%)'].astype(str) + '%',
                title=f"Absentismo por Mes y Geografía - {nombre_rango}"
            )
            fig_bar.update_traces(textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True)

            st.subheader(f"📈 {nombre_rango}: Gráfico de Líneas con Objetivo")
            fig_line = px.line(
                resumen_total,
                x='Mes_nombre',
                y='Absentismo (%)',
                color='Geografía',
                title=f"Absentismo vs. Objetivo - {nombre_rango}"
            )
            for geo in resumen_total['Geografía'].unique():
                fig_line.add_scatter(
                    x=resumen_total[resumen_total['Geografía'] == geo]['Mes_nombre'],
                    y=[umbral] * len(resumen_total[resumen_total['Geografía'] == geo]),
                    mode='lines',
                    name=f'Objetivo {geo}',
                    line=dict(dash='dash', color='gray')
                )
            st.plotly_chart(fig_line, use_container_width=True)

        st.subheader("📋 Datos consolidados de todos los rangos")
        df_final = pd.concat(resumen_completo, ignore_index=True)
        st.dataframe(df_final[['Rango', 'Geografía', 'Mes_nombre', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']])

        if st.button("📥 Exportar a Excel"):
            export_df = df_final[['Rango', 'Geografía', 'Mes_nombre', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']]
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Comparativo')
            st.download_button(
                label="📂 Descargar archivo Excel",
                data=buffer.getvalue(),
                file_name="comparativo_absentismo_rangos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )




