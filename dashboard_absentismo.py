import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard Absentismo Multianual", layout="wide")
st.title("📊 Dashboard de Absentismo (Multianual y Dinámico)")

uploaded_file = st.file_uploader("📁 Sube el archivo Excel con las ausencias", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Inicio'] = pd.to_datetime(df['Inicio'])
    df['Año'] = df['Inicio'].dt.year
    df['Mes'] = df['Inicio'].dt.month
    df['Mes_nombre'] = df['Inicio'].dt.strftime('%b')

    años_disponibles = sorted(df['Año'].unique())
    geografias = sorted(df['Geografía'].dropna().unique())
    funciones = sorted(df['Función'].dropna().unique())
    codigos = sorted(df['Codigo'].dropna().unique())

    geografias_seleccionadas = st.multiselect("🌍 Selecciona geografía(s):", geografias, default=geografias)
    funciones_seleccionadas = st.multiselect("👥 Selecciona función(es):", funciones, default=funciones)
    codigos_seleccionados = st.multiselect("📌 Selecciona códigos de ausencia:", codigos, default=codigos)

    st.sidebar.header("⚙️ Configuración por geografía + año")
    config = {}
    for geo in geografias_seleccionadas:
        for año in años_disponibles:
            st.sidebar.subheader(f"{geo} - {año}")
            jornada = st.sidebar.number_input(f"Jornada mensual ({geo}, {año})", value=140, step=1, key=f"jornada_{geo}_{año}")
            empleados = {}
            for mes in range(1, 13):
                mes_nombre = datetime(2023, mes, 1).strftime('%B')
                empleados[mes] = st.sidebar.number_input(
                    f"{geo} - {año} - {mes_nombre} - Empleados", value=100, step=1, key=f"{geo}_{año}_{mes}"
                )
            config[(geo, año)] = {"jornada": jornada, "empleados": empleados}

    st.subheader("📅 Define los rangos a comparar")
    n_rangos = st.number_input("¿Cuántos rangos?", min_value=1, max_value=10, value=1)
    rangos = []
    for i in range(n_rangos):
        col1, col2 = st.columns([1, 2])
        with col1:
            nombre = st.text_input(f"Nombre rango #{i+1}", f"Rango {i+1}", key=f"nombre_rango_{i}")
        with col2:
            fechas = st.date_input(
                f"Fechas para {nombre}",
                value=[datetime(2023, 1, 1), datetime(2023, 3, 31)],
                key=f"fecha_rango_{i}"
            )
        if isinstance(fechas, (list, tuple)) and len(fechas) == 2:
            rangos.append((nombre, pd.to_datetime(fechas[0]), pd.to_datetime(fechas[1])))

    umbral = st.number_input("📏 Índice de absentismo objetivo (%)", min_value=0.0, max_value=100.0, value=4.0, step=0.1)

    if rangos:
        df = df[df['Geografía'].isin(geografias_seleccionadas)]
        df = df[df['Función'].isin(funciones_seleccionadas)]
        df = df[df['Codigo'].isin(codigos_seleccionados)]

        resultados = []

        for nombre_rango, inicio, fin in rangos:
            df_rango = df[(df['Inicio'] >= inicio) & (df['Inicio'] <= fin)].copy()
            for geo in geografias_seleccionadas:
                df_geo = df_rango[df_rango['Geografía'] == geo]
                resumen = df_geo.groupby(['Año', 'Mes']).size().reset_index(name='Bajas')
                resumen['Mes_nombre'] = resumen['Mes'].apply(lambda m: datetime(2023, m, 1).strftime('%b'))
                resumen['Geografía'] = geo
                resumen['Rango'] = nombre_rango

                resumen['Horas por baja'] = resumen.apply(lambda row: config[(geo, row['Año'])]['jornada'] / 28, axis=1)
                resumen['Horas de ausencia'] = resumen['Bajas'] * resumen['Horas por baja']
                resumen['Horas teóricas'] = resumen.apply(
                    lambda row: config[(geo, row['Año'])]['empleados'][row['Mes']] * config[(geo, row['Año'])]['jornada'], axis=1
                )
                resumen['Absentismo (%)'] = (resumen['Horas de ausencia'] / resumen['Horas teóricas']) * 100
                resumen['Absentismo (%)'] = resumen['Absentismo (%)'].round(2)

                resultados.append(resumen)

        final = pd.concat(resultados, ignore_index=True)

        for nombre_rango in final['Rango'].unique():
            subset = final[final['Rango'] == nombre_rango]
            st.subheader(f"📊 {nombre_rango} - Barras")
            fig_bar = px.bar(
                subset,
                x='Mes_nombre',
                y='Absentismo (%)',
                color='Geografía',
                barmode='group',
                text='Absentismo (%)',
                facet_col='Año'
            )
            fig_bar.update_traces(textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True, key=f"bar_{nombre_rango}")

            st.subheader(f"📈 {nombre_rango} - Líneas")
            fig_line = px.line(
                subset,
                x='Mes_nombre',
                y='Absentismo (%)',
                color='Geografía',
                facet_col='Año',
                title=f"{nombre_rango} - vs Índice objetivo"
            )
            for geo in subset['Geografía'].unique():
                for año in subset['Año'].unique():
                    fig_line.add_scatter(
                        x=subset[(subset['Geografía'] == geo) & (subset['Año'] == año)]['Mes_nombre'],
                        y=[umbral] * len(subset[(subset['Geografía'] == geo) & (subset['Año'] == año)]),
                        mode='lines',
                        name=f'Objetivo {geo} - {año}',
                        line=dict(dash='dash', color='gray')
                    )
            st.plotly_chart(fig_line, use_container_width=True, key=f"line_{nombre_rango}")

        st.subheader("📋 Datos consolidados")
        st.dataframe(final)

        if st.button("📥 Exportar a Excel"):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                final.to_excel(writer, index=False, sheet_name='Absentismo')
            st.download_button(
                "📂 Descargar Excel",
                data=buffer.getvalue(),
                file_name="absentismo_multianual.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
