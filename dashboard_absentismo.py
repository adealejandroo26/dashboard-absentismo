import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
from io import BytesIO
import json

st.set_page_config(page_title="Dashboard Absentismo Multianual", layout="wide")
st.title("\U0001F4CA Dashboard de Absentismo (Multianual y Dinámico)")

config_file = st.file_uploader("\U0001F4C2 Cargar configuración (JSON)", type=["json"])
if config_file:
    saved_config = json.load(config_file)
    raw_config = saved_config.get("config", {})
    config = {}
    for k, v in raw_config.items():
        geo, ano = k.split("__")
        config[(geo, int(ano))] = v
else:
    saved_config = {}
    config = {}

uploaded_file = st.file_uploader("\U0001F4C1 Sube el archivo Excel con las ausencias", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Inicio'] = pd.to_datetime(df['Inicio'])
    df['Año'] = df['Inicio'].dt.year
    df['Mes'] = df['Inicio'].dt.month
    df['Mes_nombre'] = df['Inicio'].dt.strftime('%b')

    anos_disponibles = sorted(df['Año'].unique())
    geografias = sorted(df['Geografía'].dropna().unique())
    funciones = sorted(df['Función'].dropna().unique())
    codigos = sorted(df['Codigo'].dropna().unique())

    geografias_seleccionadas = st.multiselect("\U0001F30D Selecciona geografía(s):", geografias, default=saved_config.get("geografias", geografias))
    funciones_seleccionadas = st.multiselect("\U0001F465 Selecciona función(es):", funciones, default=saved_config.get("funciones", funciones))
    codigos_seleccionados = st.multiselect("\U0001F4CC Selecciona códigos de ausencia:", codigos, default=saved_config.get("codigos", codigos))

    st.sidebar.header("\u2699\ufe0f Configuración por geografía + año")
    for geo in geografias_seleccionadas:
        for ano in anos_disponibles:
            st.sidebar.subheader(f"{geo} - {ano}")
            jornada = st.sidebar.number_input(f"Jornada mensual ({geo}, {ano})", value=config.get((geo, ano), {}).get("jornada", 140), step=1, key=f"jornada_{geo}_{ano}")
            empleados = {}
            for mes in range(1, 13):
                mes_nombre = datetime(2023, mes, 1).strftime('%B')
                empleados[mes] = st.sidebar.number_input(
                    f"{geo} - {ano} - {mes_nombre} - Empleados", value=config.get((geo, ano), {}).get("empleados", {}).get(mes, 100), step=1, key=f"{geo}_{ano}_{mes}"
                )
            config[(geo, ano)] = {"jornada": jornada, "empleados": empleados}

    st.subheader("\U0001F4C5 Define los rangos a comparar")
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

    umbral = st.number_input("\U0001F4CF Índice de absentismo objetivo (%)", min_value=0.0, max_value=100.0, value=saved_config.get("umbral", 4.0), step=0.1)

    if st.button("💾 Guardar configuración"):
        export_dict = {
            "geografias": geografias_seleccionadas,
            "funciones": funciones_seleccionadas,
            "codigos": codigos_seleccionados,
            "umbral": umbral,
            "config": {
                f"{geo}__{ano}": data for (geo, ano), data in config.items()
            }
        }
        config_bytes = BytesIO()
        config_bytes.write(json.dumps(export_dict, indent=2).encode('utf-8'))
        config_bytes.seek(0)
        st.download_button(
            "⬇️ Descargar configuración",
            data=config_bytes,
            file_name="configuracion_absentismo.json",
            mime="application/json"
        )

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
                for ano in subset['Año'].unique():
                    fig_line.add_scatter(
                        x=subset[(subset['Geografía'] == geo) & (subset['Año'] == ano)]['Mes_nombre'],
                        y=[umbral] * len(subset[(subset['Geografía'] == geo) & (subset['Año'] == ano)]),
                        mode='lines',
                        name=f'Objetivo {geo} - {ano}',
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
