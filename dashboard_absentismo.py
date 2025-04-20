import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard Absentismo Simplificado", layout="wide")
st.title("📊 Absentismo: 1 baja = jornada mensual ÷ 28")

uploaded_file = st.file_uploader("Sube el archivo Excel con los datos de ausencias", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Inicio'] = pd.to_datetime(df['Inicio'])
    df['Año'] = df['Inicio'].dt.year
    df['Mes'] = df['Inicio'].dt.month
    df['Mes_nombre'] = df['Inicio'].dt.strftime('%B')

    geografias = sorted(df['Geografía'].dropna().unique())
    geografias_seleccionadas = st.multiselect("Selecciona geografía(s):", geografias, default=geografias)

    funciones_disponibles = sorted(df['Función'].dropna().unique())
    funciones_seleccionadas = st.multiselect("Selecciona función(es):", funciones_disponibles, default=funciones_disponibles)

    codigos_disponibles = sorted(df['Codigo'].dropna().unique())
    codigos_seleccionados = st.multiselect("Selecciona códigos de ausencia:", codigos_disponibles, default=codigos_disponibles)

    st.sidebar.header("⚙️ Configuración por geografía")
    configuracion = {}
    for geo in geografias_seleccionadas:
        st.sidebar.subheader(f"🌍 {geo}")
        jornada_fija = st.sidebar.number_input(f"Jornada mensual {geo} (h)", min_value=1, value=140, step=1, key=f"jornada_{geo}")
        empleados_por_mes = {}
        for mes in range(1, 13):
            nombre_mes = datetime(2023, mes, 1).strftime('%B')
            empleados = st.sidebar.number_input(f"{geo} - {nombre_mes} - Empleados", min_value=0, value=100, step=1, key=f"{geo}_{mes}")
            empleados_por_mes[mes] = empleados
        configuracion[geo] = {
            "jornada_mensual": jornada_fija,
            "empleados_mes": empleados_por_mes
        }

    st.subheader("📆 Selecciona rango de fechas")
    fecha_min = df['Inicio'].min()
    fecha_max = df['Inicio'].max()
    rango = st.date_input("Rango", [fecha_min, fecha_max], min_value=fecha_min, max_value=fecha_max)

    if len(rango) == 2:
        inicio_rango, fin_rango = pd.to_datetime(rango[0]), pd.to_datetime(rango[1])

        df_filtrado = df[
            (df['Inicio'] >= inicio_rango) &
            (df['Inicio'] <= fin_rango) &
            (df['Geografía'].isin(geografias_seleccionadas)) &
            (df['Codigo'].isin(codigos_seleccionados)) &
            (df['Función'].isin(funciones_seleccionadas))
        ]

        resumen_total = pd.DataFrame()

        for geo in geografias_seleccionadas:
            df_geo = df_filtrado[df_filtrado['Geografía'] == geo]
            resumen = df_geo.groupby('Mes').size().reset_index(name='Bajas')
            resumen['Geografía'] = geo
            resumen['Mes_nombre'] = resumen['Mes'].apply(lambda m: datetime(2023, m, 1).strftime('%B'))

            resumen['Horas por baja'] = configuracion[geo]['jornada_mensual'] / 28
            resumen['Horas de ausencia'] = resumen['Bajas'] * resumen['Horas por baja']
            resumen['Horas teóricas'] = resumen['Mes'].apply(
                lambda m: configuracion[geo]['empleados_mes'].get(m, 0) * configuracion[geo]['jornada_mensual']
            )
            resumen['Absentismo (%)'] = (resumen['Horas de ausencia'] / resumen['Horas teóricas']) * 100
            resumen['Absentismo (%)'] = resumen['Absentismo (%)'].round(2)

            resumen_total = pd.concat([resumen_total, resumen], ignore_index=True)

        st.subheader("📊 Gráfico de Absentismo (%)")
        fig = px.bar(
            resumen_total,
            x='Mes_nombre',
            y='Absentismo (%)',
            color='Geografía',
            barmode='group',
            text='Absentismo (%)',
            title='Absentismo por Mes y Geografía'
        )
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Detalle")
        st.dataframe(resumen_total[['Geografía', 'Mes_nombre', 'Bajas', 'Horas por baja', 'Horas de ausencia', 'Horas teóricas', 'Absentismo (%)']])

        if st.button("📥 Exportar a Excel"):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                resumen_total.to_excel(writer, index=False, sheet_name='Resumen')
            st.download_button(
                label="📂 Descargar Excel",
                data=buffer.getvalue(),
                file_name="absentismo_simplificado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

