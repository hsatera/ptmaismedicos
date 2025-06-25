
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Relatório de Supervisores", layout="wide")

st.title("📊 Relatório de Supervisores e Tutores")

st.sidebar.header("⚙️ Upload de Arquivo")
uploaded_file = st.sidebar.file_uploader("Faça upload do arquivo XLS ou XLSX", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        if file_extension == 'xls':
            df = pd.read_excel(uploaded_file, skiprows=5, engine='xlrd')
        elif file_extension == 'xlsx':
            df = pd.read_excel(uploaded_file, skiprows=5)
        else:
            st.error("Formato não suportado. Por favor, envie um arquivo .xls ou .xlsx.")
            st.stop()

        df.columns = [str(col).encode('latin1').decode('utf-8') if isinstance(col, str) else col for col in df.columns]
        df = df[['Município', 'Supervisor', 'Tutor']]

        df['Supervisor'] = df['Supervisor'].fillna('')
        df['Tutor'] = df['Tutor'].fillna('')
        df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

        st.subheader("🔍 Dados Carregados")
        st.dataframe(df)

        st.subheader("👩‍🏫 Número de Supervisores por Tutor")
        supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()
        st.dataframe(supervisores_por_tutor.reset_index().rename(columns={'Supervisor': 'Nº de Supervisores'}))

        st.subheader("👨‍⚕️ Número de Médicos supervisionados por Supervisor")
        medicos_por_supervisor = df.groupby('Supervisor').size()
        st.dataframe(medicos_por_supervisor.reset_index().rename(columns={0: 'Nº de Médicos'}))

        st.subheader("⚠️ Supervisores com menos de 8 ou mais de 10 médicos")
        supervisores_extremos = medicos_por_supervisor[(medicos_por_supervisor < 8) | (medicos_por_supervisor > 10)]
        st.dataframe(supervisores_extremos.reset_index().rename(columns={0: 'Nº de Médicos'}))

        st.subheader("🌎 Supervisores com mais de 2 municípios")
        supervisores_municipios = df.groupby('Supervisor')['Município'].nunique()
        supervisores_multi = supervisores_municipios[supervisores_municipios > 2]
        st.dataframe(supervisores_multi.reset_index().rename(columns={'Município': 'Nº de Municípios'}))

        st.subheader("📈 Média de médicos por supervisor")
        total_medicos = df['Supervisor'].count()
        total_supervisores = df['Supervisor'].nunique()
        media_medicos = total_medicos / total_supervisores
        st.metric("Média de médicos por supervisor", f"{media_medicos:.2f}")

        st.subheader("🗺️ Relatório Detalhado por Município")
        relatorio_municipios = []

        for municipio, grupo in df.groupby('Município'):
            supervisores = '; '.join(sorted(grupo['Supervisor'].unique()))
            tutores = '; '.join(sorted(grupo['Tutor'].unique()))
            relatorio = f"**{municipio}**\n- Supervisores: {supervisores}\n- Tutores: {tutores}"
            relatorio_municipios.append(relatorio)

        for item in relatorio_municipios:
            st.markdown(item)
            st.markdown("---")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

else:
    st.info("Por favor, faça upload do arquivo XLS ou XLSX para gerar o relatório.")
