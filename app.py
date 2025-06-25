import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Relatório de Supervisores", layout="wide")

st.title("📊 Relatório de Supervisores e Tutores")

st.sidebar.header("⚙️ Upload de Arquivo")
uploaded_file = st.sidebar.file_uploader("Faça upload do arquivo XLS ou XLSX", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        # Ler o Excel pulando as primeiras 5 linhas
        file_extension = uploaded_file.name.split('.')[-1].lower()
        if file_extension == 'xls':
            df_excel = pd.read_excel(uploaded_file, skiprows=5, engine='xlrd')
        elif file_extension == 'xlsx':
            df_excel = pd.read_excel(uploaded_file, skiprows=5)
        else:
            st.error("Formato não suportado. Por favor, envie um arquivo .xls ou .xlsx.")
            st.stop()

        # Converter para CSV na memória usando ';' como delimitador e codificação utf-8
        csv_buffer = io.StringIO()
        df_excel.to_csv(csv_buffer, sep=';', index=False, encoding='utf-8')
        csv_buffer.seek(0)

        # Ler CSV da memória com encoding latin1 (para simular seu script original)
        df = pd.read_csv(csv_buffer, delimiter=';', encoding='utf-8')

        # Ajuste para colunas com acentuação — usar decode para simular seu método
        df.columns = [str(col).encode('latin1').decode('utf-8') if isinstance(col, str) else col for col in df.columns]

        # Filtrar as colunas importantes
        df = df[['Município', 'Supervisor', 'Tutor']]

        # Preencher NaN
        df['Supervisor'] = df['Supervisor'].fillna('')
        df['Tutor'] = df['Tutor'].fillna('')
        df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

        # 1) Número de supervisores por tutor
        supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()

        # 2) Número de médicos supervisionados por supervisor
        medicos_por_supervisor = df.groupby('Supervisor').size()

        # 3) Supervisores com menos de 8 ou mais de 10 médicos
        supervisores_extremos = medicos_por_supervisor[(medicos_por_supervisor < 8) | (medicos_por_supervisor > 10)]

        # 4) Supervisores com mais de 2 municípios
        supervisores_municipios = df.groupby('Supervisor')['Município'].nunique()
        supervisores_multi = supervisores_municipios[supervisores_municipios > 2]

        # 5) Média de médicos por supervisor
        total_medicos = df['Supervisor'].count()
        total_supervisores = df['Supervisor'].nunique()
        media_medicos = total_medicos / total_supervisores if total_supervisores else 0

        # 6) Relatório detalhado por município
        relatorio_municipios = []
        for municipio, grupo in df.groupby('Município'):
            supervisores = '; '.join(sorted(grupo['Supervisor'].unique()))
            tutores = '; '.join(sorted(grupo['Tutor'].unique()))
            relatorio = f"{municipio}. Supervisores: {supervisores}. Tutores: {tutores}."
            relatorio_municipios.append(relatorio)

        # Mostrar tudo formatado no Streamlit
        st.subheader("Número de supervisores por tutor:")
        st.dataframe(supervisores_por_tutor.reset_index().rename(columns={'Supervisor': 'Nº de Supervisores'}))

        st.subheader("Número de médicos supervisionados por supervisor:")
        st.dataframe(medicos_por_supervisor.reset_index().rename(columns={0: 'Nº de Médicos'}))

        st.subheader("Supervisores com menos de 8 ou mais de 10 médicos:")
        st.dataframe(supervisores_extremos.reset_index().rename(columns={0: 'Nº de Médicos'}))

        st.subheader("Supervisores com mais de 2 municípios:")
        st.dataframe(supervisores_multi.reset_index().rename(columns={'Município': 'Nº de Municípios'}))

        st.subheader("Média de médicos por supervisor:")
        st.metric(label="Média de médicos por supervisor", value=f"{media_medicos:.2f}")

        st.subheader("Relatório Detalhado por Município:")
        for item in relatorio_municipios:
            st.markdown(f"- {item}")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

else:
    st.info("Por favor, faça upload do arquivo XLS ou XLSX para gerar o relatório.")
