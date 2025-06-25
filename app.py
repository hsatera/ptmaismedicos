import streamlit as st
import pandas as pd
import pyexcel
import tempfile
import os

st.set_page_config(page_title="Relatório de Supervisores", layout="wide")

st.title("📊 Relatório de Supervisores e Tutores")

uploaded_file = st.sidebar.file_uploader("Faça upload do arquivo XLS ou XLSX", type=["xls", "xlsx"])

def processar_dataframe(df):
    # Ajusta colunas para UTF-8 (caso necessário)
    df.columns = [str(col).encode('latin1').decode('utf-8') if isinstance(col, str) else col for col in df.columns]

    # Seleciona colunas importantes
    df = df[['Município', 'Supervisor', 'Tutor']]

    # Preenche NaN com string vazia
    df['Supervisor'] = df['Supervisor'].fillna('')
    df['Tutor'] = df['Tutor'].fillna('')

    # Remove linhas com supervisor ou tutor vazios
    df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

    return df

if uploaded_file is not None:
    try:
        file_ext = uploaded_file.name.split('.')[-1].lower()

        if file_ext == 'xls':
            # Salvar temporariamente o arquivo .xls
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp_xls:
                tmp_xls.write(uploaded_file.read())
                tmp_xls_path = tmp_xls.name

            # Caminho para o arquivo convertido .xlsx
            tmp_xlsx_path = tmp_xls_path.replace(".xls", ".xlsx")

            # Converter .xls para .xlsx
            pyexcel.save_book_as(file_name=tmp_xls_path, dest_file_name=tmp_xlsx_path)

            # Ler o arquivo .xlsx convertido com pandas
            df = pd.read_excel(tmp_xlsx_path, skiprows=5, engine='openpyxl')

            # Apagar arquivos temporários
            os.remove(tmp_xls_path)
            os.remove(tmp_xlsx_path)

        elif file_ext == 'xlsx':
            df = pd.read_excel(uploaded_file, skiprows=5, engine='openpyxl')

        else:
            st.error("Formato não suportado. Por favor, envie um arquivo .xls ou .xlsx.")
            st.stop()

        # Processar dataframe para limpeza e filtros
        df = processar_dataframe(df)

        # Mostrar dados carregados
        st.subheader("🔍 Dados Carregados")
        st.dataframe(df)

        # Número de supervisores por tutor
        st.subheader("👩‍🏫 Número de Supervisores por Tutor")
        supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()
        st.dataframe(supervisores_por_tutor.reset_index().rename(columns={'Supervisor': 'Nº de Supervisores'}))

        # Número de médicos supervisionados por supervisor
        st.subheader("👨‍⚕️ Número de Médicos supervisionados por Supervisor")
        medicos_por_supervisor = df.groupby('Supervisor').size()
        st.dataframe(medicos_por_supervisor.reset_index().rename(columns={0: 'Nº de Médicos'}))

        # Supervisores com menos de 8 ou mais de 10 médicos
        st.subheader("⚠️ Supervisores com menos de 8 ou mais de 10 médicos")
        supervisores_extremos = medicos_por_supervisor[(medicos_por_supervisor < 8) | (medicos_por_supervisor > 10)]
        st.dataframe(supervisores_extremos.reset_index().rename(columns={0: 'Nº de Médicos'}))

        # Supervisores com mais de 2 municípios
        st.subheader("🌎 Supervisores com mais de 2 municípios")
        supervisores_municipios = df.groupby('Supervisor')['Município'].nunique()
        supervisores_multi = supervisores_municipios[supervisores_municipios > 2]
        st.dataframe(supervisores_multi.reset_index().rename(columns={'Município': 'Nº de Municípios'}))

        # Média de médicos por supervisor
        st.subheader("📈 Média de médicos por supervisor")
        total_medicos = df['Supervisor'].count()
        total_supervisores = df['Supervisor'].nunique()
        media_medicos = total_medicos / total_supervisores
        st.metric("Média de médicos por supervisor", f"{media_medicos:.2f}")

        # Relatório detalhado por município
        st.subheader("🗺️ Relatório Detalhado por Município")
        for municipio, grupo in df.groupby('Município'):
            supervisores = '; '.join(sorted(grupo['Supervisor'].unique()))
            tutores = '; '.join(sorted(grupo['Tutor'].unique()))
            relatorio = f"**{municipio}**\n- Supervisores: {supervisores}\n- Tutores: {tutores}"
            st.markdown(relatorio)
            st.markdown("---")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

else:
    st.info("Por favor, faça upload do arquivo XLS ou XLSX para gerar o relatório.")
