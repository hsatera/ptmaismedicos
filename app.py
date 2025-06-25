import streamlit as st
import pandas as pd
import tempfile
import os

st.set_page_config(page_title="Relatório de Supervisores", layout="wide")

st.title("📊 Relatório de Supervisores e Tutores")

uploaded_file = st.sidebar.file_uploader("Faça upload do arquivo XLS ou XLSX", type=["xls", "xlsx"])

def processar_dataframe(df):
    # Adjust columns to UTF-8 (if necessary)
    # Using errors='ignore' to prevent UnicodeDecodeError for non-string columns or those that can't be decoded
    df.columns = [str(col).encode('latin1').decode('utf-8', errors='ignore') if isinstance(col, str) else col for col in df.columns]

    # Select important columns
    # Ensure these columns exist before attempting to select them
    required_columns = ['Município', 'Supervisor', 'Tutor']
    for col in required_columns:
        if col not in df.columns:
            st.warning(f"A coluna '{col}' não foi encontrada no arquivo. Por favor, verifique se o cabeçalho está correto.")
            return pd.DataFrame() # Return empty DataFrame if critical columns are missing

    df = df[required_columns]

    # Fill NaN with empty string
    df['Supervisor'] = df['Supervisor'].fillna('')
    df['Tutor'] = df['Tutor'].fillna('')

    # Remove rows with empty supervisor or tutor
    df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

    return df

if uploaded_file is not None:
    try:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df = None # Initialize df

        if file_ext == 'xls':
            # Use xlrd for older .xls files
            try:
                df = pd.read_excel(uploaded_file, skiprows=5, engine='xlrd')
            except Exception as e:
                st.error(f"Erro ao ler o arquivo XLS com xlrd: {e}. O arquivo pode estar corrompido ou em um formato não suportado por xlrd.")
                st.stop()
        elif file_ext == 'xlsx':
            # Use openpyxl for .xlsx files
            try:
                df = pd.read_excel(uploaded_file, skiprows=5, engine='openpyxl')
            except Exception as e:
                st.error(f"Erro ao ler o arquivo XLSX com openpyxl: {e}. O arquivo pode estar corrompido.")
                st.stop()
        else:
            st.error("Formato não suportado. Por favor, envie um arquivo .xls ou .xlsx.")
            st.stop()

        if df is not None and not df.empty:
            # Process dataframe for cleaning and filtering
            df = processar_dataframe(df)

            if not df.empty:
                # Show loaded data
                st.subheader("🔍 Dados Carregados")
                st.dataframe(df)

                # Number of supervisors per tutor
                st.subheader("👩‍🏫 Número de Supervisores por Tutor")
                supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()
                st.dataframe(supervisores_por_tutor.reset_index().rename(columns={'Supervisor': 'Nº de Supervisores'}))

                # Number of supervised doctors per supervisor
                st.subheader("👨‍⚕️ Número de Médicos supervisionados por Supervisor")
                medicos_por_supervisor = df.groupby('Supervisor').size()
                st.dataframe(medicos_por_supervisor.reset_index().rename(columns={0: 'Nº de Médicos'}))

                # Supervisors with less than 8 or more than 10 doctors
                st.subheader("⚠️ Supervisores com menos de 8 ou mais de 10 médicos")
                supervisores_extremos = medicos_por_supervisor[(medicos_por_supervisor < 8) | (medicos_por_supervisor > 10)]
                st.dataframe(supervisores_extremos.reset_index().rename(columns={0: 'Nº de Médicos'}))

                # Supervisors with more than 2 municipalities
                st.subheader("🌎 Supervisores com mais de 2 municípios")
                supervisores_municipios = df.groupby('Supervisor')['Município'].nunique()
                supervisores_multi = supervisores_municipios[supervisores_municipios > 2]
                st.dataframe(supervisores_multi.reset_index().rename(columns={'Município': 'Nº de Municípios'}))

                # Average doctors per supervisor
                st.subheader("📈 Média de médicos por supervisor")
                total_medicos = df['Supervisor'].count()
                total_supervisores = df['Supervisor'].nunique()
                if total_supervisores > 0:
                    media_medicos = total_medicos / total_supervisores
                    st.metric("Média de médicos por supervisor", f"{media_medicos:.2f}")
                else:
                    st.info("Não há supervisores para calcular a média.")

                # Detailed report by municipality
                st.subheader("🗺️ Relatório Detalhado por Município")
                for municipio, grupo in df.groupby('Município'):
                    supervisores = '; '.join(sorted(grupo['Supervisor'].unique()))
                    tutores = '; '.join(sorted(grupo['Tutor'].unique()))
                    relatorio = f"**{municipio}**\n- Supervisores: {supervisores}\n- Tutores: {tutores}"
                    st.markdown(relatorio)
                    st.markdown("---")
            else:
                st.warning("O DataFrame resultante após o processamento está vazio. Verifique os dados no arquivo.")

    except Exception as e:
        st.error(f"Erro inesperado ao processar o arquivo: {e}")

else:
    st.info("Por favor, faça upload do arquivo XLS ou XLSX para gerar o relatório.")
