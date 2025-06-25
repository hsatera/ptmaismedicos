import pandas as pd
import streamlit as st
from io import StringIO

# Função para ler o arquivo Excel (.xls ou .xlsx)
def read_excel_file(uploaded_file):
    try:
        # Tenta ler o arquivo, pulando as 5 primeiras linhas (0 a 4)
        df = pd.read_excel(uploaded_file, skiprows=5, engine=None)
        return df
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None


# ---- Interface Streamlit ----
st.title("Análise de Plano de Trabalho")

uploaded_file = st.file_uploader("Carregar arquivo XLS ou XLSX", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Ler o arquivo
    df = read_excel_file(uploaded_file)

    if df is not None:
        st.subheader("Pré-visualização dos Dados")
        st.dataframe(df.head())

        try:
            # Seleciona colunas de interesse
            df = df[['Município', 'Supervisor', 'Tutor']]

            # Limpa dados vazios
            df['Supervisor'] = df['Supervisor'].fillna('')
            df['Tutor'] = df['Tutor'].fillna('')
            df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

            # 1) Número de supervisores únicos por tutor
            supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()

            # 2) Número de médicos supervisionados por supervisor
            medicos_por_supervisor = df.groupby('Supervisor').size()

            # 3) Supervisores com menos de 8 ou mais de 10 médicos
            supervisores_menos_de_8_ou_mais_de_10 = medicos_por_supervisor[
                (medicos_por_supervisor < 8) | (medicos_por_supervisor > 10)
            ]

            # 4) Supervisores que atuam em mais de 2 municípios
            supervisores_mais_de_2_municipios = df.groupby('Supervisor')['Município'].nunique()
            supervisores_mais_de_2_municipios = supervisores_mais_de_2_municipios[
                supervisores_mais_de_2_municipios > 2
            ]

            # 5) Média de médicos por supervisor
            total_medicos = df['Supervisor'].count()
            total_supervisores = df['Supervisor'].nunique()
            media_medicos_por_supervisor = total_medicos / total_supervisores

            # 6) Relatório detalhado por município
            relatorio_municipios = []
            for municipio, grupo in df.groupby('Município'):
                supervisores = '; '.join(grupo['Supervisor'].unique())
                tutores = '; '.join(grupo['Tutor'].unique())
                relatorio = f"{municipio}. Supervisores: {supervisores}. Tutores: {tutores}."
                relatorio_municipios.append(relatorio)
            relatorio_municipios_txt = "\n".join(relatorio_municipios)

            # ---- Exibir resultados ----
            st.subheader("Número de supervisores por tutor:")
            st.write(supervisores_por_tutor)

            st.subheader("Número de médicos supervisionados por supervisor:")
            st.write(medicos_por_supervisor)

            st.subheader("Supervisores com menos de 8 ou mais de 10 médicos:")
            st.write(supervisores_menos_de_8_ou_mais_de_10)

            st.subheader("Supervisores com mais de 2 municípios:")
            st.write(supervisores_mais_de_2_municipios)

            st.subheader("Média de médicos por supervisor:")
            st.write(f"{media_medicos_por_supervisor:.2f}")

            st.subheader("Relatório Detalhado por Município:")
            st.text(relatorio_municipios_txt)

            # ---- Gerar arquivo para download ----
            output = StringIO()
            output.write("Número de supervisores por tutor:\n")
            output.write(supervisores_por_tutor.to_string())
            output.write("\n\nNúmero de médicos supervisionados por supervisor:\n")
            output.write(medicos_por_supervisor.to_string())
            output.write("\n\nSupervisores com menos de 8 ou mais de 10 médicos:\n")
            output.write(supervisores_menos_de_8_ou_mais_de_10.to_string())
            output.write("\n\nSupervisores com mais de 2 municípios:\n")
            output.write(supervisores_mais_de_2_municipios.to_string())
            output.write("\n\nMédia de médicos por supervisor:\n")
            output.write(f"{media_medicos_por_supervisor:.2f}\n")
            output.write("\n\nRelatório Detalhado por Município:\n")
            output.write(relatorio_municipios_txt)

            st.download_button(
                label="Baixar relatório completo",
                data=output.getvalue(),
                file_name="relatorio_supervisao.txt",
                mime="text/plain"
            )

        except Exception as e:
            st.error(f"Ocorreu um erro no processamento dos dados: {e}")
