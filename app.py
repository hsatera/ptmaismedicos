import pandas as pd
import streamlit as st
from io import StringIO

# --- Funções Auxiliares ---

def to_title_case(name):
    if pd.isna(name) or name == '':
        return ''
    return str(name).strip().title()

# --- Interface Streamlit ---
st.set_page_config(layout="wide", page_title="Análise Mais Médicos 2026")
st.title("📊 Sistema de Análise: Plano de Trabalho")

uploaded_file = st.file_uploader("Carregar arquivo (CSV ou XLSX)", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        # 1. Leitura do Arquivo (Detecta se é CSV ou Excel)
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, delimiter=';', encoding='latin1', skiprows=5)
            # Correção de encoding nas colunas
            df.columns = [col.encode('latin1').decode('utf-8', errors='ignore') for col in df.columns]
        else:
            df = pd.read_excel(uploaded_file, skiprows=5)
            df.columns = [str(col).strip() for col in df.columns]

        # 2. Filtragem e Limpeza (Conforme sua lógica original)
        required_cols = ['Município', 'Supervisor', 'Tutor']
        if not all(col in df.columns for col in required_cols):
            st.error(f"Colunas necessárias não encontradas. O arquivo contém: {list(df.columns)}")
            st.stop()

        df = df[required_cols].copy()
        df['Supervisor'] = df['Supervisor'].fillna('').apply(to_title_case)
        df['Tutor'] = df['Tutor'].fillna('').apply(to_title_case)
        df['Município'] = df['Município'].fillna('').apply(to_title_case)

        # Remover linhas onde Supervisor ou Tutor estão vazios
        df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

        # --- PROCESSAMENTO DOS DADOS ---

        # 1) Número de supervisores (nome único) por tutor
        supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()

        # 2) Número de médicos supervisionados por supervisor
        medicos_por_supervisor = df.groupby('Supervisor').size()

        # 3) Supervisores com menos de 8 ou mais de 10 médicos
        supervisores_fora_da_meta = medicos_por_supervisor[(medicos_por_supervisor < 8) | (medicos_por_supervisor > 10)]

        # 4) Supervisores com mais de 2 municípios
        supervisores_municipios = df.groupby('Supervisor')['Município'].nunique()
        supervisores_mais_de_2_mun = supervisores_municipios[supervisores_municipios > 2]

        # 5) Totais e Média
        total_medicos = len(df)
        total_supervisores = df['Supervisor'].nunique()
        media_medicos_por_supervisor = total_medicos / total_supervisores if total_supervisores > 0 else 0

        # 6) Relatório Detalhado por Município
        relatorio_municipios = []
        for municipio, grupo in df.groupby('Município'):
            supervisores = '; '.join(grupo['Supervisor'].unique())
            tutores = '; '.join(grupo['Tutor'].unique())
            relatorio_municipios.append(f"{municipio}. Supervisores: {supervisores}. Tutores: {tutores}.")
        relatorio_municipios_txt = "\n".join(relatorio_municipios)

        # --- EXIBIÇÃO NA INTERFACE ---
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Médicos", total_medicos)
        col2.metric("Total de Supervisores", total_supervisores)
        col3.metric("Média Méd/Sup", f"{media_medicos_por_supervisor:.2f}")

        st.divider()

        tab1, tab2, tab3 = st.tabs(["Alertas de Gestão", "Por Tutor", "Por Município"])
        
        with tab1:
            st.warning("⚠️ Supervisores fora da meta ( < 8 ou > 10 médicos)")
            st.dataframe(supervisores_fora_da_meta, use_container_width=True)
            
            st.error("🚨 Supervisores em mais de 2 municípios")
            st.dataframe(supervisores_mais_de_2_mun, use_container_width=True)

        with tab2:
            st.write("### Supervisores por Tutor")
            st.bar_chart(supervisores_por_tutor)
            st.dataframe(supervisores_por_tutor)

        with tab3:
            st.write("### Detalhamento por Município")
            st.text_area("Lista resumida", relatorio_municipios_txt, height=300)

        # --- GERAR ARQUIVO DE SAÍDA ---
        
        output_content = (
            "--- RELATÓRIO DE ANÁLISE DO PLANO DE TRABALHO ---\n\n"
            f"Número de supervisores por tutor:\n{supervisores_por_tutor.to_string()}\n\n"
            f"Número de médicos supervisionados por supervisor:\n{medicos_por_supervisor.to_string()}\n\n"
            f"Supervisores com menos de 8 ou mais de 10 médicos:\n{supervisores_fora_da_meta.to_string()}\n\n"
            f"Supervisores com mais de 2 municípios:\n{supervisores_mais_de_2_mun.to_string()}\n\n"
            f"Média de médicos por supervisor: {media_medicos_por_supervisor:.2f}\n\n"
            f"Relatório Detalhado por Município:\n{relatorio_municipios_txt}"
        )

        st.download_button(
            label="📥 Baixar Resultados (.txt)",
            data=output_content,
            file_name="analise_plano_trabalho.txt",
            mime="text/plain"
        )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
        st.exception(e)
