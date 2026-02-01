import pandas as pd
import streamlit as st
from io import StringIO
import altair as alt

# --- Funções Auxiliares ---

def read_excel_file(uploaded_file):
    try:
        # Removido o engine fixo para maior compatibilidade
        df = pd.read_excel(uploaded_file, skiprows=5)
        # Limpa espaços em branco nos nomes das colunas
        df.columns = [str(col).strip() for col in df.columns]
        return df
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None

def to_title_case(name):
    if pd.isna(name) or name == '':
        return ''
    return str(name).strip().title()

# --- Interface ---
st.set_page_config(layout="wide", page_title="Análise Plano de Trabalho")
st.title("📊 Análise de Plano de Trabalho")

uploaded_file = st.file_uploader("Carregar arquivo XLSX (Plano de Trabalho)", type=["xlsx"])

if uploaded_file is not None:
    df = read_excel_file(uploaded_file)

    if df is not None:
        try:
            # Lista de colunas esperadas
            required_columns = ['Município', 'Supervisor', 'Tutor', 'Nome Região', 'Início Atividades']
            
            # Verificação robusta de colunas
            if not all(col in df.columns for col in required_columns):
                st.error(f"Colunas esperadas não encontradas. Colunas lidas: {list(df.columns)}")
                st.stop() 

            # Seleção e Limpeza
            df = df[required_columns].copy()
            df['Supervisor'] = df['Supervisor'].apply(to_title_case)
            df['Tutor'] = df['Tutor'].apply(to_title_case)
            
            # Dropna em colunas cruciais antes de prosseguir
            df = df.dropna(subset=['Supervisor', 'Tutor', 'Início Atividades'])

            # Conversão de data
            df['Início Atividades'] = pd.to_datetime(df['Início Atividades'], errors='coerce')
            df = df.dropna(subset=['Início Atividades'])
            df['Ano_Mes'] = df['Início Atividades'].dt.strftime('%Y-%m') 

            # --- Lógica de Análise (Cálculos das Variáveis) ---
            
            total_medicos = len(df)
            total_supervisores = df['Supervisor'].nunique()
            total_tutores = df['Tutor'].nunique()
            
            # Cálculo da média com proteção contra divisão por zero
            media_medicos = total_medicos / total_supervisores if total_supervisores > 0 else 0

            # --- Visualização no Streamlit ---
            col1, col2, col3 = st.columns(3)
            col1.metric("Total de Médicos", total_medicos)
            col2.metric("Total de Supervisores", total_supervisores)
            col3.metric("Média Médicos/Supervisor", f"{media_medicos:.2f}")

            # --- Construção do Relatório de Texto ---
            report_text = "--- RELATÓRIO DE ANÁLISE - MAIS MÉDICOS ---\n\n"
            report_text += f"Data do Processamento: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}\n"
            report_text += f"Total de Médicos Analisados: {total_medicos}\n"
            report_text += f"Total de Supervisores: {total_supervisores}\n"
            report_text += f"Total de Tutores: {total_tutores}\n"
            report_text += f"Média de Médicos por Supervisor: {media_medicos:.2f}\n"
            report_text += "\n--- Distribuição por Região ---\n"
            
            # Adicionando contagem por região ao relatório
            regiao_counts = df['Nome Região'].value_counts()
            for regiao, count in regiao_counts.items():
                report_text += f"{regiao}: {count} médicos\n"

            st.write("### Resumo por Região")
            st.dataframe(regiao_counts)

            # Botão de Download
            st.download_button(
                label="📥 Baixar relatório completo (.txt)",
                data=report_text,
                file_name=f"relatorio_plano_trabalho.txt",
                mime="text/plain"
            )

        except Exception as e:
            st.error(f"Erro no processamento dos dados: {e}")
            st.exception(e)
