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
st.title("Análise de Plano de Trabalho")

uploaded_file = st.file_uploader("Carregar arquivo XLSX", type=["xlsx"])

if uploaded_file is not None:
    df = read_excel_file(uploaded_file)

    if df is not None:
        try:
            # Lista de colunas esperadas
            required_columns = ['Município', 'Supervisor', 'Tutor', 'Nome Região', 'Início Atividades']
            
            # Verificação robusta de colunas
            if not all(col in df.columns for col in required_columns):
                st.error(f"Colunas encontradas: {list(df.columns)}")
                st.stop() # Interrompe a execução aqui para mostrar o erro

            # Seleção e Limpeza
            df = df[required_columns].copy()
            df['Supervisor'] = df['Supervisor'].apply(to_title_case)
            df['Tutor'] = df['Tutor'].apply(to_title_case)
            
            # Dropna em colunas cruciais antes de prosseguir
            df = df.dropna(subset=['Supervisor', 'Tutor', 'Início Atividades'])

            # Conversão de data
            df['Início Atividades'] = pd.to_datetime(df['Início Atividades'], errors='coerce')
            df = df.dropna(subset=['Início Atividades'])
            df['Ano_Mes'] = df['Início Atividades'].dt.strftime('%Y-%m') # Formato string fixo

            # --- (Restante da sua lógica de análise permanece igual) ---
            # ... (Cálculos de médias e agrupamentos) ...

            # DICA: Para o download_button, use encode se der erro de Buffer
            report_text = "--- Relatório de Análise ---\n\n"
            report_text += f"Média de Médicos: {total_medicos/total_supervisores:.2f}"
            # ... adicione o resto ao report_text ...

            st.download_button(
                label="Baixar relatório completo",
                data=report_text,
                file_name="relatorio.txt",
                mime="text/plain"
            )

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
            st.exception(e) # Isso mostra o traceback completo para você debugar
