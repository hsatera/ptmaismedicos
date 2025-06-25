import pandas as pd
import streamlit as st
from io import StringIO

# Função para ler o arquivo Excel (.xls ou .xlsx)
# Tenta ler o arquivo, pulando as 5 primeiras linhas (índices 0 a 4)
def read_excel_file(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, skiprows=5, engine=None)
        return df
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None

# ---- Interface Streamlit ----
st.set_page_config(layout="wide") # Define o layout da página para largura total
st.title("Análise de Plano de Trabalho")

uploaded_file = st.file_uploader("Carregar arquivo XLS ou XLSX", type=["xlsx"])

if uploaded_file is not None:
    # Ler o arquivo
    df = read_excel_file(uploaded_file)

    if df is not None:
        st.subheader("Pré-visualização dos Dados Carregados")
        # Exibe apenas as primeiras linhas da planilha, conforme solicitado
        st.dataframe(df.head(), use_container_width=True)

        try:
            # Seleciona as colunas de interesse, incluindo 'Nome Região'
            # Assegura que as colunas 'Município', 'Supervisor', 'Tutor' e 'Nome Região' existem
            required_columns = ['Município', 'Supervisor', 'Tutor', 'Nome Região']
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                raise KeyError(f"Colunas ausentes no arquivo: {', '.join(missing_cols)}. Verifique se o arquivo possui as colunas esperadas.")

            df = df[required_columns]

            # Normaliza a capitalização das colunas de texto para Title Case
            # Primeiro converte para minúsculas, depois aplica o Title Case
            for col in ['Município', 'Supervisor', 'Tutor', 'Nome Região']:
                if col in df.columns:
                    df[col] = df[col].astype(str).apply(lambda x: x.lower().title())

            # Limpa dados vazios nas colunas 'Supervisor' e 'Tutor'
            df['Supervisor'] = df['Supervisor'].fillna('')
            df['Tutor'] = df['Tutor'].fillna('')
            # Filtra linhas onde 'Supervisor' ou 'Tutor' estão vazios
            df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

            # 1) Número de supervisores únicos por tutor
            # Agrupa por 'Tutor' e conta o número de supervisores únicos para cada um
            supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()

            # 2) Número de médicos supervisionados por supervisor
            # Conta o número de ocorrências de cada supervisor (cada linha é um médico)
            medicos_por_supervisor = df.groupby('Supervisor').size()

            # 3) Supervisores com menos de 8 ou mais de 10 médicos
            # Filtra os supervisores com base no número de médicos supervisionados
            supervisores_menos_de_8_ou_mais_de_10 = medicos_por_supervisor[
                (medicos_por_supervisor < 8) | (medicos_por_supervisor > 10)
            ]

            # 4) Supervisores que atuam em mais de 2 municípios
            # Agrupa por 'Supervisor' e conta o número de municípios únicos para cada um
            supervisores_mais_de_2_municipios = df.groupby('Supervisor')['Município'].nunique()
            # Filtra os supervisores que atuam em mais de 2 municípios
            supervisores_mais_de_2_municipios = supervisores_mais_de_2_municipios[
                supervisores_mais_de_2_municipios > 2
            ]

            # 5) Média de médicos por supervisor
            # Calcula o total de médicos (linhas após a limpeza) e o total de supervisores únicos
            total_medicos = df.shape[0]
            total_supervisores = df['Supervisor'].nunique()
            # Evita divisão por zero
            media_medicos_por_supervisor = total_medicos / total_supervisores if total_supervisores > 0 else 0

            # ---- Exibir resultados da Análise Geral ----
            st.subheader("Análise Geral")

            st.write("### Número de supervisores únicos por tutor:")
            st.dataframe(supervisores_por_tutor.reset_index(name='Total Supervisores'), use_container_width=True)

            st.write("### Número de médicos supervisionados por supervisor:")
            st.dataframe(medicos_por_supervisor.reset_index(name='Total Médicos'), use_container_width=True)

            st.write("### Supervisores com menos de 8 ou mais de 10 médicos:")
            st.dataframe(supervisores_menos_de_8_ou_mais_de_10.reset_index(name='Total Médicos'), use_container_width=True)

            st.write("### Supervisores que atuam em mais de 2 municípios:")
            st.dataframe(supervisores_mais_de_2_municipios.reset_index(name='Total Municípios'), use_container_width=True)

            st.write("### Média de médicos por supervisor:")
            st.write(f"**{media_medicos_por_supervisor:.2f}** médicos por supervisor")

            # ---- Relatório Detalhado por Região e Município (Formato Corrigido) ----
            st.subheader("Relatório Detalhado por Região e Município")

            # Agrupa por 'Nome Região' e 'Município' para obter tutores e supervisores únicos
            relatorio_df_list = []
            # Itera sobre cada grupo de Região e Município
            for (regiao, municipio), grupo in df.groupby(['Nome Região', 'Município']):
                # Coleta tutores únicos e os une em uma string, ordenados alfabeticamente
                unique_tutores = '; '.join(sorted(grupo['Tutor'].unique()))
                # Coleta supervisores únicos e os une em uma string, ordenados alfabeticamente
                unique_supervisores = '; '.join(sorted(grupo['Supervisor'].unique()))
                # Adiciona os dados à lista
                relatorio_df_list.append({
                    'Nome Região': regiao,
                    'Município': municipio,
                    'Tutores': unique_tutores,
                    'Supervisores': unique_supervisores
                })
            # Cria um DataFrame a partir da lista de dicionários
            relatorio_final_df = pd.DataFrame(relatorio_df_list)
            # Exibe o DataFrame formatado como tabela
            st.dataframe(relatorio_final_df, use_container_width=True)


            # ---- Gerar arquivo para download ----
            output = StringIO()
            output.write("--- Relatório de Análise de Plano de Trabalho ---\n\n")

            output.write("1. Número de supervisores únicos por tutor:\n")
            output.write(supervisores_por_tutor.to_string())
            output.write("\n\n2. Número de médicos supervisionados por supervisor:\n")
            output.write(medicos_por_supervisor.to_string())
            output.write("\n\n3. Supervisores com menos de 8 ou mais de 10 médicos:\n")
            output.write(supervisores_menos_de_8_ou_mais_de_10.to_string())
            output.write("\n\n4. Supervisores que atuam em mais de 2 municípios:\n")
            output.write(supervisores_mais_de_2_municipios.to_string())
            output.write("\n\n5. Média de médicos por supervisor:\n")
            output.write(f"{media_medicos_por_supervisor:.2f}\n")
            output.write("\n\n6. Relatório Detalhado por Região e Município:\n")
            # Usa to_string para o DataFrame completo, sem o índice
            output.write(relatorio_final_df.to_string(index=False))

            st.download_button(
                label="Baixar relatório completo",
                data=output.getvalue(),
                file_name="relatorio_supervisao.txt",
                mime="text/plain"
            )

        except KeyError as ke:
            st.error(f"Erro: Coluna '{ke}' não encontrada no arquivo. Verifique se o arquivo possui as colunas esperadas: 'Município', 'Supervisor', 'Tutor', 'Nome Região'.")
        except Exception as e:
            st.error(f"Ocorreu um erro no processamento dos dados: {e}")
