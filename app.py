import pandas as pd
import streamlit as st
from io import StringIO
import altair as alt # Importação da biblioteca Altair para gráficos mais avançados

# --- Funções Auxiliares ---

# Função para ler o arquivo Excel (.xls ou .xlsx)
# Tenta ler o arquivo, pulando as 5 primeiras linhas (índices 0 a 4)
def read_excel_file(uploaded_file):
    try:
        # Para arquivos .xlsx, 'openpyxl' é o motor recomendado.
        # Para .xls, 'xlrd' seria o motor, mas o tipo aceito é apenas .xlsx
        df = pd.read_excel(uploaded_file, skiprows=5, engine='openpyxl')
        return df
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel. Verifique o formato e a integridade do arquivo. Detalhes: {e}")
        return None

# Função para converter nomes para o formato Title Case (Primeira Letra Maiúscula de cada palavra)
def to_title_case(name):
    if pd.isna(name): # Verifica se o valor é NaN (Not a Number) ou vazio/nulo
        return ''
    return str(name).title() # Converte para string e aplica Title Case

# --- Interface Streamlit ---
st.set_page_config(layout="wide") # Define o layout da página para largura total, otimizando o espaço
st.title("Análise de Plano de Trabalho")

# Widget para upload do arquivo XLSX
uploaded_file = st.file_uploader("Carregar arquivo XLSX", type=["xlsx"])

if uploaded_file is not None:
    # Tenta ler o arquivo carregado
    df = read_excel_file(uploaded_file)

    if df is not None:
        st.subheader("Pré-visualização dos Dados Carregados")
        # Exibe apenas as primeiras 5 linhas da planilha para uma rápida verificação
        st.dataframe(df.head(), use_container_width=True)

        try:
            # Define as colunas essenciais para a análise
            required_columns = ['Município', 'Supervisor', 'Tutor', 'Nome Região', 'Início Atividades']
            
            # Verifica se todas as colunas necessárias estão presentes no DataFrame
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                raise KeyError(f"Colunas ausentes no arquivo: {', '.join(missing_cols)}. Por favor, verifique se o arquivo possui todas as colunas esperadas.")

            # Seleciona apenas as colunas de interesse e cria uma cópia para evitar SettingWithCopyWarning
            df = df[required_columns].copy()

            # Aplica a correção de capitalização aos nomes dos tutores e supervisores
            df['Supervisor'] = df['Supervisor'].apply(to_title_case)
            df['Tutor'] = df['Tutor'].apply(to_title_case)

            # Limpa dados vazios nas colunas 'Supervisor' e 'Tutor'
            # Substitui valores NaN por string vazia para facilitar a filtragem
            df['Supervisor'] = df['Supervisor'].fillna('')
            df['Tutor'] = df['Tutor'].fillna('')
            
            # Filtra linhas onde 'Supervisor' ou 'Tutor' estão vazios, garantindo dados válidos
            df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

            # Converte a coluna de data para o tipo datetime, tratando erros com 'coerce'
            # 'dayfirst=True' é útil para formatos de data brasileiros (DD/MM/AAAA)
            df['Início Atividades'] = pd.to_datetime(df['Início Atividades'], errors='coerce', dayfirst=True)
            
            # Remove linhas onde a conversão de data falhou (resultando em NaT - Not a Time)
            df.dropna(subset=['Início Atividades'], inplace=True)

            # Cria uma coluna 'Ano_Mes' para agrupar os dados por mês e ano
            df['Ano_Mes'] = df['Início Atividades'].dt.to_period('M').astype(str)

            # --- Análises de Dados ---

            # 1) Número de supervisores únicos por tutor
            # Agrupa por 'Tutor' e conta o número de supervisores únicos para cada um
            supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()

            # 2) Número de médicos supervisionados por supervisor
            # Conta o número de ocorrências de cada supervisor (cada linha representa um médico)
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

            # 6) Quantidade de médicos por mês e ano (usando 'Início Atividades')
            medicos_por_mes_ano = df.groupby('Ano_Mes').size().sort_index()

            # Calcula a quantidade cumulativa de médicos por mês e ano
            medicos_cumulativos = medicos_por_mes_ano.cumsum()

            # Calcula a porcentagem cumulativa de médicos em relação ao total geral
            total_medicos_geral = medicos_cumulativos.iloc[-1] if not medicos_cumulativos.empty else 1
            porcentagem_cumulativa = (medicos_cumulativos / total_medicos_geral) * 100

            # Combina os dados para o gráfico cumulativo em um DataFrame para o Altair
            df_cumulativo_chart = pd.DataFrame({
                'Médicos Cumulativos': medicos_cumulativos,
                'Percentual Cumulativo (%)': porcentagem_cumulativa
            }).fillna(0) # Preenche NaNs com 0 caso haja meses sem dados

            # --- Exibir resultados da Análise Geral ---
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

            # --- Gráfico de Quantidade de Médicos por Mês e Ano (Cumulativo em Barras com Porcentagem) ---
            st.subheader("Quantidade de Médicos por Mês e Ano (Cumulativo)")
            if not df_cumulativo_chart.empty:
                # Criar um DataFrame para o Altair que inclui o índice 'Ano_Mes'
                df_chart_reset = df_cumulativo_chart.reset_index()

                # Criar uma coluna formatada para a dica de ferramenta do percentual para melhor visualização
                df_chart_reset['Percentual Cumulativo Formatado'] = df_chart_reset['Percentual Cumulativo (%)'].apply(lambda x: f"{x:.2f}%")

                # Criar o gráfico de barras para Médicos Cumulativos
                bar_chart = alt.Chart(df_chart_reset).mark_bar().encode(
                    x=alt.X('Ano_Mes', sort=None, title='Mês/Ano de Início das Atividades'),
                    y=alt.Y('Médicos Cumulativos', title='Médicos Cumulativos', axis=alt.Axis(titleColor='#5276A7')),
                    tooltip=['Ano_Mes', 'Médicos Cumulativos']
                )

                # Criar o gráfico de linha para Percentual Cumulativo (%)
                line_chart = alt.Chart(df_chart_reset).mark_line(color='red').encode(
                    x=alt.X('Ano_Mes', sort=None), # Compartilha o eixo X com o gráfico de barras
                    y=alt.Y('Percentual Cumulativo (%)', title='Percentual Cumulativo (%)', axis=alt.Axis(titleColor='red')),
                    tooltip=['Ano_Mes', 'Percentual Cumulativo Formatado'] # Usando a coluna formatada para o tooltip
                )

                # Combinar os gráficos e configurar eixos Y independentes
                chart = alt.layer(bar_chart, line_chart).resolve_scale(
                    y='independent' # Permite que cada gráfico tenha seu próprio eixo Y
                ).properties(
                    title='Médicos Cumulativos por Mês de Chegada e Percentual de Crescimento'
                ).interactive() # Ativa zoom e pan para uma melhor exploração do gráfico

                st.altair_chart(chart, use_container_width=True)
                st.dataframe(df_cumulativo_chart, use_container_width=True)
            else:
                st.info("Não há dados de data válidos para gerar o gráfico cumulativo de médicos por mês e ano. Verifique a coluna 'Início Atividades'.")

            # --- Gráfico de Barras Empilhadas de Médicos por Município com Linha Cumulativa (em %) ---
            st.subheader("Número de Médicos por Município (Maiores para Menores) com Linha Cumulativa em %")
            
            # Contar médicos por município e ordenar
            medicos_por_municipio = df.groupby('Município').size().sort_values(ascending=False)
            
            # Calcular o número cumulativo de médicos
            medicos_cumulativos_municipio = medicos_por_municipio.cumsum()
            
            # Calcular a porcentagem cumulativa de médicos em relação ao total geral de médicos por município
            total_medicos_municipios = medicos_cumulativos_municipio.iloc[-1] if not medicos_cumulativos_municipio.empty else 1
            porcentagem_cumulativa_municipio = (medicos_cumulativos_municipio / total_medicos_municipios) * 100

            # Criar um DataFrame para o Altair
            df_municipio_chart = pd.DataFrame({
                'Município': medicos_por_municipio.index,
                'Número de Médicos': medicos_por_municipio.values,
                'Médicos Cumulativos Percentual (%)': porcentagem_cumulativa_municipio.values
            })

            if not df_municipio_chart.empty:
                # Criar uma coluna formatada para a dica de ferramenta do percentual para melhor visualização
                df_municipio_chart['Médicos Cumulativos Percentual Formatado'] = df_municipio_chart['Médicos Cumulativos Percentual (%)'].apply(lambda x: f"{x:.2f}%")

                # Criar o gráfico de barras
                bar_chart_municipio = alt.Chart(df_municipio_chart).mark_bar().encode(
                    x=alt.X('Município', sort='-y', title='Município'), # Ordenar pelo número de médicos
                    y=alt.Y('Número de Médicos', title='Número de Médicos', axis=alt.Axis(titleColor='#5276A7')),
                    tooltip=['Município', 'Número de Médicos']
                )

                # Criar o gráfico de linha cumulativa em percentual
                line_chart_municipio = alt.Chart(df_municipio_chart).mark_line(color='red').encode(
                    x=alt.X('Município', sort='-y'), # Compartilhar o eixo X
                    y=alt.Y('Médicos Cumulativos Percentual (%)', title='Médicos Cumulativos (%)', axis=alt.Axis(titleColor='red')),
                    tooltip=['Município', 'Médicos Cumulativos Percentual Formatado'] # Usando a coluna formatada para o tooltip
                )

                # Combinar os gráficos e configurar eixos Y independentes
                chart_municipio = alt.layer(bar_chart_municipio, line_chart_municipio).resolve_scale(
                    y='independent'
                ).properties(
                    title='Distribuição e Acúmulo Percentual de Médicos por Município'
                ).interactive()

                st.altair_chart(chart_municipio, use_container_width=True)
                st.dataframe(df_municipio_chart, use_container_width=True)
            else:
                st.info("Não há dados de município válidos para gerar o gráfico de médicos por município. Verifique a coluna 'Município'.")


            # --- Relatório Detalhado por Região e Município ---
            st.subheader("Relatório Detalhado por Região e Município")

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

            # --- Gerar arquivo para download ---
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
            output.write("\n\n6. Quantidade de Médicos por Mês e Ano (Cumulativo):\n")
            # Adiciona o DataFrame completo ao arquivo de saída
            output.write(df_cumulativo_chart.to_string())
            output.write("\n\n7. Relatório Detalhado por Região e Município:\n")
            # Usa to_string para o DataFrame completo, sem o índice para uma saída mais limpa
            output.write(relatorio_final_df.to_string(index=False))

            st.download_button(
                label="Baixar relatório completo",
                data=output.getvalue(),
                file_name="relatorio_supervisao.txt",
                mime="text/plain"
            )

        except KeyError as ke:
            st.error(f"Erro: Coluna '{ke}' não encontrada no arquivo. Verifique se o arquivo possui as colunas esperadas: 'Município', 'Supervisor', 'Tutor', 'Nome Região', 'Início Atividades'.")
        except Exception as e:
            st.error(f"Ocorreu um erro inesperado no processamento dos dados: {e}")

