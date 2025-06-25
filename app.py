import pandas as pd
import streamlit as st
from io import BytesIO
import altair as alt
from fpdf import FPDF # Importa a biblioteca FPDF para geração de PDF

# Função para ler o arquivo Excel (.xls ou .xlsx)
# Tenta ler o arquivo, pulando as 5 primeiras linhas (índices 0 a 4)
def read_excel_file(uploaded_file):
    try:
        # Especifica o motor 'openpyxl' para arquivos .xlsx
        df = pd.read_excel(uploaded_file, skiprows=5, engine='openpyxl')
        return df
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None

# Função para converter nomes para o formato Title Case (Primeira Letra Maiúscula)
def to_title_case(name):
    if pd.isna(name):
        return ''
    return str(name).title()

# Função para gerar o relatório em PDF
def generate_pdf_report(supervisores_por_tutor, medicos_por_supervisor, supervisores_menos_de_8_ou_mais_de_10,
                        supervisores_mais_de_2_municipios, media_medicos_por_supervisor, df_cumulativo_chart, relatorio_final_df):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Título do Relatório
    pdf.cell(200, 10, txt="--- Relatório de Análise de Plano de Trabalho ---", ln=True, align='C')
    pdf.ln(10)

    # 1. Número de supervisores únicos por tutor
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(200, 10, txt="1. Número de supervisores únicos por tutor:", ln=True)
    pdf.set_font("Arial", size=10)
    df_temp = supervisores_por_tutor.reset_index(name='Total Supervisores')
    col_width_tutor_supervisor = pdf.w / 3.5
    
    # Cabeçalho da tabela
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(col_width_tutor_supervisor, 10, "Tutor", 1, 0, 'C', 1)
    pdf.cell(col_width_tutor_supervisor, 10, "Total Supervisores", 1, 1, 'C', 1)
    # Linhas da tabela
    for index, row in df_temp.iterrows():
        pdf.cell(col_width_tutor_supervisor, 10, str(row['Tutor']), 1)
        pdf.cell(col_width_tutor_supervisor, 10, str(row['Total Supervisores']), 1, 1)
    pdf.ln(5)

    # 2. Número de médicos supervisionados por supervisor
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(200, 10, txt="2. Número de médicos supervisionados por supervisor:", ln=True)
    pdf.set_font("Arial", size=10)
    df_temp = medicos_por_supervisor.reset_index(name='Total Médicos')
    
    # Cabeçalho da tabela
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(col_width_tutor_supervisor, 10, "Supervisor", 1, 0, 'C', 1)
    pdf.cell(col_width_tutor_supervisor, 10, "Total Médicos", 1, 1, 'C', 1)
    # Linhas da tabela
    for index, row in df_temp.iterrows():
        pdf.cell(col_width_tutor_supervisor, 10, str(row['Supervisor']), 1)
        pdf.cell(col_width_tutor_supervisor, 10, str(row['Total Médicos']), 1, 1)
    pdf.ln(5)

    # 3. Supervisores com menos de 8 ou mais de 10 médicos
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(200, 10, txt="3. Supervisores com menos de 8 ou mais de 10 médicos:", ln=True)
    pdf.set_font("Arial", size=10)
    if not supervisores_menos_de_8_ou_mais_de_10.empty:
        df_temp = supervisores_menos_de_8_ou_mais_de_10.reset_index(name='Total Médicos')
        # Cabeçalho da tabela
        pdf.set_fill_color(230, 230, 230)
        pdf.cell(col_width_tutor_supervisor, 10, "Supervisor", 1, 0, 'C', 1)
        pdf.cell(col_width_tutor_supervisor, 10, "Total Médicos", 1, 1, 'C', 1)
        # Linhas da tabela
        for index, row in df_temp.iterrows():
            pdf.cell(col_width_tutor_supervisor, 10, str(row['Supervisor']), 1)
            pdf.cell(col_width_tutor_supervisor, 10, str(row['Total Médicos']), 1, 1)
    else:
        pdf.cell(200, 10, "Nenhum supervisor nesta categoria.", 0, 1)
    pdf.ln(5)

    # 4. Supervisores que atuam em mais de 2 municípios
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(200, 10, txt="4. Supervisores que atuam em mais de 2 municípios:", ln=True)
    pdf.set_font("Arial", size=10)
    if not supervisores_mais_de_2_municipios.empty:
        df_temp = supervisores_mais_de_2_municipios.reset_index(name='Total Municípios')
        # Cabeçalho da tabela
        pdf.set_fill_color(230, 230, 230)
        pdf.cell(col_width_tutor_supervisor, 10, "Supervisor", 1, 0, 'C', 1)
        pdf.cell(col_width_tutor_supervisor, 10, "Total Municípios", 1, 1, 'C', 1)
        # Linhas da tabela
        for index, row in df_temp.iterrows():
            pdf.cell(col_width_tutor_supervisor, 10, str(row['Supervisor']), 1)
            pdf.cell(col_width_tutor_supervisor, 10, str(row['Total Municípios']), 1, 1)
    else:
        pdf.cell(200, 10, "Nenhum supervisor nesta categoria.", 0, 1)
    pdf.ln(5)

    # 5. Média de médicos por supervisor
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(200, 10, txt="5. Média de médicos por supervisor:", ln=True)
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=f"{media_medicos_por_supervisor:.2f} médicos por supervisor", ln=True)
    pdf.ln(5)

    # 6. Quantidade de Médicos por Mês e Ano (Cumulativo)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(200, 10, txt="6. Quantidade de Médicos por Mês e Ano (Cumulativo):", ln=True)
    pdf.set_font("Arial", size=10)
    if not df_cumulativo_chart.empty:
        df_temp = df_cumulativo_chart.reset_index() # Transforma o índice 'Ano_Mes' em coluna
        headers = df_temp.columns.tolist()
        col_widths = [pdf.w / (len(headers) + 1)] * len(headers) # Largura igual para simplicidade
        
        # Cabeçalho da tabela
        pdf.set_fill_color(230, 230, 230)
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 10, header, 1, 0, 'C', 1)
        pdf.ln()
        # Linhas da tabela
        for index, row in df_temp.iterrows():
            for i, col in enumerate(headers):
                pdf.cell(col_widths[i], 10, str(row[col]), 1)
            pdf.ln()
    else:
        pdf.cell(200, 10, "Não há dados de data válidos para este relatório.", 0, 1)
    pdf.ln(5)

    # 7. Relatório Detalhado por Região e Município
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(200, 10, txt="7. Relatório Detalhado por Região e Município:", ln=True)
    pdf.set_font("Arial", size=10)
    if not relatorio_final_df.empty:
        # Define as larguras das colunas para o relatório detalhado
        col_widths_detail = [40, 40, 60, 60] # Ajustar conforme necessário
        headers_detail = relatorio_final_df.columns.tolist()
        
        # Cabeçalho da tabela
        pdf.set_fill_color(230, 230, 230)
        for i, header in enumerate(headers_detail):
            pdf.cell(col_widths_detail[i], 10, header, 1, 0, 'C', 1)
        pdf.ln()
        # Linhas da tabela
        for index, row in relatorio_final_df.iterrows():
            # Para colunas com texto longo, usar multi_cell para quebrar linha
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            
            for i, col in enumerate(headers_detail):
                if col in ['Tutores', 'Supervisores']:
                    # Salva a posição atual para continuar na próxima célula após o multi_cell
                    current_x = pdf.get_x()
                    current_y = pdf.get_y()
                    pdf.multi_cell(col_widths_detail[i], 5, str(row[col]), 1, 'L', 0)
                    # Retorna à linha inicial e avança para a próxima coluna
                    pdf.set_xy(current_x + col_widths_detail[i], current_y)
                else:
                    pdf.cell(col_widths_detail[i], 10, str(row[col]), 1)
            pdf.ln() # Nova linha após processar todas as colunas da linha atual
    else:
        pdf.cell(200, 10, "Nenhum relatório detalhado disponível.", 0, 1)
    pdf.ln(5)

    # Retorna o PDF como bytes
    return pdf.output(dest='S').encode('latin-1')


# ---- Interface Streamlit ----
st.set_page_config(layout="wide") # Define o layout da página para largura total
st.title("Análise de Plano de Trabalho")

uploaded_file = st.file_uploader("Carregar arquivo XLSX", type=["xlsx"])

if uploaded_file is not None:
    # Ler o arquivo
    df = read_excel_file(uploaded_file)

    if df is not None:
        st.subheader("Pré-visualização dos Dados Carregados")
        # Exibe apenas as primeiras linhas da planilha
        st.dataframe(df.head(), use_container_width=True)

        try:
            # Seleciona as colunas de interesse
            required_columns = ['Município', 'Supervisor', 'Tutor', 'Nome Região', 'Início Atividades']
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                raise KeyError(f"Colunas ausentes no arquivo: {', '.join(missing_cols)}. Verifique se o arquivo possui as colunas esperadas.")

            df = df[required_columns].copy() # Usar .copy() para evitar SettingWithCopyWarning

            # Aplica a correção de capitalização aos nomes dos tutores e supervisores
            df['Supervisor'] = df['Supervisor'].apply(to_title_case)
            df['Tutor'] = df['Tutor'].apply(to_title_case)

            # Limpa dados vazios nas colunas 'Supervisor' e 'Tutor'
            df['Supervisor'] = df['Supervisor'].fillna('')
            df['Tutor'] = df['Tutor'].fillna('')
            # Filtra linhas onde 'Supervisor' ou 'Tutor' estão vazios
            df = df[(df['Supervisor'] != '') & (df['Tutor'] != '')]

            # Converte a coluna de data para o tipo datetime, tratando erros
            df['Início Atividades'] = pd.to_datetime(df['Início Atividades'], errors='coerce', dayfirst=True)
            # Remove linhas onde a conversão de data falhou (NaN)
            df.dropna(subset=['Início Atividades'], inplace=True)

            # Cria uma coluna 'Ano_Mes' para agrupar os dados por mês e ano
            df['Ano_Mes'] = df['Início Atividades'].dt.to_period('M').astype(str)

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
            total_medicos = df.shape[0]
            total_supervisores = df['Supervisor'].nunique()
            media_medicos_por_supervisor = total_medicos / total_supervisores if total_supervisores > 0 else 0

            # 6) Quantidade de médicos por mês e ano (usando 'Início Atividades')
            medicos_por_mes_ano = df.groupby('Ano_Mes').size().sort_index()

            # Calcula a quantidade cumulativa de médicos por mês e ano
            medicos_cumulativos = medicos_por_mes_ano.cumsum()

            # Calcula a porcentagem cumulativa de médicos
            total_medicos_geral = medicos_cumulativos.iloc[-1] if not medicos_cumulativos.empty else 1
            porcentagem_cumulativa = (medicos_cumulativos / total_medicos_geral) * 100

            # Combina os dados para o gráfico cumulativo
            df_cumulativo_chart = pd.DataFrame({
                'Médicos por Mês': medicos_por_mes_ano,
                'Médicos Cumulativos': medicos_cumulativos,
                'Percentual Cumulativo (%)': porcentagem_cumulativa
            }).fillna(0) # Preenche NaNs com 0 caso haja meses sem dados

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

            # ---- Gráfico de Quantidade de Médicos por Mês e Ano (Cumulativo em Barras com Porcentagem) ----
            st.subheader("Quantidade de Médicos por Mês e Ano (Cumulativo)")
            if not df_cumulativo_chart.empty:
                df_chart_reset = df_cumulativo_chart.reset_index()

                df_chart_reset['Percentual Cumulativo Formatado'] = df_chart_reset['Percentual Cumulativo (%)'].apply(lambda x: f"{x:.2f}%")

                bar_chart = alt.Chart(df_chart_reset).mark_bar().encode(
                    x=alt.X('Ano_Mes', sort=None, title='Mês/Ano de Início das Atividades'),
                    y=alt.Y('Médicos Cumulativos', title='Médicos Cumulativos', axis=alt.Axis(titleColor='#5276A7')),
                    tooltip=['Ano_Mes', 'Médicos Cumulativos']
                )

                line_chart = alt.Chart(df_chart_reset).mark_line(color='red').encode(
                    x=alt.X('Ano_Mes', sort=None),
                    y=alt.Y('Percentual Cumulativo (%)', title='Percentual Cumulativo (%)', axis=alt.Axis(titleColor='red')),
                    tooltip=['Ano_Mes', 'Percentual Cumulativo Formatado']
                )

                chart = alt.layer(bar_chart, line_chart).resolve_scale(
                    y='independent'
                ).properties(
                    title='Médicos Cumulativos por Mês de Chegada e Percentual de Crescimento'
                ).interactive()

                st.altair_chart(chart, use_container_width=True)
                st.dataframe(df_cumulativo_chart, use_container_width=True)
            else:
                st.info("Não há dados de data válidos para gerar o gráfico cumulativo de médicos por mês e ano. Verifique a coluna 'Início Atividades'.")

            # ---- Relatório Detalhado por Região e Município (Formato Corrigido) ----
            st.subheader("Relatório Detalhado por Região e Município")

            relatorio_df_list = []
            for (regiao, municipio), grupo in df.groupby(['Nome Região', 'Município']):
                unique_tutores = '; '.join(sorted(grupo['Tutor'].unique()))
                unique_supervisores = '; '.join(sorted(grupo['Supervisor'].unique()))
                relatorio_df_list.append({
                    'Nome Região': regiao,
                    'Município': municipio,
                    'Tutores': unique_tutores,
                    'Supervisores': unique_supervisores
                })
            relatorio_final_df = pd.DataFrame(relatorio_df_list)
            st.dataframe(relatorio_final_df, use_container_width=True)

            # ---- Gerar arquivo PDF para download ----
            pdf_bytes = generate_pdf_report(
                supervisores_por_tutor, medicos_por_supervisor, supervisores_menos_de_8_ou_mais_de_10,
                supervisores_mais_de_2_municipios, media_medicos_por_supervisor, df_cumulativo_chart, relatorio_final_df
            )

            st.download_button(
                label="Baixar relatório completo em PDF",
                data=pdf_bytes,
                file_name="relatorio_supervisao.pdf",
                mime="application/pdf"
            )

        except KeyError as ke:
            st.error(f"Erro: Coluna '{ke}' não encontrada no arquivo. Verifique se o arquivo possui as colunas esperadas: 'Município', 'Supervisor', 'Tutor', 'Nome Região', 'Início Atividades'.")
        except Exception as e:
            st.error(f"Ocorreu um erro no processamento dos dados: {e}")

