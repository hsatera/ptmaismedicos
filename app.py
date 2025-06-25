import pandas as pd
import streamlit as st
from io import StringIO
import itertools # Para ciclar entre cores e símbolos

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

uploaded_file = st.file_uploader("Carregar arquivo XLS ou XLSX", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Ler o arquivo
    df = read_excel_file(uploaded_file)

    if df is not None:
        st.subheader("Pré-visualização dos Dados Carregados")
        # Exibe apenas as primeiras linhas da planilha
        st.dataframe(df.head(), use_container_width=True)

        try:
            # Seleciona as colunas de interesse, incluindo 'Nome Região'
            # Assegura que as colunas 'Município', 'Supervisor', 'Tutor' e 'Nome Região' existem
            required_columns = ['Município', 'Supervisor', 'Tutor', 'Nome Região']
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                raise KeyError(f"Colunas ausentes no arquivo: {', '.join(missing_cols)}. Verifique se o arquivo possui as colunas esperadas.")

            df = df[required_columns].copy() # Usar .copy() para evitar SettingWithCopyWarning

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

            # Obter lista única de tutores para mapeamento
            unique_tutors = sorted(df['Tutor'].unique().tolist())

            # Definir uma lista de cores e símbolos para mapeamento
            # Cores vibrantes e distintas
            colors = [
                "#FF6347", "#6A5ACD", "#3CB371", "#FFD700", "#FF69B4", # Vermelho Tomate, Azul Ardósia, Verde Médio, Ouro, Rosa Quente
                "#1E90FF", "#BA55D3", "#8A2BE2", "#DC143C", "#20B2AA", # Azul Dodger, Orquídea Média, Azul Violeta, Carmesim, Azul Água Clara
                "#DDA0DD", "#8B0000", "#008000", "#4682B4", "#B8860B", # Ameixa, Vermelho Escuro, Verde, Aço Azul, Dourado Escuro
                "#800080", "#FF4500", "#00CED1", "#DA70D6", "#C71585"  # Púrpura, Laranja Avermelhado, Turquesa Escura, Orquídea, Rosa Médio Violeta
            ]
            # Símbolos Unicode simples e claros
            symbols = [
                "●", "■", "▲", "◆", "★", "❖", "⚡", "☀️", "⚙️", "❤️",
                "☘️", "✨", "⏳", "⚛️", "✅", "➕", "➖", "➗", "✖️", "✔️",
                "🟢", "🔵", "🟡", "🟠", "🟣", "🟤", "⚫", "⚪", "🟥", "🟦" # Adicionando mais círculos coloridos
            ]

            # Criar mapeamento apenas para tutores
            tutor_color_symbol_map = {}
            for i, tutor in enumerate(unique_tutors):
                color_index = i % len(colors)
                symbol_index = i % len(symbols)
                tutor_color_symbol_map[tutor] = {
                    'color': colors[color_index],
                    'symbol': symbols[symbol_index]
                }
            
            # --- Cálculos de Análise (permanecem os mesmos) ---
            supervisores_por_tutor = df.groupby('Tutor')['Supervisor'].nunique()
            medicos_por_supervisor = df.groupby('Supervisor').size()
            supervisores_menos_de_8_ou_mais_de_10 = medicos_por_supervisor[
                (medicos_por_supervisor < 8) | (medicos_por_supervisor > 10)
            ]
            supervisores_mais_de_2_municipios = df.groupby('Supervisor')['Município'].nunique()
            supervisores_mais_de_2_municipios = supervisores_mais_de_2_municipios[
                supervisores_mais_de_2_municipios > 2
            ]
            total_medicos = df.shape[0]
            total_supervisores = df['Supervisor'].nunique()
            media_medicos_por_supervisor = total_medicos / total_supervisores if total_supervisores > 0 else 0

            # ---- Exibir Resultados da Análise Geral ----
            st.subheader("Análise Geral")

            # --- Exibir Legenda de Tutores ---
            st.write("### Legenda de Tutores")
            legend_data_tutor = []
            for tutor, props in tutor_color_symbol_map.items():
                legend_data_tutor.append({
                    "Tutor": tutor,
                    "Símbolo": f"<span style='color:{props['color']}; font-size: 20px;'>{props['symbol']}</span>",
                    "Cor Hex": props['color']
                })
            # Usar st.write com unsafe_allow_html para renderizar HTML de DataFrame
            st.write(pd.DataFrame(legend_data_tutor).to_html(escape=False), unsafe_allow_html=True)

            # --- Não exibe legenda de supervisores ---


            st.write("### Número de supervisores únicos por tutor:")
            supervisores_por_tutor_df = supervisores_por_tutor.reset_index(name='Total Supervisores')
            # Adiciona o símbolo colorido à coluna 'Tutor'
            supervisores_por_tutor_df['Tutor'] = supervisores_por_tutor_df['Tutor'].apply(
                lambda x: f"<span style='color:{tutor_color_symbol_map.get(x, {}).get('color', 'black')}; font-size: 20px;'>{tutor_color_symbol_map.get(x, {}).get('symbol', '')}</span> {x}"
            )
            # Exibe o DataFrame com HTML para o símbolo
            st.write(supervisores_por_tutor_df[['Tutor', 'Total Supervisores']].to_html(escape=False), unsafe_allow_html=True)


            st.write("### Número de médicos supervisionados por supervisor:")
            medicos_por_supervisor_df = medicos_por_supervisor.reset_index(name='Total Médicos')
            # Supervisores sem símbolo colorido
            st.dataframe(medicos_por_supervisor_df[['Supervisor', 'Total Médicos']], use_container_width=True)

            st.write("### Supervisores com menos de 8 ou mais de 10 médicos:")
            supervisores_menos_de_8_ou_mais_de_10_df = supervisores_menos_de_8_ou_mais_de_10.reset_index(name='Total Médicos')
            # Supervisores sem símbolo colorido
            st.dataframe(supervisores_menos_de_8_ou_mais_10_df[['Supervisor', 'Total Médicos']], use_container_width=True)


            st.write("### Supervisores que atuam em mais de 2 municípios:")
            supervisores_mais_de_2_municipios_df = supervisores_mais_de_2_municipios.reset_index(name='Total Municípios')
            # Supervisores sem símbolo colorido
            st.dataframe(supervisores_mais_de_2_municipios_df[['Supervisor', 'Total Municípios']], use_container_width=True)

            st.write("### Média de médicos por supervisor:")
            st.write(f"**{media_medicos_por_supervisor:.2f}** médicos por supervisor")

            # ---- Relatório Detalhado por Região e Município (Formato Corrigido) ----
            st.subheader("Relatório Detalhado por Região e Município")

            relatorio_df_list = []
            for (regiao, municipio), grupo in df.groupby(['Nome Região', 'Município']):
                unique_tutores = sorted(grupo['Tutor'].unique())
                tutores_with_symbols = []
                for tutor in unique_tutores:
                    props = tutor_color_symbol_map.get(tutor, {'color': 'black', 'symbol': ''})
                    # Formata cada tutor com seu símbolo colorido
                    tutores_with_symbols.append(f"<span style='color:{props['color']}; font-size: 20px;'>{props['symbol']}</span> {tutor}")

                unique_supervisores = sorted(grupo['Supervisor'].unique())
                supervisores_plain = []
                for supervisor in unique_supervisores:
                    # Supervisores sem símbolo colorido
                    supervisores_plain.append(f"{supervisor}")

                relatorio_df_list.append({
                    'Nome Região': regiao,
                    'Município': municipio,
                    'Tutores': '<br>'.join(tutores_with_symbols), # Usa <br> para novas linhas na célula da tabela HTML
                    'Supervisores': '<br>'.join(supervisores_plain) # Supervisores sem HTML
                })
            relatorio_final_df = pd.DataFrame(relatorio_df_list)
            st.write(relatorio_final_df.to_html(escape=False), unsafe_allow_html=True)


            # ---- Gerar arquivo para download ----
            output = StringIO()
            output.write("--- Relatório de Análise de Plano de Trabalho ---\n\n")

            output.write("--- Legenda de Tutores ---\n")
            for tutor, props in tutor_color_symbol_map.items():
                output.write(f"Tutor: {tutor}, Símbolo: {props['symbol']}, Cor: {props['color']}\n")
            output.write("\n") # Adiciona uma linha em branco para separar

            output.write("1. Número de supervisores únicos por tutor:\n")
            supervisores_por_tutor_for_download = supervisores_por_tutor_df.copy()
            # Remove tags HTML para o download em texto
            supervisores_por_tutor_for_download['Tutor'] = supervisores_por_tutor_for_download['Tutor'].apply(lambda x: x.split('</span> ')[-1])
            supervisores_por_tutor_for_download['Símbolo'] = supervisores_por_tutor_for_download['Tutor'].apply(lambda x: tutor_color_symbol_map.get(x, {}).get('symbol', ''))
            output.write(supervisores_por_tutor_for_download[['Símbolo', 'Tutor', 'Total Supervisores']].to_string(index=False))
            output.write("\n\n2. Número de médicos supervisionados por supervisor:\n")
            medicos_por_supervisor_for_download = medicos_por_supervisor_df.copy()
            # Não é necessário remover tags HTML, pois os supervisores não as terão
            output.write(medicos_por_supervisor_for_download.to_string(index=False))
            output.write("\n\n3. Supervisores com menos de 8 ou mais de 10 médicos:\n")
            supervisores_menos_de_8_ou_mais_de_10_for_download = supervisores_menos_de_8_ou_mais_de_10_df.copy()
            output.write(supervisores_menos_de_8_ou_mais_de_10_for_download.to_string(index=False))
            output.write("\n\n4. Supervisores que atuam em mais de 2 municípios:\n")
            supervisores_mais_de_2_municipios_for_download = supervisores_mais_de_2_municipios_df.copy()
            output.write(supervisores_mais_de_2_municipios_for_download.to_string(index=False))
            output.write("\n\n5. Média de médicos por supervisor:\n")
            output.write(f"{media_medicos_por_supervisor:.2f}\n")
            output.write("\n\n6. Relatório Detalhado por Região e Município:\n")
            
            # Prepara o relatório detalhado para download (sem tags HTML)
            relatorio_df_list_for_download = []
            for (regiao, municipio), grupo in df.groupby(['Nome Região', 'Município']):
                unique_tutores = sorted(grupo['Tutor'].unique())
                tutores_plain = []
                for tutor in unique_tutores:
                    props = tutor_color_symbol_map.get(tutor, {'color': 'black', 'symbol': ''})
                    tutores_plain.append(f"{props['symbol']} {tutor}")

                unique_supervisores = sorted(grupo['Supervisor'].unique())
                supervisores_plain = []
                for supervisor in unique_supervisores:
                    supervisores_plain.append(f"{supervisor}") # Supervisores sem símbolo no download

                relatorio_df_list_for_download.append({
                    'Nome Região': regiao,
                    'Município': municipio,
                    'Tutores': '; '.join(tutores_plain),
                    'Supervisores': '; '.join(supervisores_plain)
                })
            relatorio_final_df_for_download = pd.DataFrame(relatorio_df_list_for_download)
            output.write(relatorio_final_df_for_download.to_string(index=False))


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
