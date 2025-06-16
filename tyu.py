import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.ticker as ticker
import locale

# Supondo que 'df_filtrado' já esteja definido e contenha as colunas 'Data' e 'Total'
# Vou criar um df_filtrado de exemplo para que o código rode sem erros
data = {'Data': ['2015', '2016', '2017', '2018', '2019'],
        'Total': [18000, 10000, 15000, 8000, 20000]}
df_filtrado = pd.DataFrame(data)

st.subheader("Faturamento da empresa")
vendas_ao_longo_do_tempo = df_filtrado.groupby("Data")["Total"].sum().reset_index()

# Garante que a coluna 'Data' seja do tipo datetime para melhor formatação
if not pd.api.types.is_datetime64_any_dtype(vendas_ao_longo_do_tempo['Data']):
    try:
        vendas_ao_longo_do_tempo['Data'] = pd.to_datetime(vendas_ao_longo_do_tempo['Data'], format='%Y')  # Especifica o formato do ano
    except Exception as e:
        st.warning(f"Não foi possível converter a coluna 'Data' para datetime. Verifique o formato dos dados. Erro: {e}")

# Define o estilo do gráfico usando Seaborn (opcional, mas pode dar um visual agradável)
sns.set_theme(style="whitegrid")
plt.style.use('seaborn-v0_8-whitegrid')

# Inicializa o locale para formatar a moeda corretamente
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Cria a figura com um tamanho adequado
fig_tempo = plt.figure(figsize=(10, 6))
ax_tempo = fig_tempo.add_subplot(1, 1, 1)

# Plota o gráfico de linhas com cor e marcadores específicos
ax_tempo.plot(vendas_ao_longo_do_tempo['Data'], vendas_ao_longo_do_tempo['Total'], marker='o', linestyle='-', color='#483D8B', linewidth=2, markersize=8) # Cor roxo escuro

# Define os rótulos dos eixos e o título
ax_tempo.set_xlabel('Ano', fontsize=12)
ax_tempo.set_ylabel('Faturamento (R$)', fontsize=12)
ax_tempo.set_title('Faturamento da Empresa ao Longo do Tempo', fontsize=14, fontweight='bold')

# Formata o eixo Y para exibir os valores como moeda (R$)
ax_tempo.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: locale.currency(x, grouping=True)))
ax_tempo.tick_params(axis='y', labelsize=10)

# Formata o eixo X para exibir apenas o ano e aumentar o tamanho da fonte
ax_tempo.xaxis.set_major_locator(ticker.AutoLocator())
ax_tempo.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: pd.to_datetime(x).year))
plt.xticks(rotation=0, fontsize=10)  # Mantém a rotação em 0 e define o tamanho da fonte

# Adiciona uma grade para melhorar a leitura
ax_tempo.grid(True, linestyle='--', alpha=0.7)

# Ajusta o layout para evitar sobreposição de elementos
plt.tight_layout()

# Exibe o gráfico no Streamlit
st.pyplot(fig_tempo)
