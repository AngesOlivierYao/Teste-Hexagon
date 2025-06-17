import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from sqlalchemy import create_engine

# Configurações iniciais
st.set_page_config(page_title='Painel de Vendas', layout='wide')

# Estilo customizado da sidebar
st.markdown("""
    <style>
    [data-testid="stSidebar"] {
        background-color: skyblue;
    }
    </style>
""", unsafe_allow_html=True)
#Titulo
st.title('📊 Painel de Vendas')
# Espaçamento após o título
st.markdown("<div style='margin-bottom: 30px;'></div>", unsafe_allow_html=True)


# Conexão com banco de dados
@st.cache_data

def carregar_dados():
    engine = create_engine(
        "mssql+pyodbc://@ANGES/AdventureWorks2017?driver=ODBC+Driver+18+for+SQL+Server&trusted_connection=yes&TrustServerCertificate=yes"
    )
    query = """
    SELECT
        soh.OrderDate,
        soh.TotalDue,
        sp.Name AS Region,
        p.Name AS Product
    FROM AdventureWorks2017.Sales.SalesOrderHeader AS soh
    JOIN AdventureWorks2017.Sales.SalesOrderDetail AS sod ON soh.SalesOrderID = sod.SalesOrderID
    JOIN AdventureWorks2017.Production.Product AS p ON sod.ProductID = p.ProductID
    JOIN AdventureWorks2017.Person.Address AS a ON soh.ShipToAddressID = a.AddressID
    JOIN AdventureWorks2017.Person.StateProvince AS sp ON a.StateProvinceID = sp.StateProvinceID
    """
    df = pd.read_sql(query, engine).drop_duplicates()
    df.rename(columns={"OrderDate": "Data", "TotalDue": "Total", "Region": "Regiao", "Product": "Produto"}, inplace=True)
    df['Data'] = pd.to_datetime(df['Data'])
    df['Ano'] = df['Data'].dt.year
    return df

df = carregar_dados()

# --- SIDEBAR COM FILTROS ---
st.sidebar.header("Filtros")

data_inicio = st.sidebar.date_input("Data de Início", df['Data'].min())
data_fim = st.sidebar.date_input("Data de Fim", df['Data'].max())

produtos = ['Todos'] + sorted(df['Produto'].unique().tolist())
produtos_selecionados = st.sidebar.multiselect("Produto(s)", produtos, default=['Todos'])

regioes = ['Todas'] + sorted(df['Regiao'].unique().tolist())
regioes_selecionadas = st.sidebar.multiselect("Região(ões)", regioes, default=['Todas'])

# --- APLICAR FILTROS ---
df_filtrado = df[(df['Data'] >= pd.to_datetime(data_inicio)) & (df['Data'] <= pd.to_datetime(data_fim))]

if 'Todos' not in produtos_selecionados:
    df_filtrado = df_filtrado[df_filtrado['Produto'].isin(produtos_selecionados)]

if 'Todas' not in regioes_selecionadas:
    df_filtrado = df_filtrado[df_filtrado['Regiao'].isin(regioes_selecionadas)]

# Usar dados filtrados ou originais
if not df_filtrado.empty:
    dados = df_filtrado
else:
    dados = df

# --- CARTÕES DE MÉTRICAS ---
st.subheader("")
col1, col2 = st.columns(2)
with col1:
    st.markdown("<h4>Total de Vendas</h4>", unsafe_allow_html=True)
    st.write(f"**R$ {dados['Total'].sum():,.2f}**")

with col2:
    st.markdown("<h4>Ticket Médio (Mediana)</h4>", unsafe_allow_html=True)
    st.write(f"**R$ {dados['Total'].median():,.2f}**")

# --- ESPAÇAMENTO VISUAL ---
st.markdown("---")
st.markdown("### Top 10 Produtos por Total de Vendas")
st.markdown("&nbsp;")

# --- GRÁFICO DE TOP PRODUTOS ---
top_produtos = dados.groupby('Produto')['Total'].sum().sort_values(ascending=False).head(10)
fig1, ax1 = plt.subplots(figsize=(10, 5))
sns.barplot(x=top_produtos.values, y=top_produtos.index, ax=ax1, color='skyblue')

# Estética do gráfico de barras
ax1.set_xlabel('')
ax1.set_xticks([])
ax1.set_ylabel('')
for spine in ['top', 'right', 'left', 'bottom']:
    ax1.spines[spine].set_visible(False)
ax1.tick_params(axis='y', left=False)
ax1.grid(False)

# Valores dentro das barras
for bar, value in zip(ax1.patches, top_produtos.values):
    ax1.text(
        value * 0.98,
        bar.get_y() + bar.get_height() / 2,
        f'R$ {value:,.2f}',
        va='center', ha='right', fontsize=9, color='black'
    )

st.pyplot(fig1)

# --- ESPAÇAMENTO ENTRE GRÁFICOS ---
st.markdown("&nbsp;")
st.markdown("---")
st.markdown("### Vendas Totais por Ano")
st.markdown("&nbsp;")

# --- GRÁFICO DE VENDAS POR ANO ---

vendas_ano = dados.groupby('Ano')['Total'].sum().reset_index()

fig2, ax2 = plt.subplots(figsize=(10, 6))
sns.lineplot(data=vendas_ano, x='Ano', y='Total', marker='o', color='skyblue', ax=ax2)

# Coloca valores um pouco à esquerda e acima dos pontos
for _, row in vendas_ano.iterrows():
    ax2.text(row['Ano'] - 0.1, row['Total'] * 1.01, f'R$ {row["Total"]:,.0f}', ha='right', va='bottom', fontsize=8)

# Formatação do gráfico
ax2.set_xlabel('Ano')
ax2.set_ylabel('')
ax2.grid(False)
ax2.set_xticks(vendas_ano['Ano']) 
ax2.get_yaxis().set_visible(False)
sns.despine(left=True, top=True, right=True)
st.pyplot(fig2)
