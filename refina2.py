import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine
import numpy as np

# Configurações visuais
sns.set(style='whitegrid')

# Conexão com o banco de dados
engine = create_engine(
    "mssql+pyodbc://@ANGES/AdventureWorks2017?driver=ODBC+Driver+18+for+SQL+Server&trusted_connection=yes&TrustServerCertificate=yes"
)

# Consulta SQL
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

# Carregar os dados
df = pd.read_sql(query, engine).drop_duplicates()

# Renomear colunas
df.rename(columns={
    "OrderDate": "Data",
    "TotalDue": "Total",
    "Region": "Regiao",
    "Product": "Produto"
}, inplace=True)

# Conversão e coluna de ano
df['Data'] = pd.to_datetime(df['Data'])
df['Ano'] = df['Data'].dt.year

# Verificar dados nulos
print("Valores nulos por coluna:\n", df.isnull().sum())

# --- OUTLIERS ---
plt.figure(figsize=(6, 4))
sns.boxplot(x=df['Total'])
plt.title('Boxplot - Total de Vendas')
plt.tight_layout()
plt.show()

Q1 = df['Total'].quantile(0.25)
Q3 = df['Total'].quantile(0.75)
IQR = Q3 - Q1
limite_inf = Q1 - 1.5 * IQR
limite_sup = Q3 + 1.5 * IQR

outliers = df[(df['Total'] < limite_inf) | (df['Total'] > limite_sup)]
print("\nOutliers encontrados (IQR):")
print(outliers if not outliers.empty else "Nenhum outlier encontrado.")


# --- INÍCIO DA FILTRAGEM INTERATIVA ---
df_filtrado = df.copy()

try:
    # --- 1. Filtro por Data ---
    print("\n--- Filtro por Data ---")
    data_inicio_str = input("Digite a data de início (AAAA-MM-DD): ").strip()
    data_fim_str = input("Digite a data de fim (AAAA-MM-DD): ").strip()

    data_inicio = pd.to_datetime(data_inicio_str)
    data_fim = pd.to_datetime(data_fim_str)

    df_filtrado = df_filtrado[(df_filtrado['Data'] >= data_inicio) & (df_filtrado['Data'] <= data_fim)].copy()

    if df_filtrado.empty:
        print(f"\nNão há dados para o período entre {data_inicio_str} e {data_fim_str}.")
    else:
        print(f"Dados filtrados para o período: {data_inicio_str} a {data_fim_str}")

    # --- 2. Filtro por Produto ---
    print("\n--- Filtro por Produto ---")
    produtos_disponiveis = sorted(df_filtrado['Produto'].unique())
    print("Produtos disponíveis (digite 'todos' para selecionar todos):")
    for produto in produtos_disponiveis:
        print(f"- {produto}")

    produtos_selecionados_str = input("\nDigite os nomes dos produtos separados por vírgula: ").strip()

    if produtos_selecionados_str.lower() != 'todos':
        produtos_lista = [p.strip() for p in produtos_selecionados_str.split(',')]
        df_filtrado = df_filtrado[df_filtrado['Produto'].isin(produtos_lista)].copy()

        if df_filtrado.empty:
            print("Nenhum dado encontrado para os produtos selecionados.")
            print("Verifique os nomes digitados.")

    # --- 3. Filtro por Região ---
    print("\n--- Filtro por Região ---")
    regioes_disponiveis = sorted(df_filtrado['Regiao'].unique())
    print("Regiões disponíveis (digite 'todas' para selecionar todas):")
    for regiao in regioes_disponiveis:
        print(f"- {regiao}")

    regioes_selecionadas_str = input("\nDigite os nomes das regiões separados por vírgula: ").strip()

    if regioes_selecionadas_str.lower() != 'todas':
        regioes_lista = [r.strip() for r in regioes_selecionadas_str.split(',')]
        df_filtrado = df_filtrado[df_filtrado['Regiao'].isin(regioes_lista)].copy()

        if df_filtrado.empty:
            print("Nenhum dado encontrado para as regiões selecionadas.")
            print("Verifique os nomes digitados.")

    # --- Feedback Final ---
    print("\n--- Análise com Filtros Aplicados ---")
    print("Exibindo as primeiras linhas dos dados filtrados:")
    print(df_filtrado[['Data', 'Produto', 'Regiao', 'Total']].head())

    print("\nVendas Totais por Região:")
    print(df_filtrado.groupby('Regiao')['Total'].sum().sort_values(ascending=False))

    print("\nVendas Totais por Produto:")
    print(df_filtrado.groupby('Produto')['Total'].sum().sort_values(ascending=False))

    print("\nVendas Totais por Ano:")
    print(df_filtrado.groupby('Ano')['Total'].sum().sort_values(ascending=False))

except ValueError:
    print("\nErro: Formato de data inválido. Por favor, use AAAA-MM-DD.")
except Exception as e:
    print(f"\nOcorreu um erro inesperado: {e}")

# Após a aplicação dos filtros...
# Se df_filtrado não estiver vazio, usamos ele. Caso contrário, usamos df original.
if not df_filtrado.empty:
    dados_para_analise = df_filtrado
    print("\nUsando dados filtrados para os gráficos e métricas.")
else:
    dados_para_analise = df
    print("\nFiltros não aplicados ou sem resultados. Usando dados originais.")


# --- CARTÕES VISUAIS (AJUSTADO) ---

# --- MÉTRICAS CHAVE ---
total_vendas = df_filtrado['Total'].sum()
ticket_medio = df_filtrado['Total'].median()

# --- MÉTRICAS CHAVE ---
total_vendas = dados_para_analise['Total'].sum()
ticket_medio = dados_para_analise['Total'].median()

plt.figure(figsize=(10, 4))
for i, (titulo, valor, cor) in enumerate([
    ('Total de Vendas', total_vendas, '#1f77b4'),
    ('Ticket Médio (Mediana)', ticket_medio, '#2ca02c')
]):
    ax = plt.subplot(1, 2, i + 1)
    ax.set_facecolor('#e0f2f7')
    
    # Título do cartão
    ax.text(0.5, 0.7, titulo, ha='center', va='center', fontsize=16, color='dimgray')
    
    # Valor em preto
    ax.text(0.5, 0.3, f'R$ {valor:,.2f}', ha='center', va='center',
            fontsize=22, fontweight='bold', color='black')  # <- aqui está o ajuste
    
    ax.axis('off')
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.patch.set_edgecolor('black')
    ax.patch.set_linewidth(1)

plt.suptitle('Métricas de Vendas', fontsize=18, y=1.02)
plt.tight_layout()
plt.show()

# Top N Produtos por Total de Vendas
top_n = 20
top_produtos = df_filtrado.groupby('Produto')['Total'].sum().sort_values(ascending=False).head(top_n)
# --- TOP PRODUTOS ---
top_produtos = dados_para_analise.groupby('Produto')['Total'].sum().sort_values(ascending=False).head(top_n)
# Gráfico de barras horizontal com estilo limpo
plt.figure(figsize=(12, 8))
ax = sns.barplot(
    x=top_produtos.values,
    y=top_produtos.index,
    color='skyblue'  # Cor definida
)

# Remover grades, bordas e eixos
ax.set_xlabel('')
ax.set_ylabel('')
ax.set_xticks([])
for spine in ax.spines.values():
    spine.set_visible(False)
ax.grid(False)

# Rótulos dentro das barras com texto preto
for bar in ax.patches:
    width = bar.get_width()
    y_pos = bar.get_y() + bar.get_height() / 2
    ax.text(width * 0.98, y_pos, f'R$ {width:,.2f}',
            ha='right', va='center', fontsize=10, color='black')

# Título
plt.title(f'Top {top_n} Produtos por Vendas Totais', fontsize=16)
plt.tight_layout()
plt.show()


# --- GRÁFICO DE LINHA POR ANO (AJUSTADO) ---
vendas_ano = df_filtrado.groupby('Ano')['Total'].sum().reset_index()
# --- GRÁFICO DE LINHA ---
vendas_ano = dados_para_analise.groupby('Ano')['Total'].sum().reset_index()


plt.figure(figsize=(10, 6))
sns.lineplot(data=vendas_ano, x='Ano', y='Total', marker='o', color='skyblue')

# Posicionar os valores um pouco à esquerda dos pontos
for _, row in vendas_ano.iterrows():
    plt.text(row['Ano'] - 0.1, row['Total'], f'R$ {row["Total"]:,.0f}',
             va='bottom', ha='right', fontsize=8)

# Configurações do gráfico
plt.title('Vendas Totais por Ano', fontsize=16)
plt.xlabel('Ano')
plt.ylabel('')
plt.grid(False)
plt.xticks(vendas_ano['Ano'], rotation=45)

# Remover o eixo Y
plt.gca().axes.get_yaxis().set_visible(False)

# Estilo final
sns.despine(left=True, top=True, right=True)
plt.tight_layout()
plt.show()

