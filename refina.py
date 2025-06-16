import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine
import numpy as np # Importar numpy para tick_params

# String de conexão com SQLAlchemy
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
df = pd.read_sql(query, engine)

# Remover duplicatas
df = df.drop_duplicates()
print(df)

# Verificar tipos de dados
print("Tipos de dados antes da conversão:")
print(df.dtypes)

# Verificar valores nulos
print("\nValores nulos por coluna:")
print(df.isnull().sum())

# Boxplots para detectar outliers
plt.subplot(1, 2, 2)
sns.boxplot(x=df['TotalDue'])
plt.title('Boxplot - TotalDue')
plt.tight_layout()
plt.show()

# Cálculo do IQR
Q1 = df['TotalDue'].quantile(0.25)
Q3 = df['TotalDue'].quantile(0.75)
IQR = Q3 - Q1
limite_inferior = Q1 - 1.5 * IQR
limite_superior = Q3 + 1.5 * IQR
print(f"\nIQR TotalDue: {IQR}")

# Identificando os outliers
outliers_total = df[(df['TotalDue'] < limite_inferior) | (df['TotalDue'] > limite_superior)]

print("\nOutliers detectados na coluna 'Total' pelo método IQR:")
if not outliers_total.empty:
    print(outliers_total)
else:
    print("Nenhum outlier detectado na coluna 'Total' pelo método IQR.")

# Frequência de produtos
print("\nFrequência dos produtos:")
print(df['Product'].value_counts())

# Renomear as colunas
df.rename(columns={
        "OrderDate": "Data",
        "TotalDue": "Total",
        "Region": "Regiao",
        "Product": "Produto"
    }, inplace=True)

# Criar as colunas 'Ano'
df['Data'] = pd.to_datetime(df['Data'])
df['Ano'] = df['Data'].dt.year

# --- NOVOS CÁLCULOS: Ticket Médio (Mediana) e Valor Total de Vendas ---
print("\n--- Métricas de Vendas Chave ---")

# Calcular o valor total das vendas
valor_total_vendas = df['Total'].sum()
print(f"Valor Total de Vendas: R$ {valor_total_vendas:,.2f}")

# Calcular o ticket médio usando a mediana
# Se 'Total' é o total do pedido, a mediana de 'Total' é o ticket médio mediano.
ticket_medio_mediana = df['Total'].median()
print(f"Ticket Médio (Mediana): R$ {ticket_medio_mediana:,.2f}")
# --- FIM DOS NOVOS CÁLCULOS ---

# 1. Vendas totais por região e por produto
total_vendas_por_regiao = df.groupby('Regiao')['Total'].sum().sort_values(ascending=False)
total_vendas_por_produto = df.groupby('Produto')['Total'].sum().sort_values(ascending=True)

print(total_vendas_por_regiao)
print(total_vendas_por_produto)

# 2. Vendas totais por período de tempo.
total_vendas_por_Data = df.groupby('Data')['Total'].sum().sort_values(ascending=False)

print(total_vendas_por_Data)

# --- Quantidade Total de Todos os Itens (Produtos) ---
print("\n--- Quantidade Total de TODOS os Itens ---")
quantidade_total_itens = df['Total'].count()
print(f"Quantidade total de itens vendidos: {quantidade_total_itens}")

# --- Início da Filtragem Interativa ---
# Começamos com o DataFrame completo e aplicaremos os filtros sequencialmente
df_filtrado = df.copy()

try:
    # --- 1. Filtro por Data ---
    print("\n--- Filtro por Data ---")
    data_inicio_str = input("Digite a data de início (AAAA-MM-DD2): ")
    data_fim_str = input("Digite a data de fim (AAAA-MM-DD): ")

    data_inicio = pd.to_datetime(data_inicio_str)
    data_fim = pd.to_datetime(data_fim_str)
    
    df_filtrado = df_filtrado[(df_filtrado['Data'] >= data_inicio) & (df_filtrado['Data'] <= data_fim)].copy()
    
    if df_filtrado.empty:
        print(f"\nNão há dados para o período entre {data_inicio_str} e {data_fim_str}.")
    else:
        print(f"Dados filtrados para o período: {data_inicio_str} a {data_fim_str}")

    # --- 2. Filtro por Produto ---
    print("\n--- Filtro por Produto ---")
    produtos_disponiveis = df_filtrado['Produto'].unique().tolist()
    print("Produtos disponíveis para filtragem (digite 'todos' para selecionar todos):")
    for produto in sorted(produtos_disponiveis):
        print(f"- {produto}")

    produtos_selecionados_str = input("\nDigite os nomes dos produtos que deseja analisar, separados por vírgula (ex: 'Mountain Bike, Road Bike'): ").strip()

    if produtos_selecionados_str.lower() != 'todos':
        produtos_lista = [p.strip() for p in produtos_selecionados_str.split(',')]
        df_filtrado = df_filtrado[df_filtrado['Produto'].isin(produtos_lista)].copy()

        if df_filtrado.empty:
            print(f"Nenhum dado encontrado para os produtos selecionados neste período.")
            print("Verifique se os nomes dos produtos foram digitados corretamente.")
        
    # --- 3. Filtro por Região ---
    print("\n--- Filtro por Região ---")
    regioes_disponiveis = df_filtrado['Regiao'].unique().tolist()
    print("Regiões disponíveis para filtragem (digite 'todas' para selecionar todas):")
    for regiao in sorted(regioes_disponiveis):
        print(f"- {regiao}")

    regioes_selecionadas_str = input("\nDigite os nomes das regiões que deseja analisar, separadas por vírgula (ex: 'California, Washington'): ").strip()

    if regioes_selecionadas_str.lower() != 'todas':
        regioes_lista = [r.strip() for r in regioes_selecionadas_str.split(',')]
        df_filtrado = df_filtrado[df_filtrado['Regiao'].isin(regioes_lista)].copy()

        if df_filtrado.empty:
            print(f"Nenhum dado encontrado para as regiões selecionadas nesta combinação de data/produto.")
            print("Verifique se os nomes das regiões foram digitados corretamente.")

    # --- Análise Final com Dados Filtrados ---
    print("\n--- Análise de Vendas com Filtros Aplicados ---")
    print("Cabeçalho do DataFrame final filtrado (data, produto e região):")
    print(df_filtrado.head())

    # 1. Vendas totais por região e por produto
    print("\n--- Vendas Totais por Região ---")
    total_vendas_por_regiao = df_filtrado.groupby('Regiao')['Total'].sum().sort_values(ascending=False)
    print(total_vendas_por_regiao)

    print("\n--- Vendas Totais por Produto ---")
    total_vendas_por_produto = df_filtrado.groupby('Produto')['Total'].sum().sort_values(ascending=True)
    print(total_vendas_por_produto)

    # 2. Vendas totais por período de tempo (ano)
    print("\n--- Vendas Totais por Ano ---")
    total_vendas_por_ano = df_filtrado.groupby('Ano')['Total'].sum().sort_values(ascending=False)
    print(total_vendas_por_ano) 

except ValueError:
    print("\nErro: Formato de data inválido. Por favor, use AAAA-MM-DD.")
except Exception as e:
    print(f"\nOcorreu um erro inesperado: {e}")

# --- NOVOS GRÁFICOS: CARTÕES PARA TICKET MÉDIO E VALOR TOTAL ---
plt.figure(figsize=(10, 4))  # Uma figura para os dois cartões

# Cartão para Valor Total de Vendas
ax1 = plt.subplot(1, 2, 1)  # 1 linha, 2 colunas, 1º subplot
ax1.set_facecolor('#e0f2f7')  # Fundo azul claro
ax1.text(0.5, 0.7, 'Total de Vendas', ha='center', va='center', fontsize=18, color='dimgray')
ax1.text(0.5, 0.3, f'R$ {valor_total_vendas:,.2f}', ha='center', va='center',
         fontsize=24, fontweight='bold', color='#1f77b4')  # Azul mais escuro
ax1.axis('off')  # Remove os eixos
ax1.set_xlim(0, 1)
ax1.set_ylim(0, 1)
ax1.patch.set_edgecolor('black')  # Borda do cartão
ax1.patch.set_linewidth(1)  # Espessura da borda

# Cartão para Ticket Médio (Mediana)
ax2 = plt.subplot(1, 2, 2)  # 1 linha, 2 colunas, 2º subplot
ax2.set_facecolor('#e0f2f7')  # Fundo azul claro
ax2.text(0.5, 0.7, 'Ticket Médio (Mediana)', ha='center', va='center', fontsize=18, color='dimgray')
ax2.text(0.5, 0.3, f'R$ {ticket_medio_mediana:,.2f}', ha='center', va='center',
         fontsize=24, fontweight='bold', color='#2ca02c')  # Verde
ax2.axis('off')  # Remove os eixos
ax2.set_xlim(0, 1)
ax2.set_ylim(0, 1)
ax2.patch.set_edgecolor('black')  # Borda do cartão
ax2.patch.set_linewidth(1)  # Espessura da borda

plt.suptitle('Principais Métricas de Vendas', fontsize=20, y=1.05)  # Título geral para os dois cartões
plt.tight_layout(rect=[0, 0, 1, 0.98])  # Ajusta o layout para evitar sobreposição
plt.show()


# 2. Selecionar o Top N
N = 20  # ou outro valor desejado
top_n_produtos = total_vendas_por_produto.head(N)

# Verificar a ordem
print("\n--- Dados do Top N Produtos ANTES de gerar o gráfico ---")
print(top_n_produtos)
print("-" * 50)

# Criar o gráfico de barras horizontais
plt.figure(figsize=(12, 8))
ax = sns.barplot(x=top_n_produtos.values, y=top_n_produtos.index, color='skyblue')

plt.title(f'Top {N} Produtos por Vendas Totais (Ordem Crescente com Maior no Topo)', fontsize=16)

# Configurações dos Eixos
plt.xlabel('')
ax.set_xticks([])
plt.ylabel('')
ax.tick_params(axis='y', length=0)

# Remover bordas
for spine in ['left', 'right', 'top', 'bottom']:
    ax.spines[spine].set_visible(False)

# Adicionar os valores nas barras
for p in ax.patches:
    width = p.get_width()
    plt.text(width * 0.95,
             p.get_y() + p.get_height() / 2,
             f'{width:,.2f}',
             ha='right',
             va='center',
             fontsize=10,
             color='black')

# Inverter o eixo Y para mostrar o maior valor no topo
plt.gca().invert_yaxis()

plt.tight_layout()
plt.show()

# --- GRÁFICO DE LINHAS DE VENDAS TOTAIS POR ANO (AJUSTADO: Sem Moldura, Apenas Anos Absolutos no Eixo X) ---
print("\n--- Gerando Gráfico: Vendas Totais por Ano ---")

vendas_por_ano_plot = (
    df
    .groupby('Ano')['Total']
    .sum()
    .reset_index()
    .sort_values(by='Ano')
)

if vendas_por_ano_plot.empty:
    print("Não há dados de vendas por ano para gerar o gráfico de linhas.")
else:
    plt.figure(figsize=(10, 6))
    ax_line = sns.lineplot(
        data=vendas_por_ano_plot,
        x='Ano',
        y='Total',
        marker='o',
        color='skyblue'
    )

    plt.title('Vendas Totais por Ano', fontsize=16)
    plt.xlabel('Ano', fontsize=12)

    # --- MUDANÇAS AQUI: Remover Eixo Y, Grades e Ajustar Eixo X ---
    plt.ylabel('')
    ax_line.set_yticks([])
    ax_line.spines['left'].set_visible(False)
    ax_line.spines['right'].set_visible(False)
    ax_line.spines['top'].set_visible(False)
    plt.grid(False)

    anos_unicos = vendas_por_ano_plot['Ano'].unique()
    ax_line.set_xticks(anos_unicos)
    ax_line.tick_params(axis='x', length=5)
    # --- FIM DAS MUDANÇAS ---

    plt.xticks(rotation=45)

    # --- Adicionar valores aos pontos ---
    for index, row in vendas_por_ano_plot.iterrows():
        valor_formatado = f'{row["Total"]:,.2f}'
        plt.text(
            row["Ano"] - 0.15,
            row["Total"],
            valor_formatado,
            ha='right',
            va='center',
            fontsize=9,
            color='black'
        )

    plt.tight_layout()
    plt.show()
