import streamlit as st
import pyodbc
import pandas as pd
from decimal import Decimal
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine

    # Informações de conexão
server = 'ANGES'
database = 'AdventureWorks2017'
driver_name = 'ODBC Driver 17 for SQL Server'  # <--- Certifique-se de que este esteja correto
trusted_connection = 'yes'  # Ou 'no' se precisar de usuário e senha

    # Criar a string de conexão para SQLAlchemy
conn_str = f'mssql+pyodbc://@{server}/{database}?driver={driver_name}&Trusted_Connection={trusted_connection}'
engine = create_engine(conn_str)

try:
    # Definição da consulta SQL
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
        ORDER BY soh.OrderDate DESC;

-- Verificação de valores NULL ou vazios
SELECT
    SUM(CASE WHEN soh.OrderDate IS NULL THEN 1 ELSE 0 END) AS Null_OrderDate,
    SUM(CASE WHEN soh.TotalDue IS NULL THEN 1 ELSE 0 END) AS Null_TotalDue,
    SUM(CASE WHEN sp.Name IS NULL OR sp.Name = '' THEN 1 ELSE 0 END) AS Null_Region,
    SUM(CASE WHEN p.Name IS NULL OR p.Name = '' THEN 1 ELSE 0 END) AS Null_Product
FROM AdventureWorks2017.Sales.SalesOrderHeader AS soh
JOIN AdventureWorks2017.Sales.SalesOrderDetail AS sod ON soh.SalesOrderID = sod.SalesOrderID
JOIN AdventureWorks2017.Production.Product AS p ON sod.ProductID = p.ProductID
JOIN AdventureWorks2017.Person.Address AS a ON soh.ShipToAddressID = a.AddressID
JOIN AdventureWorks2017.Person.StateProvince AS sp ON a.StateProvinceID = sp.StateProvinceID;
    """

    # Executa a consulta usando SQLAlchemy e lê para o DataFrame
    df_baseDados = pd.read_sql(query, engine)

    # Converter a coluna TotalDue para Decimal
    df_baseDados["TotalDue"] = df_baseDados["TotalDue"].apply(lambda x: Decimal(str(x))).astype(object)

    # Renomear as colunas
    df_baseDados.rename(columns={
        "OrderDate": "Data",
        "TotalDue": "Total",
        "Region": "Regiao",
        "Product": "Produto"
    }, inplace=True)

    # Criar as colunas 'Ano' e 'Mes'
    df_baseDados['Data'] = pd.to_datetime(df_baseDados['Data'])
    df_baseDados['Ano'] = df_baseDados['Data'].dt.year
    df_baseDados['Mes'] = df_baseDados['Data'].dt.month

    # Exibir o DataFrame
    print(df_baseDados)

    # Título do painel
    st.title("Painel de Análise de Vendas AdventureWorks")
    st.markdown("Explore os dados de vendas por região, produto e ao longo do tempo.")

    # --- Filtros ---
    st.sidebar.header("Filtros")

    # Filtro por período
    data_inicial = st.sidebar.date_input("Data Inicial", df_baseDados['Data'].min())
    data_final = st.sidebar.date_input("Data Final", df_baseDados['Data'].max())
    df_filtrado_data = df_baseDados[(df_baseDados['Data'] >= pd.to_datetime(data_inicial)) & (df_baseDados['Data'] <= pd.to_datetime(data_final))]

    # Filtro por Produto (multiselect)
    produtos_unicos = df_filtrado_data['Produto'].unique()
    produtos_selecionados = st.sidebar.multiselect("Produtos", produtos_unicos)
    if produtos_selecionados:
        df_filtrado_produto = df_filtrado_data[df_filtrado_data['Produto'].isin(produtos_selecionados)]
    else:
        df_filtrado_produto = df_filtrado_data

    # Filtro por Região (multiselect)
    regioes_unicas = df_filtrado_produto['Regiao'].unique()
    regioes_selecionadas = st.sidebar.multiselect("Regiões", regioes_unicas)
    if regioes_selecionadas:
        df_filtrado = df_filtrado_produto[df_filtrado_produto['Regiao'].isin(regioes_selecionadas)]
    else:
        df_filtrado = df_filtrado_produto

    # --- KPI de Vendas Totais ---
    total_vendas_filtrado = df_filtrado['Total'].sum()
    st.subheader(f"Total de Vendas no Período Filtrado: R$ {total_vendas_filtrado:.2f}")
    st.markdown("---")

    # Vendas por Produto
    st.subheader("Vendas por Produto")
    vendas_por_produto = df_filtrado.groupby("Produto")["Total"].sum().sort_values(ascending=False).reset_index()
    fig_produto, ax_produto = plt.subplots(figsize=(10, 60))
    sns.barplot(x='Total', y='Produto', data=vendas_por_produto, ax=ax_produto)
    ax_produto.set_xlabel('Total de Vendas')
    ax_produto.set_ylabel('Produto')
    ax_produto.set_title('Total de Vendas por Produto')
    st.pyplot(fig_produto)


    # Vendas ao Longo do Tempo
    st.subheader("Vendas ao Longo do Tempo")
    vendas_ao_longo_do_tempo = df_filtrado.groupby("Data")["Total"].sum().reset_index()
    fig_tempo, ax_tempo = plt.subplots(figsize=(10, 6))
    sns.lineplot(x='Data', y='Total', data=vendas_ao_longo_do_tempo, marker='o', ax=ax_tempo)
    ax_tempo.set_xlabel('Data')
    ax_tempo.set_ylabel('Total de Vendas')
    ax_tempo.set_title('Vendas ao Longo do Tempo')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    st.pyplot(fig_tempo)

except Exception as e:
    print(f"Ocorreu um erro: {e}")