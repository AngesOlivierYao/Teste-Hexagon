import sqlalchemy
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from datetime import datetime

def conectar_banco(connection_string):
    """
    Tenta estabelecer uma conexão com o banco de dados SQL Server usando a string de conexão fornecida.

    Args:
        connection_string (str): A string de conexão para o banco de dados SQL Server.

    Returns:
        sqlalchemy.engine.Engine: Um objeto Engine se a conexão for bem-sucedida, None caso contrário.
    """
    try:
        engine = sqlalchemy.create_engine(connection_string)
        # Tenta uma conexão simples para verificar se está tudo OK
        with engine.connect() as connection:
            print("Conexão com o banco de dados estabelecida com sucesso!")
        return engine
    except sqlalchemy.exc.SQLAlchemyError as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
        print("Por favor, verifique a string de conexão e se o servidor está acessível.")
        return None
    except Exception as e:
        print(f"Erro inesperado ao conectar ao banco de dados: {e}")
        return None

def carregar_dados(engine, query):
    """
    Carrega os dados do banco de dados usando a consulta SQL fornecida.

    Args:
        engine (sqlalchemy.engine.Engine): O objeto Engine para usar para a conexão.
        query (str): A consulta SQL para executar.

    Returns:
        pandas.DataFrame: Um DataFrame contendo os dados, ou None em caso de erro.
    """
    try:
        df = pd.read_sql(query, engine)
        print("Dados carregados com sucesso!")
        return df
    except sqlalchemy.exc.SQLAlchemyError as e:
        print(f"Erro ao executar a consulta: {e}")
        print("Por favor, verifique a consulta SQL.")
        return None
    except Exception as e:
        print(f"Erro inesperado ao carregar os dados: {e}")
        return None

def tratar_datas(df):
    """
    Converte a coluna 'OrderDate' para o tipo datetime e extrai Ano e Mês.

    Args:
        df (pandas.DataFrame): O DataFrame para processar.

    Returns:
        pandas.DataFrame: O DataFrame com a coluna 'OrderDate' convertida e colunas 'Ano' e 'Mes' adicionadas,
                          ou None em caso de erro.
    """
    if df is None:
        return None
    try:
        df["OrderDate"] = pd.to_datetime(df["OrderDate"])
        df["Ano"] = df["OrderDate"].dt.year
        df["Mes"] = df["OrderDate"].dt.month
        return df
    except KeyError as e:
        print(f"Erro: Coluna 'OrderDate' não encontrada no DataFrame. Verifique se o nome da coluna está correto: {e}")
        return None
    except Exception as e:
        print(f"Erro inesperado ao tratar as datas: {e}")
        return None

def realizar_analise(df):
    """
    Realiza a análise dos dados de vendas, calculando resumos por região, produto, ano e mês.

    Args:
        df (pandas.DataFrame): O DataFrame contendo os dados de vendas.

    Returns:
        tuple: Uma tupla contendo os DataFrames de resumo (vendas_por_regiao, vendas_por_produto, vendas_por_ano, vendas_por_mes),
               ou None se o DataFrame de entrada for None.
    """
    if df is None:
        return None, None, None, None
    try:
        vendas_por_regiao = df.groupby("Region")["TotalDue"].sum()
        vendas_por_produto = df.groupby("Product")["TotalDue"].sum()
        vendas_por_ano = df.groupby("Ano")["TotalDue"].sum()
        vendas_por_mes = df.groupby(["Ano", "Mes"])["TotalDue"].sum()
        return vendas_por_regiao, vendas_por_produto, vendas_por_ano, vendas_por_mes
    except KeyError as e:
        print(f"Erro: Coluna não encontrada no DataFrame durante a análise: {e}")
        return None, None, None, None
    except Exception as e:
        print(f"Erro inesperado ao realizar a análise: {e}")
        return None, None, None, None

def filtrar_vendas(df, data_inicio, data_fim, produto=None, regiao=None):
    """
    Filtra os dados de vendas com base nos critérios fornecidos.

    Args:
        df (pandas.DataFrame): O DataFrame para filtrar.
        data_inicio (str ou datetime): A data de início do filtro.
        data_fim (str ou datetime): A data de fim do filtro.
        produto (str, opcional): O nome do produto para filtrar.
        regiao (int, opcional): O ID da região para filtrar.

    Returns:
        pandas.DataFrame: Um novo DataFrame contendo os dados filtrados, ou None em caso de erro.
    """
    if df is None:
        return None
    try:
        # Converter data_inicio e data_fim para datetime, se forem strings
        if isinstance(data_inicio, str):
            data_inicio = datetime.strptime(data_inicio, "%Y-%m-%d")
        if isinstance(data_fim, str):
            data_fim = datetime.strptime(data_fim, "%Y-%m-%d")

        df_filtrado = df[(df["OrderDate"] >= data_inicio) & (df["OrderDate"] <= data_fim)].copy() # Cópia para evitar Warning
        # Verifique se as colunas existem antes de filtrar
        if produto and "Product" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Product"] == produto]
        if regiao and "Region" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Region"] == regiao]
        return df_filtrado
    except KeyError as e:
        print(f"Erro: Coluna não encontrada no DataFrame ao filtrar: {e}")
        return None
    except TypeError as e:
        print(f"Erro de tipo ao filtrar as datas: {e}. Certifique-se de usar o formato de data AAAA-MM-DD.")
        return None
    except Exception as e:
        print(f"Erro inesperado ao filtrar os dados: {e}")
        return None

def gerar_graficos(vendas_por_produto, vendas_por_mes):
    """
    Gera os gráficos de barras para vendas por produto e de linhas para vendas ao longo do tempo.

    Args:
        vendas_por_produto (pandas.Series): Série contendo o resumo das vendas por produto.
        vendas_por_mes (pandas.Series): Série contendo o resumo das vendas por mês.
    """
    if vendas_por_produto is not None:
        fig = px.bar(vendas_por_produto, x=vendas_por_produto.index, y="TotalDue", title="Vendas por Produto")
        fig.show(renderer="browser")
    else:
        print("Não foi possível gerar o gráfico de vendas por produto devido a erros anteriores.")

    if vendas_por_mes is not None:
        plt.figure(figsize=(10, 5))
        plt.plot(vendas_por_mes.index, vendas_por_mes.values, marker='o', linestyle='-')
        plt.xlabel("Ano-Mês")
        plt.ylabel("Total Vendas")
        plt.title("Vendas ao Longo do Tempo")
        plt.xticks(rotation=45)
        plt.grid()
        plt.show(block=True)
    else:
        print("Não foi possível gerar o gráfico de vendas ao longo do tempo devido a erros anteriores.")

def main():
    """
    Função principal que coordena o fluxo de execução do script.
    """
    connection_string = "mssql+pyodbc://ANGES/AdventureWorks2022?driver=SQL+Server&Trusted_Connection=yes" # Coloquei aqui para ficar mais claro
    engine = conectar_banco(connection_string)
    if engine is None:
        return  # Encerra a execução se a conexão falhar

    query = """
    SELECT s.OrderDate, s.TotalDue, a.StateProvinceID AS Region, p.Name AS Product
    FROM Sales.SalesOrderHeader s
    JOIN Sales.SalesOrderDetail sd ON s.SalesOrderID = sd.SalesOrderID
    JOIN Production.Product p ON sd.ProductID = p.ProductID
    JOIN Person.Address a ON s.ShipToAddressID = a.AddressID
    """

    df = carregar_dados(engine, query)
    if df is None:
        return  # Encerra se o carregamento dos dados falhar

    df = tratar_datas(df)
    if df is None:
        return  # Encerra se o tratamento das datas falhar

    vendas_por_regiao, vendas_por_produto, vendas_por_ano, vendas_por_mes = realizar_analise(df)
    if vendas_por_regiao is None:
        return # Encerra se a análise falhar

    # Exemplo de uso do filtro (corrigido para usar o DataFrame carregado)
    df_filtrado = filtrar_vendas(df, "2024-01-01", "2024-12-31", produto="Produto X")
    if df_filtrado is not None:
        print("\nDados filtrados:")
        print(df_filtrado.head())

    gerar_graficos(vendas_por_produto, vendas_por_mes)

if __name__ == "__main__":
    main()



