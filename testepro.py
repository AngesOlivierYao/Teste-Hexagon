import pyodbc
import pandas as pd
from decimal import Decimal
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine
from sqlalchemy.exc import SQLAlchemyError # Importe esta exceção para tratamento específico

# Informações de conexão
server = 'ANGES'
database = 'AdventureWorks2017'
driver_name = 'ODBC Driver 17 for SQL Server'  # <--- Certifique-se de que este esteja correto
trusted_connection = 'yes'  # Ou 'no' se precisar de usuário e senha

# Criar a string de conexão para SQLAlchemy
# Note que 'mssql+pyodbc' é o dialeto para SQL Server usando pyodbc
conn_str = f'mssql+pyodbc://@{server}/{database}?driver={driver_name}&Trusted_Connection={trusted_connection}'

# Inicializar engine fora do try para garantir que esteja definida,
# ou lidar com a exceção se a criação da engine falhar.
# Neste caso, a criação da engine é mais provável de falhar se os parâmetros da string de conexão estiverem errados.
# O erro mais comum do banco de dados ocorrerá ao tentar ler os dados.
engine = None # Inicializa como None por segurança

try:
    engine = create_engine(conn_str) # A criação da engine pode ser a primeira coisa a falhar
    
    # Definição das consultas SQL
    # A primeira consulta busca os dados
    # A segunda consulta verifica nulos (idealmente, você executaria uma por vez ou as combinaria)
    # Para este exemplo, vamos focar na primeira consulta para carregar o DataFrame.
    query_data = """
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
    """

    # Executa a consulta usando SQLAlchemy e lê para o DataFrame
    df_baseDados = pd.read_sql(query_data, engine) # Use query_data aqui

    # Converter a coluna TotalDue para Decimal (se necessário)
    # Se 'TotalDue' já vem como um tipo numérico que o Pandas reconhece, 'Decimal' pode não ser estritamente necessário
    # e pode até converter para 'object' se não for gerenciado corretamente.
    # Se você precisa de alta precisão financeira e o banco retorna um tipo que o Pandas converte para float,
    # então essa linha faz sentido. Caso contrário, pd.to_numeric pode ser mais simples.
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
    print("DataFrame Original:")
    print(df_baseDados.head()) # Use .head() para não imprimir o DF inteiro
    print("-" * 50)

    # --- Adicionando o método Box Plot e IQR para identificar outliers na coluna 'Total' ---

    # Convertendo a coluna 'Total' para um tipo numérico (float) para cálculos
    df_baseDados['Total_Float'] = df_baseDados['Total'].astype(float)

    # 1. Visualização com Box Plot
    plt.figure(figsize=(8, 6))
    sns.boxplot(y=df_baseDados['Total_Float'])
    plt.title('Box Plot da Coluna Total')
    plt.ylabel('Total das Vendas')
    # plt.show() # Para exibir a figura se você estiver executando localmente
    plt.savefig('box_plot_total_vendas.png')
    print("Box Plot da coluna 'Total' gerado e salvo como 'box_plot_total_vendas.png'.")
    plt.close() # Fecha a figura para liberar memória

    # 2. Método do Intervalo Interquartil (IQR)
    Q1 = df_baseDados['Total_Float'].quantile(0.25)
    Q3 = df_baseDados['Total_Float'].quantile(0.75)
    IQR = Q3 - Q1

    limite_inferior = Q1 - 1.5 * IQR
    limite_superior = Q3 + 1.5 * IQR

    print(f"\nDetalhes do IQR para a coluna 'Total':")
    print(f"Q1 (25º Percentil): {Q1:.2f}")
    print(f"Q3 (75º Percentil): {Q3:.2f}")
    print(f"IQR (Q3 - Q1): {IQR:.2f}")
    print(f"Limite Inferior para Outliers: {limite_inferior:.2f}")
    print(f"Limite Superior para Outliers: {limite_superior:.2f}")

    # Identificando os outliers
    outliers_total = df_baseDados[(df_baseDados['Total_Float'] < limite_inferior) | (df_baseDados['Total_Float'] > limite_superior)]

    print("\nOutliers detectados na coluna 'Total' pelo método IQR:")
    if not outliers_total.empty:
        print(outliers_total[['Data', 'Total', 'Regiao', 'Produto']])
    else:
        print("Nenhum outlier detectado na coluna 'Total' pelo método IQR.")

    # Remover a coluna auxiliar 'Total_Float' se não for mais necessária
    df_baseDados.drop(columns=['Total_Float'], inplace=True)

    print("\nProcessamento de dados e análise de outliers concluídos com sucesso!")

# Bloco except para capturar e tratar exceções
except SQLAlchemyError as e:
    print(f"Erro de SQLAlchemy (banco de dados): {e}")
    print("Verifique suas credenciais, nome do servidor, nome do banco de dados ou a sintaxe da sua consulta SQL.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")
    print("Verifique se o driver ODBC está instalado corretamente e se os caminhos estão corretos.")
finally:
    # Este bloco é executado sempre, independentemente de ocorrer um erro ou não.
    # Pode ser usado para fechar conexões, liberar recursos, etc.
    if 'engine' in locals() and engine is not None:
        # A engine da SQLAlchemy gerencia suas próprias conexões, então fechar explicitamente não é sempre necessário,
        # mas pode ser útil em alguns contextos para garantir liberação de recursos.
        # No caso de read_sql, a conexão é aberta e fechada automaticamente.
        pass 
    print("\nExecução do script finalizada.")