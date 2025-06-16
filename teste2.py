import sqlalchemy
import pyodbc
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt

# Conectar ao banco de dados
conn = pyodbc.connect("DRIVER={SQL Server};SERVER=ANGES;DATABASE=AdventureWorks2022;Trusted_Connection=yes")
# Criação do cursor
cursor = conn.cursor()
query = """
SELECT s.OrderDate, s.TotalDue, a.StateProvinceID, p.Name
FROM Sales.SalesOrderHeader s
JOIN Sales.SalesOrderDetail sd ON s.SalesOrderID = sd.SalesOrderID
JOIN Production.Product p ON sd.ProductID = p.ProductID
JOIN Person.Address a ON s.ShipToAddressID = a.AddressID
"""
# Executa a consulta
cursor.execute(query)

# Exibe os resultados
for row in cursor.fetchall():
    print(row)

# Fecha a conexão
cursor.close()
conn.close()

