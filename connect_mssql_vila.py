import pandas as pd

import openpyxl
import pyodbc
#todo check of connectie gemaakt kan worden met SQL Alchemy
# en anders of de data automatisch gegenereerd kan worden uit de MS dbase

cnxn_str = ("Driver={SQL Server Native Client 11.0};"
            "Server=172.27.23.62,1433;"
            "Database=vila;"
            "Trusted_Connection=yes;")
cnxn = pyodbc.connect(cnxn_str)

"""DBMS: Microsoft SQL Server (ver. 14.00.2042)
Case sensitivity: plain=mixed, delimited=mixed
Driver: Microsoft JDBC Driver 11.2 for SQL Server (ver. 11.2.0.0, JDBC4.2)

Ping: 15 ms"""

testQuery = """select SalesorderNo,ShippingDate, Size, Substrate,InvoiceCustomerName
from vila_tb_SalesOrder
where ShippingDate between '2022-12-14 00:00:00' and '2022-12-15 23:59:59'
and InvoiceCustomerName in ('PRINT.COM', 'HELLOPRINT B.V','DRUKWERKDEAL.NL')
-- and LabelHeight =  '40'
and Size = '30 x 30 mm'
and Substrate = '4811 glans wit, permanent';
"""

pd30X30 = pd.read_sql(testQuery, cnxn)
list3030 = [str(row.SalesorderNo) for row in pd30X30.itertuples()]

liststr = tuple(list3030)


newquery = """
        select  SO.SalesOrderNo as ordernummer,
        sol.itemcode as itemnummer,
        sol.Quantity as totaal_aantal,
        sol.LabelsPerRoll as aantal_per_rol, sol.NoOfRolls as aantal_rollen,
        sol.LabelWindingId as rolwikkeling,
        so.CoreId as kernID,
        so.ShippingDate as verstuurdatum,
        so.InvoiceCustomerName as klantnaam,
        so.shapeId as vorm,
        SOL.CuttingDie,
        so.Size as totaalformaat,
        so.LabelWidth, so.LabelHeight,
        so.LabelName,
        so.Substrate,
        so.Dispenser,
        so.ClientOrderNo,
        so.Approved,
        so.Handling,
        so.IsQOrder,
        sol.Description
from vila_tb_SalesOrder as SO
inner join vila_tb_SalesOrderLine as SOL
ON so.SalesOrderNo = sol.SalesOrderNo   
WHERE so.SalesOrderNo in ('202272521', '202272536', '202272567', '202272615', '202272675', '202272685', '202272903', '202272940', '202273100', '202273102');"""
nq = newquery + f"WHERE so.SalesOrderNo in {tuple(list3030)};"



pd30X30_inner = pd.read_sql(newquery, cnxn)
leesfile = pd30X30_inner.to_excel("uit_SQL-TEST-excel/leesfile.xlsx", index=False)

cnxn.close()
