
import xlwings
import sndb

with sndb.get_sndb_conn() as db:
    cursor = db.cursor()
    cursor.execute("""
        SELECT
            PIP.WONumber, PIP.PartName,
            Stock.SheetName, Stock.PrimeCode, Stock.Mill,
            PIP.RectArea
        FROM PIPArchive AS PIP
        INNER JOIN StockHistory AS Stock
            ON PIP.ProgramName=Stock.ProgramName
            AND PIP.SheetName=Stock.SheetName
        WHERE PIP.PartName LIKE '1200131%'
        AND PIP.TransType = 'SN102'
    """)

    results = [['WorkOrder', 'Part', 'Sheet', 'SAP MM', 'WBS Element', 'Rectangular Area']]

    for x in cursor.fetchall():
        results.append(list(x))


wb = xlwings.Book()
wb.sheets[0].range('A1').value = results
wb.sheets[0].autofit()
