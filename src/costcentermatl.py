
import sndb
import xlwings

data = list()
with sndb.get_sndb_conn() as db:
    cursor = db.cursor()
    cursor.execute("""
        SELECT
            ArcDateTime, ProgramName, RepeatID,
            SheetName, PartName, WONumber,
            QtyInProcess, NestedArea
        FROM
            PIPArchive
        WHERE 
            PartName NOT LIKE '11%'
        AND
            PartName NOT LIKE '12%'
        AND
            TransType='SN102'
        ORDER BY
            ArcDateTime DESC
    """)

    total = 0
    for x in cursor.fetchall():
        data.append([*x])
        total += x[-1]

    print("Total: {}".format(total))

wb = xlwings.Book()
wb.sheets[0].range("A2").value = data
