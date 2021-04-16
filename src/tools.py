
import sndb
import xlwings

data = []
for line in xlwings.books.active.sheets.active.range("A2:N120").value:
    data.append([line[0], line[1], line[2]])


db = sndb.get_sndb_conn()
cursor = db.cursor()
for speed, feed, toolid in data:
    cursor.execute("""
        UPDATE Tool
        SET ToolParam5=?, ToolParam6=?
        WHERE ToolID=?
    """, speed, feed, toolid)

db.close()