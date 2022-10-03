
import sndb
import xlwings

from tqdm import tqdm

s = xlwings.books.active.sheets.active

mm = []
for row in s.range("A5:O5").expand('down').value:
    if row[8] == 'T19':
        if not row[9]:
            continue

        mm.append(row[0])

data = [[]]
for x in tqdm(set(mm), desc="fetching sql"):
    res = sndb.exec_sql_file("sql\\matl_progs.sql", x)
    data[0] = res[0]
    data.extend(res[1:])

wb = xlwings.Book()
wb.sheets[0].range("A1").value = data
wb.sheets[0].autofit()