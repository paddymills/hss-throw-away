import xlwings
import sndb
import re

from collections import namedtuple


conn = sndb.get_sndb_conn()
cursor = conn.cursor()

cursor.execute("""
SELECT *
FROM
    (
        SELECT
            WONumber, PartName,
            DueDate, DrawingNumber, Material,
            Data1, Data2, Data9
        FROM Part

        UNION
        
        SELECT
            WONumber, PartName,
            DueDate, DrawingNumber, Material,
            Data1, Data2, Data9
        FROM PartArchive
    )
    AS Part
WHERE
    PartName LIKE '1200131__A%'
AND
    WONumber LIKE '1200131%'
AND
    PartName NOT IN
    (
        SELECT Mark as PartName
        FROM JobShipMarks
    )
ORDER BY Data2
""")

to_add = list()
Part = namedtuple('Part', 'wo part date dwg matl job ship mark')

for wo, part, date, dwg, matl, job, ship, mark in cursor.fetchall():
    shipment = re.split("-|_", wo)[1]
    if shipment.startswith("apl"):
        shipment = ship
    if not job or not job.startswith("1200131"):
        job = part.split("_")[0]

    part = part.replace("_", "-")

    to_add.append(Part(wo, part, date, dwg, matl, job, shipment, mark))

conn.close()

xl_app = xlwings.App()
for job, shipment in set([(x.job, x.ship) for x in to_add]):
    wb = xl_app.books.open(r'\\hssieng\DATA\HS\SAP - Material Master_BOM\SigmaNest Work Orders\20{} Work Orders Created\{}-{}_SimTrans_WO.xls'.format(job[1:3], job, shipment))
    sheet = wb.sheets[0]

    i = 2
    while sheet.range((i, 1)).value:
        i += 1

    print("Processing {}-{}".format(job, shipment))

    for part in to_add:
        if part.job != job or part.ship != shipment:
            continue

        sheet.range((i, 1)).value = "SN84"
        sheet.range((i, 2)).value = 2               # district
        sheet.range((i, 5)).value = part.part
        sheet.range((i, 6)).value = 1               # qty
        sheet.range((i, 7)).value = part.matl
        sheet.range((i, 9)).value = part.date
        sheet.range((i, 11)).value = part.dwg
        sheet.range((i, 12)).value = 10             # priority
        sheet.range((i, 28)).value = job
        sheet.range((i, 29)).value = shipment
        sheet.range((i, 36)).value = 7              # machine
        sheet.range((i, 61)).value = part.mark
        sheet.range((i, 66)).value = "HighHeatNum"
        

        i +=1 

    wb.save()
    wb.close()

xl_app.quit()
