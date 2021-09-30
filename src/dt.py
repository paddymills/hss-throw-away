
import sndb
import xlwings

jobs = [
    "1190278A",
    "1190278A",
    "1190181A",
    "1190181A",
    "1190181A",
    "1190181A",
    "1190181A",
    "1190181A",
    "1200045A",
    "1200045A",
    "1190332F",
    "1190332F",
    "1180095L",
    "1180095L",
    "1200055K",
    "1200055K",
    "1190278C",
    "1190278C",
    "1190181A",
    "1190181A",
    "1200055B",
    "1200055B",
    "1200055C",
    "1200055C",
    "1200055B",
    "1200055B",
    "1200055B",
    "1200055L",
    "1200055L",
    "1190252B",
    "1190252B",
    "1200002C",
    "1200002C",
]

with sndb.get_sndb_conn() as db:
    cursor = db.cursor()
    
    results = []
    for job in jobs:
        part = "{}_DT%".format(job)

        cursor.execute("""
            SELECT PartName, ProgramName
            FROM PIPArchive
            WHERE PartName LIKE ?
        """, part)

        for prt, prog in cursor.fetchall():
            results.append((prt, prog))

    
wb = xlwings.Book()
wb.sheets[0].range('A1').value = results
wb.sheets[0].autofit()