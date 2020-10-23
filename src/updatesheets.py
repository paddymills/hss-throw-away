import xlwings
import pyodbc
import os
import csv
from string import Template

SNDB_PRD = "HIIWINBL18"
SNDB_DEV = "HIIWINBL5"

CONN_STR_TEMPLATE = Template(
    "DRIVER={$driver};SERVER=$server;UID=$user;PWD=$pwd;DATABASE=$db;")

cs_kwargs = dict(
    driver="SQL Server",
    server=SNDB_PRD,
    db="SNDBase91",
    user=os.getenv('SNDB_USER'),
    pwd=os.getenv('SNDB_PWD'),
)


def get_sndb_conn(dev=False, **kwargs):
    cs_kwargs.update(kwargs)

    if dev:
        cs_kwargs['server'] = SNDB_DEV
        cs_kwargs['db'] = "SNDBaseDev"

    connection_string = CONN_STR_TEMPLATE.substitute(**cs_kwargs)

    return pyodbc.connect(connection_string)


def main():
    fl = os.path.join(os.path.expanduser(r"~\Documents"), "updates.csv")
    with open(fl, "r") as csvfile:
        conn = get_sndb_conn()
        cur = conn.cursor()

        for sheet, loc in csv.reader(csvfile):
            cur.execute("""
                UPDATE Stock
                SET Location=?
                WHERE SheetName=?
            """, loc, sheet)

    conn.commit()
    conn.close()


if __name__ == "__main__":
    main()
