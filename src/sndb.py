import os
import pyodbc
import xlwings

from argparse import ArgumentParser
from string import Template
from tabulate import tabulate

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

def exec_sql_file(filename, *args, export_xl=False, no_header=False):
    data = list()

    if not os.path.exists(filename):
        print("Filename does not exist:", filename)
        return

    with open(filename, 'r') as sql_file:
        sql = sql_file.read()

    with get_sndb_conn() as db:
        cursor = db.cursor()

        if args:
            cursor.execute(sql, *args)
        else:
            cursor.execute(sql)

        # get headers
        if not no_header:
            data.append([x[0] for x in cursor.description])

        # get data
        for row in cursor.fetchall():
            # data.append([*row])
            data.append(row)

    if export_xl:
        wb = xlwings.Book()
        wb.sheets[0].range("A1").value = data
        wb.sheets[0].autofit()

    return data


if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument("sql", help="SQL file to execute")
    parser.add_argument("--xl", action="store_true", help="Export to excel spreadsheet")
    parser.add_argument("--save", help="Save path for excel spreadsheet")
    
    args = parser.parse_args()
    data = exec_sql_file(args.sql, export_xl=args.xl)

    if not args.xl:
        print(tabulate(data, headers="firstrow"))
    elif args.save:
        xlwings.books.active.save(args.save)
