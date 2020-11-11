import xlwings
import pyodbc
import sndb
import os
import csv
from string import Template


def main():
    fl = os.path.join(os.path.expanduser(r"~\Documents"), "updates.csv")
    with open(fl, "r") as csvfile:
        conn = sndb.get_sndb_conn()
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
