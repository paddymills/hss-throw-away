
from typing import Match
import pyclip
import re
import sndb
import xlwings

from argparse import ArgumentParser
from collections import defaultdict
from datetime import date, timedelta
from tabulate import tabulate
from tqdm import tqdm

FARO_REGEX = re.compile("(\w+)-\w+-\w+-(\w+)")


def generate_query_list():
    jobs = sorted(set([x.job + "-*" for x in one_year_parts()]))

    put_list_into_clipboard(jobs)
    print("{} distinct jobs to look up loaded into clipboard".format(len(jobs)))


def analyze():
    print("Fetching parts...")
    parts = one_year_parts()
    for row in parts:
        match = FARO_REGEX.match(row.part)
        if match and row.job.endswith(match.group(1)):
            row.part = "{}_{}".format(row.job, match.group(2))

    all_parts = [x.part for x in parts]

    cnf = defaultdict(int)

    wb = xlwings.books.active
    data = wb.sheets[0].range("A1").expand().value
    for row in tqdm(data, desc="Checking SAP"):
        part = row[1].replace("-", "_", 1)
        if part not in all_parts:
            continue

        qty = row[2]
        shipment = int(row[4])

        cnf["{}|{}".format(part, shipment)] += int(qty)

    needs_cnf = list()
    for row in tqdm(parts, "Checking SN"):
        name = "{}|{}".format(row.part, row.shipment)

        if cnf[name] < row.qty:
            needs_cnf.append((*name.split("|"), row.qty - cnf[name]))


    if needs_cnf:
        header = ("Part", "Shipment", "Qty")
        print(tabulate(sorted(needs_cnf), headers=header))
        new_sheet = wb.sheets.add()
        new_sheet.range("A1").value = header
        new_sheet.range("A2").value = needs_cnf
        new_sheet.autofit()


def one_year_parts():
    now = date.today() - timedelta(days=365)
    should_be_cnf = sndb.exec_sql_file("sql\\all_proj.sql", now.isoformat(), no_header=True)

    return should_be_cnf


def put_list_into_clipboard(ls):
    pyclip.copy('\r\n'.join(ls))


if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument("--analyze", action="store_true", help="Run analyze routine")
    
    args = parser.parse_args()

    if args.analyze:
        analyze()

    else:
        generate_query_list()
