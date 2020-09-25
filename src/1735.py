import xlwings
import re

from tabulate import tabulate

import geometry

NUM_FRAC = re.compile(r'(\d+\s?)?(\d+/\d+)?')


def main():
    wb = xlwings.books.active
    sheet = wb.sheets[0]

    parts = dict()
    mm_counter = 1
    job, shipment = "1200143A-1".split("-")
    for row in sheet.range("A3").expand().value:
        if not any(row):
            break

        for _i, x in enumerate(row):
            if x and type(x) is str:
                row[_i] = x.strip()

        part = parse_row(row)
        part["job"] = job
        part["shipment"] = shipment
        part["mm"] = "07{}".format(str(mm_counter).zfill(3))
        mm_counter += 1

        if part["part"] in parts.keys():
            in_dict = parts[part["part"]]
            test_keys = ("thk", "wid", "len", "grade")

            if all([part[k] == in_dict[k] for k in test_keys]):
                parts[part["part"]]["qty"] += part["qty"]
                continue
            else:
                part["part"] = get_next_part_name(parts, part["part"])

            parts[part["part"]] = part
        else:
            parts[part["part"]] = part

    print(tabulate(parts.values(), headers="keys"))

    process_xml(parts, job, shipment)


def process_xml(parts, job, shipment):
    for p in parts.values():
        p["part"] = f"{job}-{shipment}_{p['part']}"
        p["WO"] = f"PN_{job}-{shipment}"
        geo = geometry.Part(prenest=True, **p)
        geo.generate_xml()

    run = input("Run xml import? ")
    if run.upper().startswith("Y"):
        geometry.run_xml_import()


def parse_row(row):
    mm = None
    qty = 1
    thk = handle_dim(row[2].replace("PL ", ""))
    wid = handle_ft_in(row[4])
    length = handle_ft_in(row[5])
    grade = row[6] + "2"
    part = row[1]

    return dict(
        mm=mm,
        qty=qty,
        thk=thk,
        wid=wid,
        len=length,
        grade=grade,
        part=part,
    )


def handle_ft_in(dim):
    if "-" in dim:
        feet, inches = dim.split("-")
    else:
        feet, inches = 0, dim

    return int(feet) * 12 + handle_dim(inches)


def handle_dim(dim):
    integer, fraction = NUM_FRAC.match(dim).groups()

    value = 0.0
    if integer:
        value += float(integer.strip())
    if fraction:
        num, denom = fraction.split('/')
        value += int(num) / int(denom)

    return value


def get_next_part_name(parts_dict, name):
    suffixes = 'abcdefghijklmnopqrstuvwxyz'
    for suffix in suffixes:
        if name + suffix not in parts_dict.keys():
            return name + suffix


if __name__ == "__main__":
    main()
