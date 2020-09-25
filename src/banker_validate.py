
import xlwings
import re

import geometry

SHEET_NAME = re.compile(r"\d+\w-\d+")
# NUM_FRAC = re.compile(r'(\d+)?\s*(\d+/\d+)?"')
NUM_FRAC = re.compile(r'(\d+\s?)?(\d+/\d+)?"')
GRADE = re.compile(r"(?:ASTM\s*)?(A\d+)-\w*\.?\s*(\d+)\s*\w*")

EXPECTED_HEADER = ["Group", "Item", "Qty", "Description", None, "Length",
                   "Specification", "Testing", "Weight", "Pcmark", "Girder Mark"]


def main():
    wb = xlwings.books["Book1"]
    parts = dict()
    for p, m, t in wb.sheets[0].range("A2:C2").expand('down').value:
        parts[p] = t

    wb = xlwings.books.active
    for sheet in wb.sheets:
        if not SHEET_NAME.match(sheet.name):
            continue

        print("Processing sheet: {}".format(sheet.name))
        result = process_sheet(sheet)
        job, shipment = sheet.name.split("-")

        header_printed = False
        template = "|{:<30} | {:^5} | {:^5} |"
        for p in result.values():
            prt = f"{job}-{shipment}_{p['part']}".upper()
            if p["thk"] != parts[prt]:
                if not header_printed:
                    print("\r+------------------------------+-------+-------+")
                    print(template.format("part", "new", "old"))
                    print("+------------------------------+-------+-------+")
                    header_printed = True
                print(template.format(prt, p["thk"], parts[prt]))


def process_sheet(sheet):
    parts = dict()

    # validate header
    header = sheet.range((4, 1), (4, 11)).value
    assert header == EXPECTED_HEADER, "Header mismatch on sheet: " + sheet.name

    job, shipment = sheet.name.split("-")
    i = 4
    while 1:
        # increment at beginning so we can use the `continue` statement
        i += 1
        print("\rrow: {}".format(i), end='')

        row = sheet.range((i, 1), (i, 11)).value
        for _i, x in enumerate(row):
            if x and type(x) is str:
                row[_i] = x.strip()

        if not any(row):
            break

        if process(row):
            part = parse_row(row)
            part["job"] = job
            part["shipment"] = shipment

            if part["part"] in parts.keys():
                in_dict = parts[part["part"]]
                test_keys = ("thk", "wid", "len", "grade")

                if all([part[k] == in_dict[k] for k in test_keys]):
                    parts[part["part"]]["qty"] += part["qty"]
                    continue
                else:
                    part["part"] = get_next_part_name(parts, part["part"])

            parts[part["part"]] = part

    return parts


def parse_row(row):
    mm = str(int(row[1])).zfill(5)
    qty = row[2]
    thk, wid = map(handle_dim, row[4].split("X"))
    length = handle_len(row[5])
    grade = '-'.join(GRADE.match(row[6]).groups()) + "T2"
    part = row[9]

    return dict(
        mm=mm,
        qty=int(qty),
        thk=thk,
        wid=wid,
        len=length,
        grade=grade,
        part=part,
    )


def handle_dim(dim):
    integer, fraction = NUM_FRAC.match(dim).groups()

    value = 0.0
    if integer:
        value += float(integer.strip())
    if fraction:
        num, denom = fraction.split('/')
        value += int(num) / int(denom)

    return value


def handle_len(length):
    feet, inches = length.split("'-")

    return int(feet) * 12 + handle_dim(inches)


def process(row):
    if not row[1]:
        return False
    if row[3] not in ("PL", "FB"):
        return False
    if str(row[1])[0] in ("3", "4"):
        return False

    return True


def get_next_part_name(parts_dict, name):
    suffixes = 'abcdefghijklmnopqrstuvwxyz'
    for suffix in suffixes:
        if name + suffix not in parts_dict.keys():
            return name + suffix


if __name__ == "__main__":
    main()
