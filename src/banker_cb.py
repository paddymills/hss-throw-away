
import xlwings
import re

from tqdm import tqdm

import geometry

SHEET_NAME = re.compile(r"\d+\w-\d+")
# NUM_FRAC = re.compile(r'(\d+)?\s*(\d+/\d+)?"')
NUM_FRAC = re.compile(r'(\d+\s?)?(\d+/\d+)?"?')
GRADE = re.compile(r"(?:ASTM\s*)?(A\d+)-\w*\.?\s*(\d+)\s*\w*")

EXPECTED_HEADER = ["Group", "Item", "Qty", "Description", None, "Length",
                   "Specification", "Testing", "Weight", "Pcmark", "Girder Mark", "Notes"]


def main():
    parts = dict()
    wb = xlwings.books.active
    for sheet in wb.sheets:
        if not SHEET_NAME.match(sheet.name):
            continue

        print("Processing sheet: {}".format(sheet.name))
        result = process_sheet(sheet)

        for part in result.values():
            update_without_replacing(part, parts)

    for part in tqdm(parts.values(), desc="Generating xml files"):
        part["part"] = "{job}-{shipment}_{part}".format(**part)
        part["WO"] = "PN_{job}-{shipment}".format(**part)
        geo = geometry.Part(prenest=True, **part)
        geo.generate_xml()

    run = input("Run xml import? ")
    if run.upper().startswith("Y"):
        geometry.run_xml_import()


def process_sheet(sheet):
    parts = dict()

    # validate header
    header = sheet.range((4, 1), (4, len(EXPECTED_HEADER))).value
    assert header == EXPECTED_HEADER, "Header mismatch on sheet: " + sheet.name

    job, shipment = sheet.name.split("-")
    i = 4
    while 1:
        # increment at beginning so we can use the `continue` statement
        i += 1
        print("\rrow: {}".format(i), end='')

        row = sheet.range((i, 1), (i, len(EXPECTED_HEADER))).value

        # trim spaces from cells
        for _i, x in enumerate(row):
            if x and type(x) is str:
                row[_i] = x.strip()

        if not any(row):
            break

        if process(row):
            part = parse_row(row)
            part["job"] = job
            part["shipment"] = shipment

            update_without_replacing(part, parts)

    print("\r", end='')

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
    try:
        integer, fraction = NUM_FRAC.match(dim).groups()
    except AttributeError as e:
        print("\nDim parse error ({})".format(dim))
        raise e

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


def update_without_replacing(part, parts_dict):
    # duplicate part name (unlikely)
    if part["part"] in parts_dict.keys():
        in_dict = parts_dict[part["part"]]
        test_keys = ("thk", "wid", "len", "grade")

        # if another part with the same name has the same geometry
        if all([part[k] == in_dict[k] for k in test_keys]):
            parts_dict[part["part"]]["qty"] += part["qty"]
            return

        # else, append suffix to name
        else:
            part["part"] = get_next_part_name(parts_dict, part["part"])

    parts_dict[part["part"]] = part


if __name__ == "__main__":
    main()
