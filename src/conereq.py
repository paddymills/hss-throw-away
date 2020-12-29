
import os
import csv
import re
import yaml
import shutil

from glob import glob
from collections import namedtuple
from fractions import Fraction


DESC_TEMPLATE = "PL {} x {} x {}{}({})"

def main():
    csv_file = get_file()
    data = parse_csv(csv_file)
    generate_req(data)


def get_file():
    dl = os.scandir(os.path.expanduser(r"~\Downloads"))
    files = list(filter(lambda x: x.name.endswith('.csv'), dl))

    if len(files) > 1:
        print("Available .csv files")
        for i, x in enumerate(files, start=1):
            print(i, "|", x.name)

        selection = input("\nid: ")

        try:
            id = int(selection) - 1

            return files[id]
        except:
            exit()

    # only 1 CSV in downloads
    return files[0]


def parse_csv(csv_file):
    QTY = 2
    MM = 3
    THK = 5
    WID = 6
    LEN = 8
    SPEC = 9
    GRADE = 10
    TEST = 11

    MM_PATTERN = re.compile("\d{7}\w\d{2}-\d{5}")

    items = list()

    print("\nParsing file:", csv_file.name)

    Item = namedtuple('Item', [
        'qty',
        'mm',
        'thk',
        'width',
        'length',
        'spec',
        'grade',
        'test',
    ])

    with open(csv_file.path) as open_csv:
        reader = csv.reader(open_csv)
        for row in reader:
            if len(row) > MM and MM_PATTERN.match(row[MM]):
                items.append(Item(
                    qty=row[QTY],
                    mm=row[MM],
                    thk=float(row[THK]),
                    width=float(row[WID]),
                    length=float(row[LEN]),
                    spec=row[SPEC],
                    grade=row[GRADE],
                    test=row[TEST],
                ))

    return items


def generate_req(data):
    MM_RE = re.compile("(\d{7}\w)(\d{2})-(\d{2})\d{3}")

    # load yaml config
    with open("reqcfg.yaml") as y:
        cfg = yaml.safe_load(y)

    req_filename = get_req_filename()
    req_file = open(req_filename, "w")

    # start data
    date = input("Start Date: ")

    ft2_header_printed = False

    for x in data:

        # material grade
        grade = x.spec + "-" + x.grade + x.test

        if x.width * x.length >= 100000:
            if not ft2_header_printed:
                print("{:20}{:15}{:<10}{:<10}{:<20}".format(
                    "Material Master", "Grade", "Thk", "Width", "Length"))
                print("=" * 75)
                ft2_header_printed = True
            print("{:20}{:15}{:<10}{:<10}{:<20}".format(
                x.mm, grade, x.thk, x.width, x.length))
            continue

        # determine UoM
        if x.spec in ("A709M", "M270M"):
            uom = "M2"
        else:
            uom = "IN2"

        # job, ship and valuation code parsed from MM (and grade for val)
        groups = MM_RE.match(x.mm)
        job = groups[1]
        ship = int(groups[2])
        val_code = calc_valuation_group(int(groups[3]), x.grade, cfg)

        # description
        desc = calc_description(grade, x.thk, x.width, x.length)

        # weight
        weight = int(3.4032 * x.thk * x.width * (x.length / 12))

        # build line
        line = [x.qty, x.mm, uom, grade, desc, x.thk, x.width,
                x.length, job, ship, date, weight, "", val_code, "12", "0", "\n"]

        req_file.write("\t".join([str(l) for l in line]))

    req_file.close()
    shutil.copy(req_filename, r"\\HSSIENG\HSSEDSSERV\SNData\SimTrans\Outbox")


def get_req_filename():
    path = r"\\HSSIENG\HSSEDSSERV\SNData\SimTrans\PurchaseReq Files\Processed"

    REQ_RE = re.compile("sigmanest_req_(\d{4}).ready")
    req_nums = []
    for fn in os.listdir(path):
        if REQ_RE.match(fn):
            req_nums.append(int(REQ_RE.match(fn).group(1)))

    next_req = "sigmanest_req_{:0>4}.ready".format(max(req_nums) + 1)

    print("Req file:", next_req)

    return os.path.join(path, next_req)


def calc_valuation_group(group, grade, yaml_cfg):
    if int(group) == 3:
        type = "Webs"
    elif int(group) == 4:
        type = "Flanges"
    else:
        type = "Detail Plate Mtl"

    if grade.startswith("HPS50"):
        type += " 50 HPS"
    elif grade.startswith("HPS70"):
        type += " 70 HPS"
    else:
        type += " 50/50W"

    for val_code in yaml_cfg['ConfigSettings']['ValuationCodeList']:
        if val_code.endswith(type):
            return val_code


def calc_description(grade, thk, wid, length):
    def numstr(num, show_zero=False):
        if num % 1 == 0:
            return "{:n}".format(num)

        if not show_zero and num // 1 == 0:
            return "{}".format(Fraction(num % 1))

        return "{:n} {}".format(num // 1, Fraction(num % 1))

    _thk = numstr(thk)
    _wid = numstr(wid)
    _len = "{:n}'-{}".format(length // 12, numstr(length % 12, show_zero=True))

    desc = DESC_TEMPLATE.format(_thk, _wid, _len, "{}", grade)
    
    # SAP limit is 40 chars
    spaces = 40 - len(desc) - 2

    # make sure at least 1 space
    # will truncate grade if over 40, but that is Ok
    spaces = max(spaces, 1)

    return desc.format(" " * spaces)


if __name__ == "__main__":
    main()
