import re
import xlwings

from collections import defaultdict

PART_RE = re.compile(r"(\d+\w)-(\d+)_([\d\w]+)")
DESC_RE = re.compile(r"PL ([\d\.]+)[\s-]*(\d*)/?(\d*) x ([\d\.]+)[\s-]*(\d*)/?(\d*) x (\d+)'-+([\d\.]+)[\s-]*(\d*)/?(\d*)[\s\w\d\(\)]*")

def main():
    # SAP export
    wb = xlwings.books['export.XLSX']
    data = wb.sheets[0].range("E2:F2").expand('down').value
    
    mms = list()
    for mm, desc in data:
        match = DESC_RE.match(desc)
        if not match:
            print(desc, "not matched")
            continue

        groups =  match.groups()

        # thickness
        t, tnum, tdenom = groups[:3]
        thk = float(t)
        if tnum:
            thk += float(tnum) / float(tdenom)

        # width
        w, wnum, wdenom = groups[3:6]
        wid = float(w)
        if wnum:
            wid += float(wnum) / float(wdenom)

        # length
        lf, li, linum, lidenom = groups[6:]
        length = float(lf) * 12 + float(li)
        if linum:
            length += float(linum) / float(lidenom)

        mms.append([mm, (thk, wid, length)])

    # RawBOM
    wb = xlwings.books.active

    # prenest sheet list
    ordered = wb.sheets[1].range("B2:H2").expand('down').value
    sizes = dict()
    for row in ordered:
        mm = row[0]
        thk, wid, length = row[4:]
        sizes[mm] = (thk, wid, length)

    # prenest RawBOM
    s = wb.sheets[0]
    s.range("E:E").clear_contents()

    i = 1
    mismatches = list()
    sizes_not_found = list()
    last = s.range("A2").expand('down').last_cell.row
    for i in range(2, last + 1):
        print("\rProcessing line {}/{}".format(i, last), end="")

        task, part_val = s.range((i, 1), (i, 2)).value
        match = PART_RE.match(part_val)

        if not match:
            mismatches.append(i)
            continue

        job, ship, part = match.groups()
        if job[3] != "0":
            job = job[:3] + "0" + job[3:]

        def find_mm():
            size = sizes[task]

            for minus in range(5):
                shipment = int(ship)
                diff = 100
                closest_shipment_mm = None
                for mm, mm_size in mms:
                    if size == mm_size:
                        ship_diff = shipment - int(mm[8:10])
                        if abs(ship_diff) < diff:
                            diff = ship_diff
                            closest_shipment_mm = mm

            if closest_shipment_mm:
                return closest_shipment_mm 
                
            sizes_not_found.append(size)
            return "##### NOT MATCHED #####"

        s.range(i, 5).value = find_mm()

    print("")
    if mismatches:
        for m in mismatches:
            print("Line {} part not matched".format(m))

    if sizes_not_found:
        print("Sizes not found:")
        print("{:^10} | {:^10} | {:^10}".format("Thk", "Wid", "Len"))
        for twl in sorted(list(set(sizes_not_found))):
            print("{:^10} | {:^10} | {:^10}".format(*twl))


if __name__ == "__main__":
    main()
