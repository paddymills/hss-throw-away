import re
import math

G1_line = re.compile("X([-.\d]+)Y([-.\d]+)")
X_OFFSET_LIMIT = 6      # will not compare points more than this in X-direction
WEB_DEPTH_LIMIT = 24    # will not compare points less than this in Y-direction
END_MASK_LENGTH = 6     # will not compare points within this many inches from the end


def main():
    prog = 6
    # nc_file = r"\\hssieng\SNDataDev\NC\PlateProcessors\Post\{}.n".format(prog)
    nc_file = r"C:\Users\PMiller1\Downloads\46011.n"

    plasma_selected = False
    plasma_on = False   # M77/M79
    line_mode = False   # G1

    points = list()
    with open(nc_file, 'r') as nc:
        for i, line in enumerate(nc.readlines(), start=1):
            line = line.strip()

            if line == "T01=1":
                plasma_selected = True
            elif line == "T00":
                plasma_selected = False

            if not plasma_selected:
                continue

            if line == "M77":
                plasma_on = True
            elif line == "M79":
                plasma_on = False

            if line.startswith("G1"):
                line_mode = True
                line = line[2:]

            match = G1_line.match(line)
            if line_mode and match:
                fx, fy = match.groups()
                points.append((float(fy), float(fx)))
            else:
                line_mode = False

    xmin = min([x[0] for x in points]) + END_MASK_LENGTH
    xmax = max([x[0] for x in points]) - END_MASK_LENGTH
    deltas = list()
    for x, y in points:
        if not (xmin < x < xmax):
            continue

        try:
            nx, ny = find_closest_point(x, y, points)
            dist = math.sqrt((nx - x) ** 2 + (ny - y) ** 2)

            deltas.append(dist)
            # print("({}, {}), ({}, {}): {}".format(x, y, nx, ny, dist))
        except TypeError:
            # print("No matching point ({}, {})".format(x, y))
            pass

    print("Point pairs:", len(deltas))
    print("Min:", min(deltas))
    print("Max:", max(deltas))
    print("Avg:", sum(deltas) / len(deltas))


def find_closest_point(x, y, pts):
    xdiff = 500
    res = None
    for px, py in pts:
        if abs(px - x) > X_OFFSET_LIMIT:
            continue
        if abs(py - y) < WEB_DEPTH_LIMIT:
            continue
        if px == x and py == y:
            continue

        if abs(px - x) < xdiff:
            xdiff = abs(px - x)
            res = (px, py)

    return res


if __name__ == "__main__":
    main()
