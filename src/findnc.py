import os

nc_paths = [
    r"\\pl2cimnet\Farley\SigmaNest",
    r"\\pl2cimnet\Farley\SigmaNest\checked",
]

to_find = "DRILL  1.563  610  0.008"


def main():
    for path in nc_paths:
        search_files(path)


def search_files(folder):
    for fl in os.scandir(folder):
        if fl.is_dir():
            continue
        with open(fl.path, "r") as ncfile:
            if find_str(ncfile.read()):
                print(fl.name)


def find_str(nc):
    if to_find in nc:
        return True
    return False


if __name__ == "__main__":
    main()
