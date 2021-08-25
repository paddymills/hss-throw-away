
import os

ugs = r"\\hssieng\hssedsserv\Program Files\UGS"
folders = ("MACH", "resource")
nx1953 = os.path.join(ugs, "NX1953")
nx12 = os.path.join(ugs, "NX 12.0")

known_decode_error_exts = [
    "jt",
    "pdf",
    "xml",
    "exe",
    "db",
    # "py",
    # "htm",
    "dll",
    "jpg",
    "pce",
    "tif",
    "png",
    "CCF",
    "xlsx",
    "bmp",
    "zip",
    "gif",
    "docx",
    "xls",
    "prt",
    "MCF",
    # "tcl",
    "cyc",
]
decode_error_exts = []

def diff_folder(folder):
    if len(folder.split("\\")) > 5:
        return

    if not os.path.exists(os.path.join(nx12, folder)):
        print("Folder does not exist in NX12:", folder)
        return

    print("searching folder:", folder)

    for x in os.listdir(os.path.join(nx1953, folder)):
        if os.path.isdir(os.path.join(nx1953, folder, x)):
            diff_folder(os.path.join(folder, x))
        
        elif x.split(".")[-1] in known_decode_error_exts:
            continue

        elif os.path.exists(os.path.join(nx12, folder, x)):
            new = open(os.path.join(nx1953, folder, x), 'r')
            old = open(os.path.join(nx12, folder, x), 'r')

            try:
                if new.read() != old.read():
                    print("\tFile mismatch:", x)
            except UnicodeDecodeError:
                print("Error decoding: ", os.path.join(nx12, folder, x))
                decode_error_exts.append(x.split(".")[-1])

            new.close()
            old.close()
        
        else:
            print("File does not exist in NX12:", os.path.join(folder, x))

# diff_folder("MACH\\resource")
diff_folder("MACH")

for x in set(decode_error_exts):
    print("Decode error ext:", x)
