
import os
import shutil
from tqdm import tqdm

PATH = r"\\hssieng\Jobs"
dest = r"C:\temp\dxfwebs"

def main():
    to_copy = []
    for i in range(18, 23):
        year_folder = os.path.join(PATH, "20{}".format(i))
        for folder in os.listdir(year_folder):
            dxf_folder = os.path.join(year_folder, folder, "Fab", "Webs", "DXF")
            if os.path.exists(dxf_folder):
                for f in os.listdir(dxf_folder):
                    if f.endswith(".dxf"):
                        to_copy.append(os.path.join(dxf_folder, f))

    for entry in tqdm(to_copy):
        shutil.copy(entry, dest)


if __name__ == "__main__":
    main()
