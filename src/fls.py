
import os

from zipfile import ZipFile
from tqdm import tqdm

base_dir = r"\\hssfileserv3\THC\3DScanning\FARO\1200131 Banker Steel\sdcard_08-24-2021"

scans = os.path.join(base_dir, "Scans")
dest = os.path.join(base_dir, "fls")

class Zipper:

    def zip(self, name):
        self.zf = ZipFile(os.path.join(dest, folder), 'w')

        self.add_to_zip('Scans')
        self.add_to_zip('Main')
        self.add_to_zip('SHA256SUM')
        self.add_to_zip('SHA256SUM.sha')
        self.add_to_zip('SHA256SUM.sig')

        self.zf.close()


    def add_to_zip(self, folder):
        if os.path.isfile(folder):
            self.zf.write(folder)
            return

        for file in os.listdir(folder):
            full_path = os.path.join(folder, file)

            if os.path.isfile(full_path):
                self.zf.write(full_path)
            
            elif os.path.isdir(full_path):
                self.add_to_zip(full_path)


z = Zipper()
for folder in tqdm(os.listdir(scans), desc="Zipping files"):
    os.chdir(os.path.join(scans, folder))

    z.zip(folder)