import os
from tqdm import tqdm

PATH = r"\\hssieng\SNDataPrd\Parts"

def remove(ext=".QPR", remove_index=False):

    removals = 0

    for f in tqdm(os.listdir(PATH)):
        if f.endswith(ext):
            os.remove(os.path.join(PATH, f))
            removals += 1
        elif f == 'SNIndex.ind' and remove_index:
            os.remove(os.path.join(PATH, f))
        elif ".." in f or " " in f:
            print(f)
            

    print("{} files removed".format(removals))

# remove(ext=".QPR", remove_index=True)
remove(ext=".QPR")
remove(ext=".DEL")
