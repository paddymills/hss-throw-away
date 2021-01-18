import os
import shutil
import sndb
from tqdm import tqdm

def main():
    parts_dir = r"\\hssieng\SNDataPrd\PARTS"

    jobs = get_jobs()
    parts_list = get_parts_list()
    parts_to_move = list()
    keep = 0
    for part_file in os.listdir(parts_dir):
        if not part_file.endswith(".PRS"):
            continue

        part = part_file.replace(".PRS", "")

        if part[:8].upper() in jobs:
            keep += 1
        elif part in parts_list:
            keep += 1
        else:
            parts_to_move.append(part_file)

    for part in tqdm(parts_to_move, desc="Moving parts to Archive"):
        from_path = os.path.join(parts_dir, part)
        to_path = os.path.join(parts_dir, "Archive", part)

        shutil.move(from_path, to_path)


def get_parts_list():
    conn = sndb.get_sndb_conn()
    cursor = conn.cursor()
    cursor.execute("SELECT PartName FROM Part")

    parts_list = [x[0].upper() for x in cursor.fetchall()]

    conn.close()

    return parts_list

def get_jobs():
    with open("jobs.txt", "r") as jobs_file:
        return jobs_file.read().split("\n")


if __name__ == "__main__":
    main()