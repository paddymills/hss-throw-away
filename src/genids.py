from uuid import uuid4

NUM_NEW_IDS = 2


def main():
    with open("existing_ids.txt", "r") as idfile:
        ids = idfile.readlines()

    # if last line is not blank, add newline
    if ids[-1]:
        ids[-1] += "\n"
        
    TOTAL_NEEDED_IDS = len(ids) + NUM_NEW_IDS

    while len(ids) < TOTAL_NEEDED_IDS:
        id = uuid4()
        if id not in ids:
            ids.append(str(id) + "\n")

    with open("ids.txt", "w") as idfile:
        idfile.writelines(ids)


if __name__ == "__main__":
    main()
