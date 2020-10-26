from uuid import uuid4

NUM_IDS = 119


def main():
    ids = list()

    while len(ids) < NUM_IDS:
        id = uuid4()
        if id not in ids:
            ids.append(id)

    with open("ids.txt", "w") as idfile:
        idfile.write("\n".join(map(lambda x: str(x), ids)))


if __name__ == "__main__":
    main()
