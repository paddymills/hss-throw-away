
import os
from ping3 import ping

pcs = [
    ("W9245", "(Scott Aspril)"),
    ("W9129", "Michelle Brown"),
    ("W8790", "Sam Eveler"),
    ("W9248", "Steve Grilley"),
    ("W9188", "Keith Hakes"),
    ("W9246", "Amos Johnson"),
    ("W9249", "Nate Kirchner"),
    ("W9250", "Patrick Miller (workstation)"),
    ("W9103", "Patrick Miller (laptop)"),
    ("W9247", "(Mike Ortega)"),
    ("W9244", "Dan Painter"),
    ("W9251", "Jeff Reese"),
    ("W9243", "Jeff Vickers"),
    ("W9802", "Sam Eveler (laptop)"),
]

locs = [
    r"Program Files\SigmaTek",
    r"Program Files (x86)\SigmaTek",
]


for pc, owner in pcs:
    if not ping(pc):
        continue

    print("Searching {}...".format(pc))
    for loc in locs:
        try:
            for x in os.listdir(os.path.join(r"\\" + pc, "c$", loc)):
                if x != 'SigmaNEST 20 SP4':
                    print("\t{:30} | {:20}".format(loc, x))
        except FileNotFoundError:
            pass
