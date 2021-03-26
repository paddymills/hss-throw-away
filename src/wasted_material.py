
import sndb

LBS_PER_IN2 = 0.425

with open('progs.txt', 'r') as progs_file:
    progs = list()
    for line in progs_file.readlines():
        if line:
            progs.append(line.strip())

with sndb.get_sndb_conn() as db:
    crsr = db.cursor()

    area = dict(total=0, burned=0, open=0, deleted=0)

    for prog in progs:
        crsr.execute("""
            SELECT TOP 1 UsedArea, TransType
            FROM ProgArchive
            WHERE ProgramName=?
            ORDER BY AutoID DESC
        """, prog)

        p_area, transtype = crsr.fetchone()
        area['total'] += p_area
        if transtype == 'SN100':
            area['open'] += p_area
        elif transtype == 'SN101':
            area['deleted'] += p_area
        elif transtype == 'SN102':
            area['burned'] += p_area

print("     Open:{:16.4f} LBS".format( area['open']    * LBS_PER_IN2) )
print("  Deleted:{:16.4f} LBS".format( area['deleted'] * LBS_PER_IN2) )
print("   Burned:{:16.4f} LBS".format( area['burned']  * LBS_PER_IN2) )
print("================================")
print("    Total:{:16.4f} LBS".format( area['total']   * LBS_PER_IN2) )
