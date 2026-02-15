from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# ===============================
# TIME FUNCTIONS
# ===============================

def time_to_seconds(val):
    try:
        if pd.isna(val) or val == "-":
            return 0
        h, m, s = str(val).split(":")
        return int(h)*3600 + int(m)*60 + int(s)
    except:
        return 0


def seconds_to_time(sec):
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)
    return f"{h:02}:{m:02}:{s:02}"


# ===============================
# UNMERGE CELLS FUNCTION
# ===============================

def unmerge_excel(file):
    from openpyxl import load_workbook
    wb = load_workbook(file)
    ws = wb.active

    merged = list(ws.merged_cells)
    for m in merged:
        ws.unmerge_cells(str(m))

    temp = "temp_unmerged.xlsx"
    wb.save(temp)
    return temp


# ===============================
# HOME
# ===============================

@app.route("/")
def index():
    return render_template("index.html")


# ===============================
# GENERATE REPORT
# ===============================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    # ===========================
    # UNMERGE FIRST
    # ===========================

    agent_path = unmerge_excel(agent_file)
    cdr_path = unmerge_excel(cdr_file)

    # ===========================
    # LOAD AGENT PERFORMANCE
    # header row = row 3
    # ===========================

    agent = pd.read_excel(agent_path, header=2)

    agent.replace("-", "00:00:00", inplace=True)

    agent["Employee ID"] = agent.iloc[:,1].astype(str)
    agent["Agent Full Name"] = agent.iloc[:,2]

    agent["Total Login Time"] = agent.iloc[:,3]

    # BREAK
    lunch = agent.iloc[:,19]
    short = agent.iloc[:,22]
    tea = agent.iloc[:,24]

    agent["Total Break"] = (
        lunch.apply(time_to_seconds) +
        short.apply(time_to_seconds) +
        tea.apply(time_to_seconds)
    ).apply(seconds_to_time)

    # MEETING
    meet = agent.iloc[:,20]
    sysdown = agent.iloc[:,23]

    agent["Total Meeting"] = (
        meet.apply(time_to_seconds) +
        sysdown.apply(time_to_seconds)
    ).apply(seconds_to_time)

    # NET LOGIN
    agent["Total Net Login"] = (
        agent["Total Login Time"].apply(time_to_seconds) -
        agent["Total Break"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # TALK TIME
    agent["Total Talk Time"] = agent.iloc[:,5]


    # ===========================
    # LOAD CDR
    # header row = row 2
    # ===========================

    cdr = pd.read_excel(cdr_path, header=1)

    cdr["Employee ID"] = cdr.iloc[:,1].astype(str)
    cdr["CallType"] = cdr.iloc[:,6].astype(str).str.upper()
    cdr["Call Status"] = cdr.iloc[:,25].astype(str).str.upper()

    # ===========================
    # MATURE COUNT
    # ===========================

    mature = cdr[
        (cdr["Call Status"] == "CALLMATURED") |
        (cdr["Call Status"] == "TRANSFER")
    ]

    total_mature = mature.groupby("Employee ID").size()

    ib_mature = mature[
        mature["CallType"] == "CSRINBOUND"
    ].groupby("Employee ID").size()

    ob_mature = total_mature - ib_mature

    # ===========================
    # MERGE
    # ===========================

    final = agent.copy()

    final["Total Mature"] = final["Employee ID"].map(total_mature).fillna(0).astype(int)
    final["IB Mature"] = final["Employee ID"].map(ib_mature).fillna(0).astype(int)
    final["OB Mature"] = final["Employee ID"].map(ob_mature).fillna(0).astype(int)

    # ===========================
    # AHT
    # ===========================

    final["AHT"] = final.apply(
        lambda x: seconds_to_time(
            time_to_seconds(x["Total Talk Time"]) // x["Total Mature"]
        ) if x["Total Mature"] > 0 else "00:00:00",
        axis=1
    )

    # ===========================
    # SELECT FINAL COLUMNS
    # ===========================

    final = final[[
        "Employee ID",
        "Agent Full Name",
        "Total Login Time",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "AHT",
        "Total Mature",
        "IB Mature",
        "OB Mature"
    ]]

    rows = final.to_dict("records")

    return render_template("result.html", rows=rows)


# ===============================
# MAIN
# ===============================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
