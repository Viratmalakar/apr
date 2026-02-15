from flask import Flask, render_template, request
import pandas as pd
import io

app = Flask(__name__)

# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(val):
    try:
        if pd.isna(val):
            return 0
        val = str(val)
        h, m, s = val.split(":")
        return int(h)*3600 + int(m)*60 + int(s)
    except:
        return 0


def seconds_to_time(sec):
    sec = int(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h:02}:{m:02}:{s:02}"


# =========================
# UNMERGE CELLS
# =========================

def unmerge_excel(file):
    from openpyxl import load_workbook

    wb = load_workbook(file)
    ws = wb.active

    merged = list(ws.merged_cells.ranges)

    for merge in merged:
        ws.unmerge_cells(str(merge))

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return stream


# =========================
# HOME
# =========================

@app.route("/")
def home():
    return render_template("index.html")


# =========================
# GENERATE REPORT
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    # =====================
    # UNMERGE FIRST
    # =====================

    agent_stream = unmerge_excel(agent_file)
    cdr_stream = unmerge_excel(cdr_file)

    # =====================
    # LOAD DATA
    # =====================

    agent = pd.read_excel(agent_stream, header=2)
    cdr = pd.read_excel(cdr_stream, header=1)

    agent.columns = agent.columns.str.strip()
    cdr.columns = cdr.columns.str.strip()

    # replace dash
    agent = agent.replace("-", 0)

    # =====================
    # AGENT PERFORMANCE
    # =====================

    agent["Employee ID"] = agent.iloc[:, 1].astype(str)

    agent["Agent Name"] = agent.iloc[:, 2]

    login_sec = agent.iloc[:, 3].apply(time_to_seconds)

    lunch = agent.iloc[:, 19].apply(time_to_seconds)
    meeting = agent.iloc[:, 20].apply(time_to_seconds)
    short = agent.iloc[:, 22].apply(time_to_seconds)
    systemdown = agent.iloc[:, 23].apply(time_to_seconds)
    tea = agent.iloc[:, 24].apply(time_to_seconds)

    agent["Total Break"] = lunch + short + tea

    agent["Total Meeting"] = meeting + systemdown

    agent["Total Net Login"] = login_sec - agent["Total Break"]

    talk_time = agent.iloc[:, 5].apply(time_to_seconds)

    agent["Talk Time Sec"] = talk_time

    # =====================
    # CDR PROCESS
    # =====================

    cdr["Employee ID"] = cdr.iloc[:, 1].astype(str)

    call_status = cdr.iloc[:, 25].astype(str).str.upper()

    campaign = cdr.iloc[:, 6].astype(str).str.upper()

    cdr["MATURED"] = call_status.isin(["CALLMATURED", "TRANSFER"])

    cdr["IB"] = cdr["MATURED"] & (campaign == "CSRINBOUND")

    total = cdr.groupby("Employee ID")["MATURED"].sum().astype(int)

    ib = cdr.groupby("Employee ID")["IB"].sum().astype(int)

    ob = total - ib

    cdr_result = pd.DataFrame({
        "Employee ID": total.index,
        "Total Mature": total.values,
        "IB Mature": ib.values,
        "OB Mature": ob.values
    })

    # =====================
    # MERGE
    # =====================

    final = agent.merge(cdr_result, on="Employee ID", how="left")

    final = final.fillna(0)

    final["Total Mature"] = final["Total Mature"].astype(int)
    final["IB Mature"] = final["IB Mature"].astype(int)
    final["OB Mature"] = final["OB Mature"].astype(int)

    # =====================
    # AHT
    # =====================

    final["AHT"] = final.apply(
        lambda x: seconds_to_time(
            x["Talk Time Sec"] // x["Total Mature"]
        ) if x["Total Mature"] > 0 else "00:00:00",
        axis=1
    )

    # =====================
    # FORMAT TIMES
    # =====================

    final["Total Login Time"] = login_sec.apply(seconds_to_time)

    final["Total Break"] = final["Total Break"].apply(seconds_to_time)

    final["Total Meeting"] = final["Total Meeting"].apply(seconds_to_time)

    final["Total Net Login"] = final["Total Net Login"].apply(seconds_to_time)

    # =====================
    # FINAL OUTPUT
    # =====================

    final = final[[
        "Employee ID",
        "Agent Name",
        "Total Login Time",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "AHT",
        "Total Mature",
        "IB Mature",
        "OB Mature"
    ]]

    return render_template(
        "result.html",
        rows=final.to_dict("records")
    )


# =========================
# RUN
# =========================

import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

