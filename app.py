from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        h, m, s = str(t).split(":")
        return int(h)*3600 + int(m)*60 + int(s)
    except:
        return 0


def seconds_to_time(sec):
    sec = int(sec)
    h = sec//3600
    m = (sec%3600)//60
    s = sec%60
    return f"{h:02}:{m:02}:{s:02}"


# =========================
# HOME
# =========================

@app.route("/")
def index():
    return render_template("upload.html")


# =========================
# GENERATE
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    agent_path = os.path.join(UPLOAD_FOLDER, agent_file.filename)
    cdr_path = os.path.join(UPLOAD_FOLDER, cdr_file.filename)

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)

    # ======================
    # READ AGENT FILE
    # ======================

    agent = pd.read_excel(agent_path, skiprows=2)

    # remove hidden spaces
    agent.columns = agent.columns.str.strip()

    # AUTO DETECT REQUIRED COLUMNS
    emp_col = None
    name_col = None

    for col in agent.columns:
        if "employee" in col.lower():
            emp_col = col
        if "agent full name" in col.lower():
            name_col = col

    if emp_col is None or name_col is None:
        return "Employee ID or Agent Name column not found in Agent file"

    agent["Employee ID"] = agent[emp_col].astype(str).str.strip()
    agent["Agent Name"] = agent[name_col]

    agent.replace("-", "00:00:00", inplace=True)

    # break
    agent["Total Break"] = (
        agent["LunchBreak"].apply(time_to_seconds)
        + agent["TeaBreak"].apply(time_to_seconds)
        + agent["ShortBreak"].apply(time_to_seconds)
    )

    # meeting
    agent["Total Meeting"] = (
        agent["Meeting"].apply(time_to_seconds)
        + agent["SystemDown"].apply(time_to_seconds)
    )

    # login
    agent["Login Sec"] = agent["Total Login Time"].apply(time_to_seconds)

    agent["Net Login Sec"] = agent["Login Sec"] - agent["Total Break"]

    agent["Talk Sec"] = agent["Total Talk Time"].apply(time_to_seconds)


    # ======================
    # READ CDR
    # ======================

    cdr = pd.read_excel(cdr_path, skiprows=1)

    cdr.columns = cdr.columns.str.strip()

    cdr["Employee ID"] = cdr["Username"].astype(str).str.strip()

    cdr["CallMatured"] = pd.to_numeric(cdr["CallMatured"], errors="coerce").fillna(0)

    cdr["Transfer"] = pd.to_numeric(cdr["Transfer"], errors="coerce").fillna(0)

    cdr["CallType"] = cdr["CallType"].astype(str)

    cdr["IsMature"] = (
        (cdr["CallMatured"] == 1)
        | (cdr["Transfer"] == 1)
    )

    cdr["IsIB"] = (
        (cdr["IsMature"])
        & (cdr["CallType"] == "CSRINBOUND")
    )

    total = cdr.groupby("Employee ID")["IsMature"].sum().reset_index()
    ib = cdr.groupby("Employee ID")["IsIB"].sum().reset_index()

    total.columns = ["Employee ID","Total Mature"]
    ib.columns = ["Employee ID","IB Mature"]

    calls = pd.merge(total, ib, on="Employee ID", how="outer").fillna(0)

    calls["OB Mature"] = calls["Total Mature"] - calls["IB Mature"]


    # ======================
    # MERGE
    # ======================

    final = pd.merge(agent, calls, on="Employee ID", how="left").fillna(0)


    # ======================
    # AHT
    # ======================

    final["AHT"] = final.apply(
        lambda x: seconds_to_time(x["Talk Sec"]/x["Total Mature"])
        if x["Total Mature"]>0 else "00:00:00",
        axis=1
    )


    # ======================
    # FINAL FORMAT
    # ======================

    final["Total Login"] = final["Login Sec"].apply(seconds_to_time)

    final["Total Net Login"] = final["Net Login Sec"].apply(seconds_to_time)

    final["Total Break"] = final["Total Break"].apply(seconds_to_time)

    final["Total Meeting"] = final["Total Meeting"].apply(seconds_to_time)

    final = final[
        [
            "Employee ID",
            "Agent Name",
            "Total Login",
            "Total Net Login",
            "Total Break",
            "Total Meeting",
            "AHT",
            "Total Mature",
            "IB Mature",
            "OB Mature"
        ]
    ]

    data = final.to_dict(orient="records")

    return render_template("result.html", data=data)


# =========================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
