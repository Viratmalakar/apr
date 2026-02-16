from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ================================
# TIME CONVERSION FUNCTIONS
# ================================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        parts = str(t).split(":")
        if len(parts) != 3:
            return 0
        return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    except:
        return 0


def seconds_to_time(seconds):
    seconds = int(seconds)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02}:{m:02}:{s:02}"


# ================================
# HOME PAGE
# ================================

@app.route("/")
def index():
    return render_template("upload.html")


# ================================
# GENERATE REPORT
# ================================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    agent_path = os.path.join(UPLOAD_FOLDER, agent_file.filename)
    cdr_path = os.path.join(UPLOAD_FOLDER, cdr_file.filename)

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)

    # ============================
    # READ AGENT PERFORMANCE
    # ignore top 2 rows
    # ============================

    agent = pd.read_excel(agent_path, skiprows=2)

    agent.columns = agent.columns.str.strip()

    agent.replace("-", "00:00:00", inplace=True)

    agent["Employee ID"] = agent["Employee ID"].astype(str).str.strip()

    # ============================
    # BREAK CALCULATION
    # ============================

    agent["Total Break"] = (
        agent["LunchBreak"].apply(time_to_seconds)
        + agent["TeaBreak"].apply(time_to_seconds)
        + agent["ShortBreak"].apply(time_to_seconds)
    )

    agent["Total Meeting"] = (
        agent["Meeting"].apply(time_to_seconds)
        + agent["SystemDown"].apply(time_to_seconds)
    )

    agent["Total Login Seconds"] = agent["Total Login Time"].apply(time_to_seconds)

    agent["Net Login Seconds"] = (
        agent["Total Login Seconds"] - agent["Total Break"]
    )

    agent["Total Talk Seconds"] = agent["Total Talk Time"].apply(time_to_seconds)

    # ============================
    # READ CDR
    # ignore top 1 row
    # ============================

    cdr = pd.read_excel(cdr_path, skiprows=1)

    cdr.columns = cdr.columns.str.strip()

    cdr["Employee ID"] = cdr["Username"].astype(str).str.strip()

    cdr["CallType"] = cdr["CallType"].astype(str).str.upper()

    cdr["CallMatured"] = pd.to_numeric(cdr["CallMatured"], errors="coerce").fillna(0)

    cdr["Transfer"] = pd.to_numeric(cdr["Transfer"], errors="coerce").fillna(0)

    cdr["IsMature"] = (
        (cdr["CallMatured"] == 1) | (cdr["Transfer"] == 1)
    )

    cdr["IsIBMature"] = (
        (cdr["IsMature"]) & (cdr["CallType"] == "CSRINBOUND")
    )

    # ============================
    # GROUP CDR DATA
    # ============================

    total_mature = cdr.groupby("Employee ID")["IsMature"].sum().reset_index()

    ib_mature = cdr.groupby("Employee ID")["IsIBMature"].sum().reset_index()

    total_mature.columns = ["Employee ID", "Total Mature"]

    ib_mature.columns = ["Employee ID", "IB Mature"]

    calls = pd.merge(total_mature, ib_mature, on="Employee ID", how="outer").fillna(0)

    calls["OB Mature"] = calls["Total Mature"] - calls["IB Mature"]

    # ============================
    # MERGE AGENT + CDR
    # ============================

    final = pd.merge(agent, calls, on="Employee ID", how="left").fillna(0)

    # ============================
    # AHT CALCULATION
    # ============================

    final["AHT"] = final.apply(
        lambda x:
        seconds_to_time(
            int(x["Total Talk Seconds"] / x["Total Mature"])
        ) if x["Total Mature"] > 0 else "00:00:00",
        axis=1
    )

    # ============================
    # FINAL FORMAT
    # ============================

    final["Total Login"] = final["Total Login Seconds"].apply(seconds_to_time)

    final["Total Net Login"] = final["Net Login Seconds"].apply(seconds_to_time)

    final["Total Break"] = final["Total Break"].apply(seconds_to_time)

    final["Total Meeting"] = final["Total Meeting"].apply(seconds_to_time)

    final = final[
        [
            "Employee ID",
            "Agent Full Name",
            "Total Login",
            "Total Net Login",
            "Total Break",
            "Total Meeting",
            "AHT",
            "Total Mature",
            "IB Mature",
            "OB Mature",
        ]
    ]

    final.rename(columns={
        "Employee ID": "Employee ID",
        "Agent Full Name": "Agent Name"
    }, inplace=True)

    data = final.to_dict(orient="records")

    return render_template("result.html", data=data)


# ================================
# RUN
# ================================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
