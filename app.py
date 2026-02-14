from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        h, m, s = map(int, str(t).split(":"))
        return h*3600 + m*60 + s
    except:
        return 0


def seconds_to_time(sec):
    sec = int(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h:02}:{m:02}:{s:02}"


# =========================
# SAFE HEADER DETECT
# =========================

def detect_header(df):

    for i in range(min(10, len(df))):

        row = df.iloc[i].astype(str).str.lower()

        if row.str.contains("agent").any() or row.str.contains("employee").any():

            df.columns = df.iloc[i]
            df = df[i+1:]
            break

    return df.reset_index(drop=True)


# =========================
# HOME
# =========================

@app.route("/")
def index():
    return render_template("index.html")


# =========================
# GENERATE REPORT
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    # =========================
    # LOAD AGENT REPORT
    # =========================

    agent = pd.read_excel(agent_file, header=None)
    agent = detect_header(agent)

    agent.columns = agent.columns.astype(str).str.strip()

    # Employee ID from column B (Agent Name)
    agent["Employee ID"] = agent.iloc[:,1].astype(str).str.strip()

    agent["Agent Name"] = agent.iloc[:,2]

    agent["Total Login Time"] = agent.iloc[:,3]

    agent["Total Net Login"] = agent.iloc[:,4]

    agent["Total Break"] = agent.iloc[:,5]

    # Total Meeting = column U + column X
    meeting = pd.to_numeric(agent.iloc[:,20], errors="coerce").fillna(0)
    systemdown = pd.to_numeric(agent.iloc[:,23], errors="coerce").fillna(0)

    agent["Total Meeting"] = meeting + systemdown

    # Total Talk Time column F
    agent["Total Talk Time"] = agent.iloc[:,5]

    # =========================
    # LOAD CDR
    # =========================

    cdr = pd.read_excel(cdr_file, header=None)
    cdr = detect_header(cdr)

    cdr.columns = cdr.columns.astype(str).str.strip()

    # Employee ID from column B
    cdr["Employee ID"] = cdr.iloc[:,1].astype(str).str.strip()

    # Campaign column G
    cdr["Campaign"] = cdr.iloc[:,6].astype(str).str.upper()

    # Column Z = CallMatured+Transfer result
    matured_flag = pd.to_numeric(cdr.iloc[:,25], errors="coerce").fillna(0)

    cdr["MaturedFlag"] = matured_flag

    # =========================
    # COUNT CALCULATIONS
    # =========================

    total_mature = (
        cdr[cdr["MaturedFlag"] > 0]
        .groupby("Employee ID")
        .size()
        .reset_index(name="Total Mature")
    )

    ib_mature = (
        cdr[(cdr["MaturedFlag"] > 0) &
            (cdr["Campaign"]=="CSRINBOUND")]
        .groupby("Employee ID")
        .size()
        .reset_index(name="IB Mature")
    )

    calls = total_mature.merge(ib_mature, on="Employee ID", how="left")

    calls = calls.fillna(0)

    calls["OB Mature"] = calls["Total Mature"] - calls["IB Mature"]

    calls["Total Mature"] = calls["Total Mature"].astype(int)
    calls["IB Mature"] = calls["IB Mature"].astype(int)
    calls["OB Mature"] = calls["OB Mature"].astype(int)

    # =========================
    # MERGE WITH AGENT
    # =========================

    final = agent.merge(calls, on="Employee ID", how="left")

    final = final.fillna(0)

    final["Total Mature"] = final["Total Mature"].astype(int)
    final["IB Mature"] = final["IB Mature"].astype(int)
    final["OB Mature"] = final["OB Mature"].astype(int)

    # =========================
    # AHT
    # =========================

    final["TalkSec"] = final["Total Talk Time"].apply(time_to_seconds)

    final["AHTSec"] = final.apply(
        lambda x: x["TalkSec"]//x["Total Mature"]
        if x["Total Mature"]>0 else 0,
        axis=1
    )

    final["AHT"] = final["AHTSec"].apply(seconds_to_time)

    # =========================
    # FINAL OUTPUT
    # =========================

    final = final[[
        "Agent Name",
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

    final.columns = [
        "Agent Name",
        "Agent Full Name",
        "Total Login Time",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "AHT",
        "Total Mature",
        "IB Mature",
        "OB Mature"
    ]

    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")

    return render_template(
        "result.html",
        rows=final.to_dict("records"),
        report_time=report_time
    )


# =========================
# RENDER PORT FIX
# =========================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)
