from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# =============================
# TIME FUNCTIONS
# =============================

def time_to_seconds(t):
    try:
        if pd.isna(t) or t == "-":
            return 0
        h, m, s = str(t).split(":")
        return int(h)*3600 + int(m)*60 + int(s)
    except:
        return 0


def seconds_to_time(sec):
    try:
        sec = int(sec)
        h = sec // 3600
        m = (sec % 3600) // 60
        s = sec % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


# =============================
# REMOVE MERGED EFFECT
# =============================

def clean_excel(file, skip_rows):

    df = pd.read_excel(
        file,
        skiprows=skip_rows,
        engine="openpyxl"
    )

    df = df.fillna("")
    df = df.replace("-", 0)

    return df


# =============================
# HOME
# =============================

@app.route("/")
def index():
    return render_template("index.html")


# =============================
# GENERATE REPORT
# =============================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    # =========================
    # LOAD AGENT PERFORMANCE
    # =========================

    agent = clean_excel(agent_file, 2)

    agent["Employee ID"] = agent.iloc[:, 1].astype(str)
    agent["Agent Name"] = agent.iloc[:, 2]

    agent["Total Login"] = agent.iloc[:, 3]

    lunch = agent.iloc[:, 19]
    short = agent.iloc[:, 22]
    tea = agent.iloc[:, 24]

    meeting = agent.iloc[:, 20]
    system = agent.iloc[:, 23]

    agent["Total Break"] = (
        lunch.apply(time_to_seconds) +
        short.apply(time_to_seconds) +
        tea.apply(time_to_seconds)
    ).apply(seconds_to_time)

    agent["Total Meeting"] = (
        meeting.apply(time_to_seconds) +
        system.apply(time_to_seconds)
    ).apply(seconds_to_time)

    agent["Total Net Login"] = (
        agent["Total Login"].apply(time_to_seconds)
        - (
            lunch.apply(time_to_seconds)
            + short.apply(time_to_seconds)
            + tea.apply(time_to_seconds)
        )
    ).apply(seconds_to_time)

    talk_time = agent.iloc[:, 5]
    agent["Talk Time Seconds"] = talk_time.apply(time_to_seconds)

    agent = agent[
        [
            "Employee ID",
            "Agent Name",
            "Total Login",
            "Total Net Login",
            "Total Break",
            "Total Meeting",
            "Talk Time Seconds",
        ]
    ]

    # =========================
    # LOAD CDR
    # =========================

    cdr = clean_excel(cdr_file, 1)

    cdr["Employee ID"] = cdr.iloc[:, 1].astype(str)

    cdr["Call Status"] = cdr.iloc[:, 25].astype(str).str.upper()
    cdr["Call Type"] = cdr.iloc[:, 6].astype(str).str.upper()

    mature = cdr[
        cdr["Call Status"].isin(["CALLMATURED", "TRANSFER"])
    ]

    ib = mature[
        mature["Call Type"] == "CSRINBOUND"
    ]

    total_mature = mature.groupby("Employee ID").size()
    ib_mature = ib.groupby("Employee ID").size()

    # =========================
    # MERGE DATA
    # =========================

    final = agent.copy()

    final["Total Mature"] = final["Employee ID"].map(total_mature).fillna(0).astype(int)

    final["IB Mature"] = final["Employee ID"].map(ib_mature).fillna(0).astype(int)

    final["OB Mature"] = final["Total Mature"] - final["IB Mature"]

    # =========================
    # AHT
    # =========================

    final["AHT"] = (
        final["Talk Time Seconds"] /
        final["Total Mature"].replace(0, 1)
    ).apply(seconds_to_time)

    # =========================
    # FINAL SELECT
    # =========================

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
            "OB Mature",
        ]
    ]

    final = final.sort_values("Total Mature", ascending=False)

    rows = final.to_dict("records")

    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")

    return render_template(
        "result.html",
        rows=rows,
        report_time=report_time
    )


# =============================
# RUN FOR RENDER
# =============================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(
        host="0.0.0.0",
        port=port
    )
