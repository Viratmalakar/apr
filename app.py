from flask import Flask, render_template, request
import pandas as pd
from datetime import datetime

app = Flask(__name__)


# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        parts = str(t).split(":")
        if len(parts) != 3:
            return 0
        h, m, s = map(int, parts)
        return h*3600 + m*60 + s
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


# =========================
# HOME PAGE
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
    # LOAD AGENT FILE
    # =========================

    agent = pd.read_excel(agent_file, header=None)

    # Employee ID → Column B
    agent["Employee ID"] = agent.iloc[:, 1].astype(str).str.strip()

    # Agent Full Name → Column C
    agent["Agent Name"] = agent.iloc[:, 2]

    # Total Login → Column D
    agent["Total Login Time"] = agent.iloc[:, 3]

    # Total Talk Time → Column F
    agent["Total Talk Time"] = agent.iloc[:, 5]

    # Total Break → Column T + W + Y
    agent["Break1"] = agent.iloc[:, 19]
    agent["Break2"] = agent.iloc[:, 22]
    agent["Break3"] = agent.iloc[:, 24]

    agent["Total Break"] = (
        agent["Break1"].apply(time_to_seconds) +
        agent["Break2"].apply(time_to_seconds) +
        agent["Break3"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # Total Meeting → Column U + X
    agent["Meet1"] = agent.iloc[:, 20]
    agent["Meet2"] = agent.iloc[:, 23]

    agent["Total Meeting"] = (
        agent["Meet1"].apply(time_to_seconds) +
        agent["Meet2"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # Net Login
    agent["Total Net Login"] = (
        agent["Total Login Time"].apply(time_to_seconds) -
        agent["Total Break"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # =========================
    # LOAD CDR FILE
    # =========================

    cdr = pd.read_excel(cdr_file, header=None)

    # Employee ID → Column B
    cdr["Employee ID"] = cdr.iloc[:, 1].astype(str).str.strip()

    # Campaign → Column G
    cdr["Campaign"] = cdr.iloc[:, 6].astype(str).str.upper().str.strip()

    # Call Status → Column Z
    cdr["Status"] = cdr.iloc[:, 25].astype(str).str.upper().str.strip()

    # Mature Calls → CALLMATURED + TRANSFER
    mature = cdr[
        (cdr["Status"] == "CALLMATURED") |
        (cdr["Status"] == "TRANSFER")
    ]

    total_mature = mature.groupby("Employee ID").size().reset_index(name="Total Mature")

    ib_mature = mature[
        mature["Campaign"] == "CSRINBOUND"
    ].groupby("Employee ID").size().reset_index(name="IB Mature")

    calls = total_mature.merge(ib_mature, on="Employee ID", how="left").fillna(0)

    calls["Total Mature"] = calls["Total Mature"].astype(int)
    calls["IB Mature"] = calls["IB Mature"].astype(int)

    calls["OB Mature"] = (
        calls["Total Mature"] -
        calls["IB Mature"]
    ).astype(int)

    # =========================
    # MERGE BOTH FILES
    # =========================

    final = agent.merge(calls, on="Employee ID", how="left").fillna(0)

    final["Total Mature"] = final["Total Mature"].astype(int)
    final["IB Mature"] = final["IB Mature"].astype(int)
    final["OB Mature"] = final["OB Mature"].astype(int)

    # =========================
    # CALCULATE AHT
    # =========================

    final["AHT"] = final.apply(
        lambda x:
        seconds_to_time(
            time_to_seconds(x["Total Talk Time"]) //
            x["Total Mature"]
        )
        if x["Total Mature"] > 0 else "00:00:00",
        axis=1
    )

    # =========================
    # FINAL COLUMNS
    # =========================

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

    # =========================
    # SORT BY NET LOGIN
    # =========================

    final["sort"] = final["Total Net Login"].apply(time_to_seconds)

    final = final.sort_values("sort", ascending=False)

    final = final.drop("sort", axis=1)

    # =========================
    # RENDER
    # =========================

    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")

    return render_template(
        "result.html",
        rows=final.to_dict("records"),
        report_time=report_time
    )


# =========================
# MAIN
# =========================

if __name__ == "__main__":
    app.run()
