from flask import Flask, render_template, request, redirect
import pandas as pd
import os

app = Flask(__name__)

# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(val):
    try:
        if pd.isna(val):
            return 0
        h, m, s = map(int, str(val).split(":"))
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
    # LOAD AGENT PERFORMANCE
    # header row = 3
    # =========================

    agent = pd.read_excel(agent_file, header=2)

    agent = agent.fillna("0")
    agent = agent.replace("-", "0")

    agent["Employee ID"] = agent.iloc[:,1].astype(str).str.strip()
    agent["Agent Name"] = agent.iloc[:,2]

    agent["Total Login"] = agent.iloc[:,3]

    # Break calculation
    T = agent.iloc[:,19]
    W = agent.iloc[:,22]
    Y = agent.iloc[:,24]

    agent["Total Break"] = (
        T.apply(time_to_seconds) +
        W.apply(time_to_seconds) +
        Y.apply(time_to_seconds)
    ).apply(seconds_to_time)

    # Meeting calculation
    U = agent.iloc[:,20]
    X = agent.iloc[:,23]

    agent["Total Meeting"] = (
        U.apply(time_to_seconds) +
        X.apply(time_to_seconds)
    ).apply(seconds_to_time)

    # Net login
    agent["Total Net Login"] = (
        agent["Total Login"].apply(time_to_seconds)
        - agent["Total Break"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # Talk time
    agent["Total Talk Time"] = agent.iloc[:,5]

    # =========================
    # LOAD CDR
    # header row = 2
    # =========================

    cdr = pd.read_excel(cdr_file, header=1)

    cdr["Employee ID"] = cdr.iloc[:,1].astype(str).str.strip()

    cdr["Call Type"] = cdr.iloc[:,6].astype(str).str.upper()

    cdr["Call Status"] = cdr.iloc[:,25].astype(str).str.upper()

    # =========================
    # MATURE COUNT
    # =========================

    mature = cdr[
        (cdr["Call Status"] == "CALLMATURED") |
        (cdr["Call Status"] == "TRANSFER")
    ]

    total_mature = mature.groupby("Employee ID").size()

    ib_mature = mature[
        mature["Call Type"] == "CSRINBOUND"
    ].groupby("Employee ID").size()

    # =========================
    # FINAL MERGE
    # =========================

    result = pd.DataFrame()

    result["Employee ID"] = agent["Employee ID"]

    result["Agent Name"] = agent["Agent Name"]

    result["Total Login"] = agent["Total Login"]

    result["Total Net Login"] = agent["Total Net Login"]

    result["Total Break"] = agent["Total Break"]

    result["Total Meeting"] = agent["Total Meeting"]

    result["Total Mature"] = result["Employee ID"].map(total_mature).fillna(0).astype(int)

    result["IB Mature"] = result["Employee ID"].map(ib_mature).fillna(0).astype(int)

    result["OB Mature"] = (
        result["Total Mature"] - result["IB Mature"]
    )

    # =========================
    # AHT
    # =========================

    talk_sec = agent["Total Talk Time"].apply(time_to_seconds)

    result["AHT"] = (
        talk_sec /
        result["Total Mature"].replace(0,1)
    ).apply(seconds_to_time)

    # =========================
    # FINAL FORMAT
    # =========================

    result = result[[
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
    ]]

    data = result.to_dict("records")

    return render_template(
        "result.html",
        data=data
    )


# =========================
# RENDER PORT FIX
# =========================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
