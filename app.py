import os
import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)


# ======================
# TIME FUNCTIONS
# ======================

def time_to_seconds(val):

    try:

        if pd.isna(val) or val == "-" or val == "":
            return 0

        h, m, s = str(val).split(":")

        return int(h)*3600 + int(m)*60 + int(s)

    except:

        return 0


def seconds_to_time(sec):

    sec = int(sec)

    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60

    return f"{h:02}:{m:02}:{s:02}"


# ======================
# HOME
# ======================

@app.route("/")
def home():

    return render_template("index.html")


# ======================
# GENERATE REPORT
# ======================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    # read files directly (NO openpyxl save)

    agent = pd.read_excel(agent_file, header=2, engine="openpyxl")

    cdr = pd.read_excel(cdr_file, header=1, engine="openpyxl")

    agent = agent.fillna(0)

    agent.replace("-", 0, inplace=True)

    cdr = cdr.fillna("")


    # ======================
    # AGENT DATA
    # ======================

    df = pd.DataFrame()

    df["Employee ID"] = agent.iloc[:, 1].astype(str)

    df["Agent Full Name"] = agent.iloc[:, 2]

    df["Total Login Time"] = agent.iloc[:, 3]

    df["Total Talk Time"] = agent.iloc[:, 5]


    lunch = agent.iloc[:, 19]

    meeting = agent.iloc[:, 20]

    short = agent.iloc[:, 22]

    system = agent.iloc[:, 23]

    tea = agent.iloc[:, 24]


    break_sec = (

        lunch.apply(time_to_seconds)

        + short.apply(time_to_seconds)

        + tea.apply(time_to_seconds)

    )


    meeting_sec = (

        meeting.apply(time_to_seconds)

        + system.apply(time_to_seconds)

    )


    login_sec = agent.iloc[:, 3].apply(time_to_seconds)

    net_sec = login_sec - break_sec


    df["Total Break"] = break_sec.apply(seconds_to_time)

    df["Total Meeting"] = meeting_sec.apply(seconds_to_time)

    df["Total Net Login"] = net_sec.apply(seconds_to_time)


    # ======================
    # CDR CALCULATION
    # ======================

    cdr["Employee ID"] = cdr.iloc[:, 1].astype(str)

    cdr["CallType"] = cdr.iloc[:, 6]

    cdr["CallStatus"] = cdr.iloc[:, 25]


    matured = cdr[

        cdr["CallStatus"].isin(["CALLMATURED", "TRANSFER"])

    ]


    total_mature = matured.groupby("Employee ID").size()


    ib = matured[

        matured["CallType"] == "CSRINBOUND"

    ].groupby("Employee ID").size()


    df["Total Mature"] = df["Employee ID"].map(total_mature).fillna(0).astype(int)

    df["IB Mature"] = df["Employee ID"].map(ib).fillna(0).astype(int)

    df["OB Mature"] = df["Total Mature"] - df["IB Mature"]


    df["AHT"] = "00:00:00"

