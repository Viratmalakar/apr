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

        t = str(t)

        parts = t.split(":")

        if len(parts) != 3:
            return 0

        h, m, s = map(int, parts)

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
# GENERATE
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    # =========================
    # LOAD AGENT REPORT
    # ignore first 2 rows
    # =========================

    agent = pd.read_excel(
        agent_file,
        skiprows=2,
        dtype=str
    )

    agent.fillna("", inplace=True)

    agent["Employee ID"] = agent.iloc[:,1].str.strip()

    agent["Agent Full Name"] = agent.iloc[:,1]

    agent["Total Login Time"] = agent.iloc[:,3]

    agent["Total Net Login"] = agent.iloc[:,4]

    agent["Total Break"] = agent.iloc[:,5]

    meeting_u = pd.to_timedelta(agent.iloc[:,20], errors="coerce").fillna(pd.Timedelta(0))
    meeting_x = pd.to_timedelta(agent.iloc[:,23], errors="coerce").fillna(pd.Timedelta(0))

    agent["Total Meeting"] = meeting_u + meeting_x

    agent["Total Talk Time"] = agent.iloc[:,5]


    # =========================
    # LOAD CDR REPORT
    # ignore first row
    # =========================

    cdr = pd.read_excel(
        cdr_file,
        skiprows=1,
        dtype=str
    )

    cdr.fillna("", inplace=True)

    cdr["Employee ID"] = cdr.iloc[:,1].str.strip()

    cdr["CallMatured"] = pd.to_numeric(cdr.iloc[:,25], errors="coerce").fillna(0)

    cdr["Transfer"] = pd.to_numeric(cdr.iloc[:,26], errors="coerce").fillna(0)

    cdr["Campaign"] = cdr.iloc[:,6].str.upper()


    # Mature condition
    cdr["is_mature"] = (
        (cdr["CallMatured"] == 1) |
        (cdr["Transfer"] == 1)
    ).astype(int)


    # =========================
    # COUNT
    # =========================

    total = cdr.groupby("Employee ID")["is_mature"].sum().reset_index()

    total.rename(columns={"is_mature":"Total Mature"}, inplace=True)


    ib = (
        cdr[cdr["Campaign"]=="CSRINBOUND"]
        .groupby("Employee ID")["is_mature"]
        .sum()
        .reset_index()
    )

    ib.rename(columns={"is_mature":"IB Mature"}, inplace=True)


    summary = total.merge(ib, on="Employee ID", how="left")

    summary["IB Mature"] = summary["IB Mature"].fillna(0).astype(int)

    summary["OB Mature"] = summary["Total Mature"] - summary["IB Mature"]


    # =========================
    # MERGE
    # =========================

    final = agent.merge(summary, on="Employee ID", how="left")

    final.fillna(0, inplace=True)


    final["Total Mature"] = final["Total Mature"].astype(int)

    final["IB Mature"] = final["IB Mature"].astype(int)

    final["OB Mature"] = final["OB Mature"].astype(int)


    # =========================
    # AHT
    # =========================

    def calc_aht(row):

        talk = time_to_seconds(row["Total Talk Time"])

        mature = row["Total Mature"]

        if mature == 0:
            return "00:00:00"

        return seconds_to_time(talk/mature)


    final["AHT"] = final.apply(calc_aht, axis=1)


    # =========================
    # FINAL
    # =========================

    final = final[[
        "Employee ID",
        "Agent Full Name",
        "Total Login Time",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "Total Mature",
        "IB Mature",
        "OB Mature",
        "Total Talk Time",
        "AHT"
    ]]


    return render_template(
        "result.html",
        rows=final.to_dict("records"),
        report_time=datetime.now().strftime("%d %b %Y %I:%M %p")
    )


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run()
