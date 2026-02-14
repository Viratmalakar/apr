from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ================= TIME FUNCTIONS =================

def time_to_seconds(t):

    try:

        if pd.isna(t):
            return 0

        t = str(t).strip()

        h, m, s = t.split(":")

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


# ================= HOME =================

@app.route("/")
def index():

    return render_template("index.html")


# ================= GENERATE =================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    agent_path = os.path.join(UPLOAD_FOLDER, agent_file.filename)
    cdr_path = os.path.join(UPLOAD_FOLDER, cdr_file.filename)

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)


    # =====================================================
    # AGENT PERFORMANCE REPORT
    # =====================================================

    agent = pd.read_excel(agent_path, header=2)

    # Column B Employee ID
    agent["Employee ID"] = agent.iloc[:,1].astype(str).str.strip()

    # Column C Agent Name
    agent["Agent Full Name"] = agent.iloc[:,2].astype(str).str.strip()

    # Column D Login Time
    agent["login_sec"] = agent.iloc[:,3].apply(time_to_seconds)

    # Column F Talk Time
    agent["talk_sec"] = agent.iloc[:,5].apply(time_to_seconds)

    # Column Z Break
    agent["break_sec"] = agent.iloc[:,25].apply(time_to_seconds)

    # Column U + X Meeting
    agent["meeting_sec"] = (
        agent.iloc[:,20].apply(time_to_seconds)
        +
        agent.iloc[:,23].apply(time_to_seconds)
    )

    # Net Login
    agent["net_sec"] = agent["login_sec"] - agent["break_sec"]



    # =====================================================
    # CDR REPORT
    # =====================================================

    cdr = pd.read_excel(cdr_path, header=2)

    # Column B Employee ID
    cdr["Employee ID"] = (
        cdr.iloc[:,1]
        .astype(str)
        .str.strip()
    )

    # Column G Campaign
    cdr["Campaign"] = (
        cdr.iloc[:,6]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Column Z Mature
    cdr["matured"] = pd.to_numeric(
        cdr.iloc[:,25],
        errors="coerce"
    ).fillna(0)


    # ================= TOTAL MATURE COUNT =================

    total_mature = cdr[
        cdr["matured"] > 0
    ].groupby(
        "Employee ID"
    ).size().reset_index(
        name="Total Mature"
    )


    # ================= IB MATURE COUNT =================

    ib = cdr[
        (cdr["matured"] > 0)
        &
        (
            cdr["Campaign"]
            .str.contains("CSRINBOUND", na=False)
        )
    ].groupby(
        "Employee ID"
    ).size().reset_index(
        name="IB Mature"
    )


    # =====================================================
    # MERGE
    # =====================================================

    final = agent.merge(
        total_mature,
        on="Employee ID",
        how="left"
    )

    final = final.merge(
        ib,
        on="Employee ID",
        how="left"
    )

    final.fillna(0, inplace=True)


    # OB Mature
    final["OB Mature"] = (
        final["Total Mature"]
        -
        final["IB Mature"]
    )


    # ================= AHT =================

    final["AHT_sec"] = (
        final["talk_sec"]
        /
        final["Total Mature"].replace(0,1)
    )


    # ================= CONVERT BACK TO TIME =================

    final["Total Login Time"] = (
        final["login_sec"]
        .apply(seconds_to_time)
    )

    final["Total Break"] = (
        final["break_sec"]
        .apply(seconds_to_time)
    )

    final["Total Net Login"] = (
        final["net_sec"]
        .apply(seconds_to_time)
    )

    final["Total Meeting"] = (
        final["meeting_sec"]
        .apply(seconds_to_time)
    )

    final["Total Talk Time"] = (
        final["talk_sec"]
        .apply(seconds_to_time)
    )

    final["AHT"] = (
        final["AHT_sec"]
        .apply(seconds_to_time)
    )


    # ================= FINAL OUTPUT =================

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
        table=final.to_html(index=False)
    )


# ================= RENDER PORT FIX =================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(
        host="0.0.0.0",
        port=port
    )
