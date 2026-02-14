from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ---------- TIME FUNCTIONS ----------

def time_to_seconds(t):

    try:
        if pd.isna(t):
            return 0

        parts = str(t).split(":")
        return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])

    except:
        return 0


def seconds_to_time(sec):

    sec = int(sec)

    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60

    return f"{h:02}:{m:02}:{s:02}"


# ---------- HOME ----------

@app.route("/")
def index():
    return render_template("index.html")


# ---------- GENERATE REPORT ----------

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

    # Column B = Employee ID
    agent["Employee ID"] = agent.iloc[:,1].astype(str)

    # Column C = Agent Full Name
    agent["Agent Full Name"] = agent.iloc[:,2].astype(str)

    # Column D = Total Login Time
    agent["login_sec"] = agent.iloc[:,3].apply(time_to_seconds)

    # Column F = Total Talk Time
    agent["talk_sec"] = agent.iloc[:,5].apply(time_to_seconds)

    # Column Z = Total Break
    agent["break_sec"] = agent.iloc[:,25].apply(time_to_seconds)

    # Column U + Column X = Total Meeting
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

    # Column B = Employee ID
    cdr["Employee ID"] = cdr.iloc[:,1].astype(str)

    # Column G = Campaign
    cdr["Campaign"] = cdr.iloc[:,6].astype(str)

    # Column Z = Matured
    cdr["matured"] = pd.to_numeric(
        cdr.iloc[:,25],
        errors="coerce"
    ).fillna(0)


    # Total Mature
    total_mature = cdr.groupby("Employee ID")[
        "matured"
    ].sum().reset_index()

    total_mature.rename(
        columns={"matured":"Total Mature"},
        inplace=True
    )


    # IB Mature
    ib = cdr[
        cdr["Campaign"]=="CSRINBOUND"
    ].groupby("Employee ID")[
        "matured"
    ].sum().reset_index()

    ib.rename(
        columns={"matured":"IB Mature"},
        inplace=True
    )



    # =====================================================
    # MERGE DATA
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


    # =====================================================
    # AHT CALCULATION
    # =====================================================

    final["AHT_sec"] = (
        final["talk_sec"]
        /
        final["Total Mature"].replace(0,1)
    )



    # =====================================================
    # CONVERT BACK TO TIME FORMAT
    # =====================================================

    final["Total Login Time"] = final["login_sec"].apply(seconds_to_time)

    final["Total Break"] = final["break_sec"].apply(seconds_to_time)

    final["Total Net Login"] = final["net_sec"].apply(seconds_to_time)

    final["Total Meeting"] = final["meeting_sec"].apply(seconds_to_time)

    final["Total Talk Time"] = final["talk_sec"].apply(seconds_to_time)

    final["AHT"] = final["AHT_sec"].apply(seconds_to_time)



    # =====================================================
    # FINAL OUTPUT
    # =====================================================

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


# ---------- RUN ----------

if __name__ == "__main__":
    app.run()
