from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ---------------- TIME FUNCTIONS ----------------

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


# ---------------- HOME PAGE ----------------

@app.route("/")
def index():

    return render_template("index.html")


# ---------------- GENERATE REPORT ----------------

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

    agent.columns = agent.columns.str.strip()

    # Column B = Employee ID
    agent["Employee ID"] = agent.iloc[:,1].astype(str)

    # Column C = Agent Full Name
    agent["Agent Full Name"] = agent.iloc[:,2].astype(str)

    # Login Time
    agent["login_sec"] = agent["Total Login Time"].apply(time_to_seconds)

    # Break Time
    if "Total Break" in agent.columns:

        agent["break_sec"] = agent["Total Break"].apply(time_to_seconds)

    else:

        agent["break_sec"] = (
            agent.get("SHORTBREAK","0:00:00").apply(time_to_seconds) +
            agent.get("TEABREAK","0:00:00").apply(time_to_seconds) +
            agent.get("LUNCHBREAK","0:00:00").apply(time_to_seconds)
        )

    # Net Login
    if "Total Net Login" in agent.columns:

        agent["net_sec"] = agent["Total Net Login"].apply(time_to_seconds)

    else:

        agent["net_sec"] = agent["login_sec"] - agent["break_sec"]

    # Meeting
    if "Total Meeting" in agent.columns:

        agent["meeting_sec"] = agent["Total Meeting"].apply(time_to_seconds)

    else:

        agent["meeting_sec"] = agent.get(
            "SYSTEMDOWN","0:00:00"
        ).apply(time_to_seconds)

    # Talk Time
    agent["talk_sec"] = agent["Total Talk Time"].apply(time_to_seconds)



    # =====================================================
    # CDR REPORT
    # =====================================================

    cdr = pd.read_excel(cdr_path, header=2)

    cdr.columns = cdr.columns.str.strip()

    # Column B = Employee ID
    cdr["Employee ID"] = cdr.iloc[:,1].astype(str)

    # Column G = Campaign
    cdr["Campaign"] = cdr.iloc[:,6].astype(str)

    # Column Q = CallMatured
    cdr["CallMatured"] = pd.to_numeric(
        cdr.iloc[:,16], errors="coerce"
    ).fillna(0)

    # Column R = Transfer
    cdr["Transfer"] = pd.to_numeric(
        cdr.iloc[:,17], errors="coerce"
    ).fillna(0)

    # Mature calculation
    cdr["matured"] = cdr["CallMatured"] + cdr["Transfer"]


    # Total Mature
    total_mature = cdr.groupby("Employee ID")[
        "matured"
    ].sum().reset_index()

    total_mature.rename(
        columns={"matured":"Total Mature"},
        inplace=True
    )


    # IB Mature (Campaign = CSRINBOUND)
    ib = cdr[
        cdr["Campaign"]=="CSRINBOUND"
    ].groupby("Employee ID")[
        "matured"
    ].sum().reset_index()

    ib.rename(
        columns={"matured":"IB Mature"},
        inplace=True
    )


    # IVR Hit
    ivr = cdr[
        cdr["Campaign"]=="CSRINBOUND"
    ].groupby("Employee ID").size().reset_index(
        name="Total IVR Hit"
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

    final = final.merge(
        ivr,
        on="Employee ID",
        how="left"
    )

    final.fillna(0, inplace=True)


    # OB Mature
    final["OB Mature"] = (
        final["Total Mature"]
        - final["IB Mature"]
    )


    # =====================================================
    # AHT CALCULATION
    # =====================================================

    final["AHT_sec"] = (
        final["talk_sec"]
        / final["Total Mature"].replace(0,1)
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
        "Total IVR Hit",
        "Total Talk Time",
        "AHT"
    ]]


    return render_template(
        "result.html",
        table=final.to_html(index=False)
    )



# ---------------- RUN ----------------

if __name__ == "__main__":

    app.run()
