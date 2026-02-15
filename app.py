from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)


# ================================
# TIME FUNCTIONS
# ================================

def time_to_seconds(val):
    try:
        if pd.isna(val):
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


# ================================
# HOME PAGE
# ================================

@app.route("/")
def index():
    return render_template("index.html")


# ================================
# GENERATE DASHBOARD
# ================================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    # ================================
    # LOAD AGENT PERFORMANCE
    # ================================

    agent = pd.read_excel(agent_file, header=2)

    agent = agent.fillna(0)
    agent.replace("-", 0, inplace=True)

    # Column mapping by position
    emp_id = agent.iloc[:, 1].astype(str)
    agent_name = agent.iloc[:, 2]

    total_login = agent.iloc[:, 3]
    talk_time = agent.iloc[:, 5]

    lunch = agent.iloc[:, 19]
    meeting = agent.iloc[:, 20]
    short = agent.iloc[:, 22]
    systemdown = agent.iloc[:, 23]
    tea = agent.iloc[:, 24]

    total_break_sec = (
        lunch.apply(time_to_seconds) +
        short.apply(time_to_seconds) +
        tea.apply(time_to_seconds)
    )

    total_meeting_sec = (
        meeting.apply(time_to_seconds) +
        systemdown.apply(time_to_seconds)
    )

    net_login_sec = (
        total_login.apply(time_to_seconds) - total_break_sec
    )

    agent_df = pd.DataFrame()

    agent_df["Employee ID"] = emp_id
    agent_df["Agent Name"] = agent_name
    agent_df["Total Login"] = total_login
    agent_df["Total Net Login"] = net_login_sec.apply(seconds_to_time)
    agent_df["Total Break"] = total_break_sec.apply(seconds_to_time)
    agent_df["Total Meeting"] = total_meeting_sec.apply(seconds_to_time)
    agent_df["Talk Sec"] = talk_time.apply(time_to_seconds)

    # ================================
    # LOAD CDR REPORT
    # ================================

    cdr = pd.read_excel(cdr_file, header=1)

    cdr_emp = cdr.iloc[:, 1].astype(str)
    call_type = cdr.iloc[:, 6].astype(str)
    call_status = cdr.iloc[:, 25].astype(str)

    cdr_df = pd.DataFrame()

    cdr_df["Employee ID"] = cdr_emp
    cdr_df["CallType"] = call_type
    cdr_df["CallStatus"] = call_status

    # Total Mature
    total_mature = (
        cdr_df[
            cdr_df["CallStatus"].isin(["CALLMATURED", "TRANSFER"])
        ]
        .groupby("Employee ID")
        .size()
        .reset_index(name="Total Mature")
    )

    # IB Mature
    ib_mature = (
        cdr_df[
            (cdr_df["CallStatus"].isin(["CALLMATURED", "TRANSFER"])) &
            (cdr_df["CallType"] == "CSRINBOUND")
        ]
        .groupby("Employee ID")
        .size()
        .reset_index(name="IB Mature")
    )

    # Merge
    final = agent_df.merge(total_mature, on="Employee ID", how="left")
    final = final.merge(ib_mature, on="Employee ID", how="left")

    final["Total Mature"] = final["Total Mature"].fillna(0).astype(int)
    final["IB Mature"] = final["IB Mature"].fillna(0).astype(int)

    final["OB Mature"] = final["Total Mature"] - final["IB Mature"]

    # ================================
    # AHT
    # ================================

    final["AHT Sec"] = final.apply(
        lambda x: x["Talk Sec"] // x["Total Mature"]
        if x["Total Mature"] > 0 else 0,
        axis=1
    )

    final["AHT"] = final["AHT Sec"].apply(seconds_to_time)

    # ================================
    # FINAL FORMAT
    # ================================

    final = final[[
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

    return render_template(
        "result.html",
        rows=final.to_dict(orient="records")
    )


# ================================
# RUN
# ================================

if __name__ == "__main__":
    app.run()
