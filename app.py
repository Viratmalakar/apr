from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# =========================
# TIME CONVERSION FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        h, m, s = str(t).split(":")
        return int(h)*3600 + int(m)*60 + int(s)
    except:
        return 0


def seconds_to_time(sec):
    sec = int(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


# =========================
# HOME PAGE
# =========================

@app.route("/")
def index():
    return render_template("upload.html")


# =========================
# GENERATE DASHBOARD
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    if "agent_file" not in request.files or "cdr_file" not in request.files:
        return "Please upload both files"

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    # =========================
    # READ FILES
    # =========================

    agent = pd.read_excel(agent_file)
    cdr = pd.read_excel(cdr_file)

    # =========================
    # CLEAN HEADERS
    # =========================

    agent.columns = agent.columns.str.strip()
    cdr.columns = cdr.columns.str.strip()

    # =========================
    # REQUIRED COLUMNS (FINAL LOGIC)
    # =========================

    emp_col = agent.columns[0]
    name_col = agent.columns[2]
    login_col = agent.columns[3]
    lunch_col = agent.columns[19]
    tea_col = agent.columns[22]
    short_col = agent.columns[24]
    meeting_col = agent.columns[20]
    system_col = agent.columns[23]

    # =========================
    # CALCULATIONS
    # =========================

    agent["Total Break"] = (
        agent[lunch_col].apply(time_to_seconds) +
        agent[tea_col].apply(time_to_seconds) +
        agent[short_col].apply(time_to_seconds)
    )

    agent["Total Meeting"] = (
        agent[meeting_col].apply(time_to_seconds) +
        agent[system_col].apply(time_to_seconds)
    )

    agent["Net Login"] = (
        agent[login_col].apply(time_to_seconds)
        - agent["Total Break"]
    )

    # =========================
    # CDR CALCULATION
    # =========================

    cdr.columns = cdr.columns.str.strip()

    cdr["Employee ID"] = cdr.iloc[:,0].astype(str)

    mature = cdr[cdr.iloc[:,5] == "Mature"]

    total_mature = mature.groupby("Employee ID").size()
    ib_mature = mature[mature.iloc[:,3] == "IB"].groupby("Employee ID").size()
    ob_mature = mature[mature.iloc[:,3] == "OB"].groupby("Employee ID").size()

    talk_time = mature.groupby("Employee ID")[cdr.columns[6]].apply(
        lambda x: sum(time_to_seconds(i) for i in x)
    )

    # =========================
    # BUILD FINAL DATAFRAME
    # =========================

    result = pd.DataFrame()

    result["Employee ID"] = agent[emp_col].astype(str)
    result["Agent Name"] = agent[name_col]

    result["Total Login"] = agent[login_col].apply(time_to_seconds).apply(seconds_to_time)

    result["Total Net Login"] = agent["Net Login"].apply(seconds_to_time)

    result["Total Break"] = agent["Total Break"].apply(seconds_to_time)

    result["Total Meeting"] = agent["Total Meeting"].apply(seconds_to_time)

    result["Total Mature"] = result["Employee ID"].map(total_mature).fillna(0).astype(int)

    result["IB Mature"] = result["Employee ID"].map(ib_mature).fillna(0).astype(int)

    result["OB Mature"] = result["Employee ID"].map(ob_mature).fillna(0).astype(int)

    # =========================
    # AHT CALCULATION
    # =========================

    def calc_aht(emp):
        tm = talk_time.get(emp, 0)
        count = total_mature.get(emp, 0)

        if count == 0:
            return "00:00:00"

        return seconds_to_time(tm // count)

    result["AHT"] = result["Employee ID"].apply(calc_aht)

    # =========================
    # SUMMARY
    # =========================

    summary = {

        "total_agents": len(result),

        "total_login": seconds_to_time(
            sum(time_to_seconds(i) for i in result["Total Login"])
        ),

        "total_net_login": seconds_to_time(
            sum(time_to_seconds(i) for i in result["Total Net Login"])
        ),

        "total_mature": result["Total Mature"].sum()

    }

    return render_template(
        "result.html",
        tables=result.to_dict(orient="records"),
        summary=summary
    )


# =========================
# RUN APP (RENDER READY)
# =========================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)
