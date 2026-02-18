from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import os
from datetime import datetime
import io

app = Flask(__name__)

# =========================
# TIME CONVERSION FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        parts = str(t).split(":")
        if len(parts) == 3:
            return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        return 0
    except:
        return 0


def seconds_to_time(sec):
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h:02}:{m:02}:{s:02}"


# =========================
# HOME PAGE
# =========================

@app.route("/")
def index():
    return render_template("upload.html")


# =========================
# GENERATE REPORT
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    if "agent_file" not in request.files or "cdr_file" not in request.files:
        return "Upload both files"

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    agent = pd.read_excel(agent_file)
    cdr = pd.read_excel(cdr_file)

    # Remove blank header row automatically
    agent = agent[agent.iloc[:,0] != "Agent Name"]

    # Reset index
    agent = agent.reset_index(drop=True)

    # =========================
    # AUTO DETECT REQUIRED COLUMNS
    # =========================

    emp_col = agent.columns[0]
    name_col = agent.columns[1]
    login_col = agent.columns[3]
    lunch_col = agent.columns[19]
    tea_col = agent.columns[22]
    short_col = agent.columns[24]
    meeting_col = agent.columns[20]
    sysdown_col = agent.columns[23]

    # =========================
    # CALCULATIONS
    # =========================

    agent["Total Login"] = agent[login_col]

    agent["Total Break"] = (
        agent[lunch_col].apply(time_to_seconds) +
        agent[tea_col].apply(time_to_seconds) +
        agent[short_col].apply(time_to_seconds)
    ).apply(seconds_to_time)

    agent["Total Meeting"] = (
        agent[meeting_col].apply(time_to_seconds) +
        agent[sysdown_col].apply(time_to_seconds)
    ).apply(seconds_to_time)

    agent["Total Net Login"] = (
        agent[login_col].apply(time_to_seconds) -
        agent["Total Break"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # =========================
    # CDR CALCULATIONS
    # =========================

    cdr["Employee ID"] = cdr.iloc[:,1].astype(str)

    total_mature = cdr.groupby("Employee ID").size()
    ib_mature = cdr[cdr.iloc[:,5]=="IB"].groupby("Employee ID").size()
    ob_mature = cdr[cdr.iloc[:,5]=="OB"].groupby("Employee ID").size()

    talk_time = cdr.groupby("Employee ID")[cdr.columns[8]].apply(
        lambda x: sum(time_to_seconds(i) for i in x)
    )

    # =========================
    # BUILD FINAL TABLE
    # =========================

    final = pd.DataFrame()

    final["Employee ID"] = agent[emp_col].astype(str)
    final["Agent Name"] = agent[name_col]
    final["Total Login"] = agent["Total Login"]
    final["Total Net Login"] = agent["Total Net Login"]
    final["Total Break"] = agent["Total Break"]
    final["Total Meeting"] = agent["Total Meeting"]

    final["Total Mature"] = final["Employee ID"].map(total_mature).fillna(0).astype(int)
    final["IB Mature"] = final["Employee ID"].map(ib_mature).fillna(0).astype(int)
    final["OB Mature"] = final["Employee ID"].map(ob_mature).fillna(0).astype(int)

    final["AHT"] = final.apply(
        lambda row: seconds_to_time(
            int(talk_time.get(row["Employee ID"],0) / row["Total Mature"])
        ) if row["Total Mature"]>0 else "00:00:00",
        axis=1
    )

    # =========================
    # SAVE OUTPUT
    # =========================

    output = io.BytesIO()
    final.to_excel(output, index=False)
    output.seek(0)

    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")

    return render_template(
        "result.html",
        tables=final.values,
        headers=final.columns,
        report_time=report_time
    )


# =========================
# DOWNLOAD EXCEL
# =========================

@app.route("/download", methods=["POST"])
def download():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    agent = pd.read_excel(agent_file)
    cdr = pd.read_excel(cdr_file)

    # same logic reused
    # simplified export

    output = io.BytesIO()
    agent.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Agent_Performance_Report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# RESET
# =========================

@app.route("/reset")
def reset():
    return redirect("/")


# =========================
# RENDER COMPATIBILITY
# =========================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
