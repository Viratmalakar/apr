from flask import Flask, render_template, request, send_file, session, redirect, url_for
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = "final_project_secret"

OUTPUT_FILE = "Agent_Performance_Report.xlsx"


# =========================
# TIME FORMAT FUNCTION
# =========================

def format_time(seconds):
    try:
        seconds = int(seconds)
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"
    except:
        return "00:00:00"


def time_to_seconds(val):

    if pd.isna(val):
        return 0

    try:
        parts = str(val).split(":")
        if len(parts) == 3:
            return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
    except:
        pass

    return 0


# =========================
# PROCESS AGENT PERFORMANCE
# =========================

def process_agent(file):

    df = pd.read_excel(file, skiprows=2)

    df = df.replace("-", "00:00:00")

    df = df.rename(columns={
        df.columns[1]: "Employee ID",
        df.columns[2]: "Agent Name",
        df.columns[3]: "Login",
        df.columns[19]: "Lunch",
        df.columns[22]: "ShortBreak",
        df.columns[24]: "TeaBreak",
        df.columns[20]: "Meeting",
        df.columns[23]: "SystemDown",
        df.columns[5]: "TalkTime"
    })

    df["Employee ID"] = df["Employee ID"].astype(str)

    df["Login_sec"] = df["Login"].apply(time_to_seconds)
    df["Lunch_sec"] = df["Lunch"].apply(time_to_seconds)
    df["Short_sec"] = df["ShortBreak"].apply(time_to_seconds)
    df["Tea_sec"] = df["TeaBreak"].apply(time_to_seconds)
    df["Meeting_sec"] = df["Meeting"].apply(time_to_seconds)
    df["System_sec"] = df["SystemDown"].apply(time_to_seconds)
    df["Talk_sec"] = df["TalkTime"].apply(time_to_seconds)

    df["Break_sec"] = df["Lunch_sec"] + df["Short_sec"] + df["Tea_sec"]
    df["NetLogin_sec"] = df["Login_sec"] - df["Break_sec"]
    df["TotalMeeting_sec"] = df["Meeting_sec"] + df["System_sec"]

    df = df[["Employee ID","Agent Name","Login_sec","NetLogin_sec","Break_sec","TotalMeeting_sec","Talk_sec"]]

    return df


# =========================
# PROCESS CDR REPORT
# =========================

def process_cdr(file):

    df = pd.read_excel(file, skiprows=1)

    df = df.rename(columns={
        df.columns[1]: "Employee ID",
        df.columns[6]: "CallType",
        df.columns[25]: "CallStatus"
    })

    df["Employee ID"] = df["Employee ID"].astype(str)

    df["Mature"] = df["CallStatus"].isin(["CALLMATURED","TRANSFER"])

    total = df[df["Mature"]].groupby("Employee ID").size()

    inbound = df[(df["Mature"]) & (df["CallType"]=="CSRINBOUND")].groupby("Employee ID").size()

    result = pd.DataFrame({
        "Total Mature": total,
        "IB Mature": inbound
    }).fillna(0)

    result["OB Mature"] = result["Total Mature"] - result["IB Mature"]

    result.reset_index(inplace=True)

    return result


# =========================
# MERGE REPORTS
# =========================

def merge_reports(agent_df, cdr_df):

    df = pd.merge(agent_df, cdr_df, on="Employee ID", how="left")

    df = df.fillna(0)

    df["AHT_sec"] = df["Talk_sec"] / df["Total Mature"].replace(0,1)

    final = pd.DataFrame({

        "Employee ID": df["Employee ID"],
        "Agent Name": df["Agent Name"],
        "Total Login": df["Login_sec"].apply(format_time),
        "Total Net Login": df["NetLogin_sec"].apply(format_time),
        "Total Break": df["Break_sec"].apply(format_time),
        "Total Meeting": df["TotalMeeting_sec"].apply(format_time),
        "AHT": df["AHT_sec"].apply(format_time),
        "Total Mature": df["Total Mature"].astype(int),
        "IB Mature": df["IB Mature"].astype(int),
        "OB Mature": df["OB Mature"].astype(int)

    })

    return final


# =========================
# ROUTES
# =========================

@app.route("/")
def index():

    return render_template("upload.html")


@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files.get("agent_file")
    cdr_file = request.files.get("cdr_file")

    if not agent_file or not cdr_file:
        return "Upload both files", 400

    agent_df = process_agent(agent_file)
    cdr_df = process_cdr(cdr_file)

    final_df = merge_reports(agent_df, cdr_df)

    final_df.to_excel(OUTPUT_FILE, index=False)

    session["report_ready"] = True

    return render_template(
        "result.html",
        tables=[final_df.to_dict(orient="records")],
        report_time=datetime.now().strftime("%d %b %Y %I:%M %p")
    )


@app.route("/download")
def download():

    if os.path.exists(OUTPUT_FILE):
        return send_file(OUTPUT_FILE, as_attachment=True)

    return redirect("/")


@app.route("/reset")
def reset():

    session.clear()

    return redirect("/")


# =========================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
