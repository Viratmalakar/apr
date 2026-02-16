import os
import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

# ================================
# SAFE TIME CONVERT FUNCTION
# ================================
def safe_time_to_seconds(val):
    try:
        if pd.isna(val):
            return 0

        val = str(val).strip()

        if val in ["", "-", "nan", "NaT"]:
            return 0

        parts = val.split(":")
        if len(parts) != 3:
            return 0

        h, m, s = parts
        return int(h)*3600 + int(m)*60 + int(s)

    except:
        return 0


def seconds_to_time(seconds):
    try:
        seconds = int(seconds)
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


# ================================
# REMOVE MERGED CELLS SAFE METHOD
# ================================
def clean_excel(file, skip_rows):
    df = pd.read_excel(file, header=None, skiprows=skip_rows)

    df = df.fillna("")

    return df


# ================================
# MAIN PAGE
# ================================
@app.route("/")
def index():
    return render_template("upload.html")


# ================================
# GENERATE REPORT
# ================================
@app.route("/generate", methods=["POST"])
def generate():

    try:

        agent_file = request.files["agent"]
        cdr_file = request.files["cdr"]

        # ================================
        # READ FILES (skip headers)
        # ================================
        agent_df = clean_excel(agent_file, 2)
        cdr_df = clean_excel(cdr_file, 1)

        # ================================
        # AGENT PERFORMANCE PROCESS
        # ================================

        agent = pd.DataFrame()

        agent["Employee ID"] = agent_df.iloc[:,1].astype(str).str.strip()
        agent["Agent Name"] = agent_df.iloc[:,2].astype(str).str.strip()

        login = agent_df.iloc[:,3].apply(safe_time_to_seconds)

        lunch = agent_df.iloc[:,19].apply(safe_time_to_seconds)
        short = agent_df.iloc[:,22].apply(safe_time_to_seconds)
        tea = agent_df.iloc[:,24].apply(safe_time_to_seconds)

        meeting = agent_df.iloc[:,20].apply(safe_time_to_seconds)
        system = agent_df.iloc[:,23].apply(safe_time_to_seconds)

        talk = agent_df.iloc[:,5].apply(safe_time_to_seconds)

        agent["Total Login"] = login.apply(seconds_to_time)

        break_sec = lunch + short + tea
        meeting_sec = meeting + system

        net_sec = login - break_sec

        agent["Total Net Login"] = net_sec.apply(seconds_to_time)
        agent["Total Break"] = break_sec.apply(seconds_to_time)
        agent["Total Meeting"] = meeting_sec.apply(seconds_to_time)

        agent["TalkSec"] = talk

        # ================================
        # CDR PROCESS
        # ================================

        cdr = pd.DataFrame()

        cdr["Employee ID"] = cdr_df.iloc[:,1].astype(str).str.strip()
        cdr["CallType"] = cdr_df.iloc[:,6].astype(str).str.upper().str.strip()
        cdr["CallStatus"] = cdr_df.iloc[:,25].astype(str).str.upper().str.strip()

        matured = cdr["CallStatus"].isin(["CALLMATURED","TRANSFER"])

        total = cdr[matured].groupby("Employee ID").size()

        ib = cdr[matured & (cdr["CallType"]=="CSRINBOUND")].groupby("Employee ID").size()

        result = agent.copy()

        result["Total Mature"] = result["Employee ID"].map(total).fillna(0)
        result["IB Mature"] = result["Employee ID"].map(ib).fillna(0)

        result["OB Mature"] = result["Total Mature"] - result["IB Mature"]

        # ================================
        # AHT CALCULATION
        # ================================
        def calc_aht(row):
            if row["Total Mature"] == 0:
                return "00:00:00"
            return seconds_to_time(row["TalkSec"] / row["Total Mature"])

        result["AHT"] = result.apply(calc_aht, axis=1)

        result = result.drop(columns=["TalkSec"])

        # ================================
        # FINAL COLUMN ORDER
        # ================================
        result = result[
            [
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
            ]
        ]

        result = result.sort_values("Employee ID")

        table = result.to_dict(orient="records")

        return render_template("result.html", data=table)

    except Exception as e:

        return str(e)


# ================================
# PORT FIX FOR RENDER
# ================================
if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)
