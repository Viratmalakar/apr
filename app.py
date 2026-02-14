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
        h, m, s = map(int, str(t).split(":"))
        return h*3600 + m*60 + s
    except:
        return 0


def seconds_to_time(sec):
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)
    return f"{h:02}:{m:02}:{s:02}"


# =========================
# SAFE HEADER DETECT
# =========================

def detect_header(df):

    max_rows = min(10, len(df))

    for i in range(max_rows):

        row = df.iloc[i].tolist()

        safe_row = []

        for x in row:

            if pd.isna(x):
                safe_row.append("")
            else:
                safe_row.append(str(x).lower())

        if any(
            ("agent" in x)
            or ("employee" in x)
            or ("username" in x)
            for x in safe_row
        ):

            df.columns = df.iloc[i]

            df = df[i+1:]

            return df.reset_index(drop=True)

    return df.reset_index(drop=True)


# =========================
# HOME
# =========================

@app.route("/")
def index():
    return render_template("index.html")


# =========================
# GENERATE REPORT
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files.get("agent_file")
    cdr_file = request.files.get("cdr_file")

    if not agent_file or not cdr_file:
        return "Upload both files"


    # =========================
    # LOAD AGENT FILE
    # =========================

    agent = pd.read_excel(agent_file, header=None)
    agent = detect_header(agent)

    agent.columns = agent.columns.astype(str).str.strip()

    agent["Employee ID"] = agent.iloc[:,1].astype(str).str.strip()
    agent["Agent Name"] = agent.iloc[:,2]

    agent["Total Login Time"] = agent.iloc[:,3]
    agent["Total Net Login"] = agent.iloc[:,4]
    agent["Total Break"] = agent.iloc[:,5]

    # Meeting = U + X
    agent["Total Meeting"] = agent.iloc[:,20].apply(time_to_seconds) + agent.iloc[:,23].apply(time_to_seconds)
    agent["Total Meeting"] = agent["Total Meeting"].apply(seconds_to_time)

    agent["Total Talk Time"] = agent.iloc[:,5]


    # =========================
    # LOAD CDR FILE
    # =========================

    cdr = pd.read_excel(cdr_file, header=None)
    cdr = detect_header(cdr)

    cdr.columns = cdr.columns.astype(str).str.strip()

    cdr["Employee ID"] = cdr.iloc[:,1].astype(str).str.strip()

    # Campaign column G
    cdr["Campaign"] = cdr.iloc[:,6].astype(str).str.upper().str.strip()

    # Mature column Z
    mature_col = cdr.iloc[:,25].astype(str).str.upper()

    cdr["MatureFlag"] = mature_col.isin(["CALLMATURED", "TRANSFER"])


    # =========================
    # COUNT LOGIC
    # =========================

    total_mature = (
        cdr[cdr["MatureFlag"]]
        .groupby("Employee ID")
        .size()
        .reset_index(name="Total Mature")
    )

    ib_mature = (
        cdr[
            (cdr["MatureFlag"])
            &
            (cdr["Campaign"]=="CSRINBOUND")
        ]
        .groupby("Employee ID")
        .size()
        .reset_index(name="IB Mature")
    )


    # =========================
    # MERGE
    # =========================

    final = agent.merge(total_mature, on="Employee ID", how="left")
    final = final.merge(ib_mature, on="Employee ID", how="left")

    final = final.fillna(0)

    final["Total Mature"] = final["Total Mature"].astype(int)
    final["IB Mature"] = final["IB Mature"].astype(int)

    final["OB Mature"] = final["Total Mature"] - final["IB Mature"]


    # =========================
    # AHT
    # =========================

    final["TalkSec"] = final["Total Talk Time"].apply(time_to_seconds)

    final["AHTsec"] = final.apply(
        lambda x: x["TalkSec"]/x["Total Mature"]
        if x["Total Mature"]>0 else 0,
        axis=1
    )

    final["AHT"] = final["AHTsec"].apply(seconds_to_time)


    # =========================
    # FINAL SELECT
    # =========================

    final = final[[
        "Employee ID",
        "Agent Name",
        "Total Login Time",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "Total Mature",
        "IB Mature",
        "OB Mature",
        "AHT"
    ]]


    final = final.sort_values(
        by="Total Net Login",
        key=lambda x: x.apply(time_to_seconds),
        ascending=False
    )


    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")


    return render_template(
        "result.html",
        rows=final.to_dict("records"),
        report_time=report_time
    )


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run(debug=True)
