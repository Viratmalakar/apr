from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(t):

    try:
        if pd.isna(t):
            return 0

        parts = str(t).split(":")

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
# SAFE HEADER DETECT
# =========================

def detect_header(df):

    for i in range(min(10, len(df))):

        row = df.iloc[i].tolist()

        row = [str(x).lower() if pd.notna(x) else "" for x in row]

        if any(
            "agent" in x
            or "employee" in x
            or "username" in x
            for x in row
        ):

            df.columns = df.iloc[i]

            df = df[i+1:]

            return df.reset_index(drop=True)

    return df


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

        return redirect("/")


    # =====================
    # LOAD AGENT FILE
    # =====================

    agent = pd.read_excel(agent_file, header=None)

    agent = detect_header(agent)

    agent.columns = agent.columns.str.strip()


    agent["Employee ID"] = agent.iloc[:,1].astype(str).str.strip()

    agent["Agent Name"] = agent.iloc[:,2]

    agent["Total Login Time"] = agent.iloc[:,3]

    agent["Total Break"] = agent.iloc[:,20]

    agent["Total Meeting"] = (
        pd.to_numeric(agent.iloc[:,20], errors="coerce").fillna(0)
        +
        pd.to_numeric(agent.iloc[:,23], errors="coerce").fillna(0)
    )

    agent["Total Talk Time"] = agent.iloc[:,5]


    # =====================
    # LOAD CDR FILE
    # =====================

    cdr = pd.read_excel(cdr_file, header=None)

    cdr = detect_header(cdr)

    cdr.columns = cdr.columns.str.strip()


    # Column B = Employee ID
    cdr["Employee ID"] = cdr.iloc[:,1].astype(str).str.strip()

    # Column G = Campaign
    cdr["Campaign"] = cdr.iloc[:,6].astype(str).str.upper().str.strip()

    # Column Z = CallMatured + Transfer result
    matured_col = pd.to_numeric(cdr.iloc[:,25], errors="coerce").fillna(0)


    # =====================
    # TOTAL MATURE COUNT
    # =====================

    cdr["is_mature"] = matured_col > 0


    total_mature = (

        cdr[cdr["is_mature"]]

        .groupby("Employee ID")

        .size()

        .reset_index(name="Total Mature")

    )


    ib_mature = (

        cdr[
            (cdr["is_mature"])
            &
            (cdr["Campaign"] == "CSRINBOUND")
        ]

        .groupby("Employee ID")

        .size()

        .reset_index(name="IB Mature")

    )


    calls = total_mature.merge(

        ib_mature,

        on="Employee ID",

        how="left"

    )

    calls = calls.fillna(0)


    calls["OB Mature"] = (

        calls["Total Mature"]

        -

        calls["IB Mature"]

    )


    # convert to int
    calls["Total Mature"] = calls["Total Mature"].astype(int)
    calls["IB Mature"] = calls["IB Mature"].astype(int)
    calls["OB Mature"] = calls["OB Mature"].astype(int)


    # =====================
    # MERGE AGENT + CALLS
    # =====================

    final = agent.merge(

        calls,

        on="Employee ID",

        how="left"

    )

    final = final.fillna(0)


    # =====================
    # AHT CALCULATION
    # =====================

    final["Talk_sec"] = final["Total Talk Time"].apply(time_to_seconds)

    final["AHT_sec"] = final.apply(

        lambda x:

        x["Talk_sec"] // x["Total Mature"]

        if x["Total Mature"] > 0 else 0,

        axis=1

    )

    final["AHT"] = final["AHT_sec"].apply(seconds_to_time)


    # =====================
    # FINAL OUTPUT
    # =====================

    final_df = pd.DataFrame({

        "Employee ID": final["Employee ID"],

        "Agent Name": final["Agent Name"],

        "Total Login Time": final["Total Login Time"],

        "Total Break": final["Total Break"],

        "Total Meeting": final["Total Meeting"],

        "Total Mature": final["Total Mature"].astype(int),

        "IB Mature": final["IB Mature"].astype(int),

        "OB Mature": final["OB Mature"].astype(int),

        "AHT": final["AHT"]

    })


    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")


    return render_template(

        "result.html",

        rows=final_df.to_dict("records"),

        report_time=report_time

    )


# =========================
# EXPORT
# =========================

@app.route("/export", methods=["POST"])
def export():

    data = request.json

    df = pd.DataFrame(data)

    output = io.BytesIO()

    df.to_excel(output, index=False)

    output.seek(0)

    return send_file(

        output,

        download_name="Agent_Report.xlsx",

        as_attachment=True

    )


# =========================
# RUN
# =========================

if __name__ == "__main__":

    app.run(debug=True)
