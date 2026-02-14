from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)


# ================================
# TIME FORMAT FUNCTION
# ================================

def to_seconds(val):
    try:
        h, m, s = str(val).split(":")
        return int(h)*3600 + int(m)*60 + int(s)
    except:
        return 0


def to_hhmmss(seconds):
    try:
        seconds = int(seconds)
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


# ================================
# HOME PAGE
# ================================

@app.route("/")
def home():
    return render_template("index.html")


# ================================
# GENERATE REPORT
# ================================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    agent = pd.read_excel(agent_file, header=2)
    cdr = pd.read_excel(cdr_file, header=2)

    agent.columns = agent.columns.str.strip()
    cdr.columns = cdr.columns.str.strip()

    # ================================
    # AGENT CLEAN
    # ================================

    agent["Employee ID"] = (
        agent.iloc[:,1]
        .astype(str)
        .str.replace(".0","", regex=False)
        .str.strip()
    )

    agent["Agent Full Name"] = agent.iloc[:,2]

    agent["Total Login Time"] = agent.iloc[:,3]
    agent["Total Break"] = agent.iloc[:,5]
    agent["Total Talk Time"] = agent.iloc[:,6]

    agent["Total Meeting"] = (
        agent.iloc[:,20].apply(to_seconds)
        + agent.iloc[:,23].apply(to_seconds)
    ).apply(to_hhmmss)


    # ================================
    # CDR CLEAN
    # ================================

    cdr["Employee ID"] = (
        cdr.iloc[:,1]
        .astype(str)
        .str.replace(".0","", regex=False)
        .str.strip()
    )

    cdr["Campaign"] = (
        cdr.iloc[:,6]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    cdr["MatureFlag"] = (
        cdr.iloc[:,25]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    cdr["is_mature"] = cdr["MatureFlag"].isin(
        ["CALLMATURED", "TRANSFER", "1"]
    )


    # ================================
    # TOTAL MATURE COUNT
    # ================================

    total_mature = (
        cdr[cdr["is_mature"]]
        .groupby("Employee ID")
        .size()
        .reset_index(name="Total Mature")
    )


    # ================================
    # IB MATURE COUNT
    # ================================

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


    # ================================
    # MERGE
    # ================================

    df = agent.merge(total_mature, on="Employee ID", how="left")
    df = df.merge(ib_mature, on="Employee ID", how="left")

    df["Total Mature"] = df["Total Mature"].fillna(0)
    df["IB Mature"] = df["IB Mature"].fillna(0)

    df["OB Mature"] = (
        df["Total Mature"]
        - df["IB Mature"]
    )


    # ================================
    # NET LOGIN
    # ================================

    df["Total Net Login"] = (
        df["Total Login Time"].apply(to_seconds)
        - df["Total Break"].apply(to_seconds)
    ).apply(to_hhmmss)


    # ================================
    # AHT
    # ================================

    df["AHT"] = (
        df["Total Talk Time"].apply(to_seconds)
        / df["Total Mature"].replace(0,1)
    ).apply(to_hhmmss)


    df = df.fillna(0)


    result = df[
        [
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
        ]
    ]

    return render_template(
        "result.html",
        tables=result.to_dict(orient="records")
    )


# ================================
# RUN
# ================================

if __name__ == "__main__":
    app.run()
