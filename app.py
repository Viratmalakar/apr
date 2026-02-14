from flask import Flask, render_template, request
import pandas as pd

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
    try:
        h = int(sec // 3600)
        m = int((sec % 3600) // 60)
        s = int(sec % 60)
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


# =========================
# AUTO HEADER DETECT
# =========================

def detect_header(df):

    for i in range(10):

        row = df.iloc[i].astype(str).str.lower().tolist()

        if any("agent" in x for x in row):

            df.columns = df.iloc[i]
            df = df[i+1:]
            break

    return df.reset_index(drop=True)


# =========================
# HOME
# =========================

@app.route("/")
def index():
    return render_template("index.html")


# =========================
# PROCESS
# =========================

@app.route("/process", methods=["POST"])
def process():

    agent_file = request.files.get("agent_file")
    cdr_file = request.files.get("cdr_file")

    if not agent_file or not cdr_file:
        return render_template("index.html")


    # =========================
    # LOAD AGENT FILE
    # =========================

    agent = pd.read_excel(agent_file, header=None)
    agent = detect_header(agent)

    agent.columns = agent.columns.str.strip()


    agent["Employee ID"] = (
        agent.iloc[:,1]
        .astype(str)
        .str.replace(".0","", regex=False)
        .str.strip()
    )

    agent["Agent Name"] = agent.iloc[:,2]
    agent["Total Login Time"] = agent.iloc[:,3]
    agent["Total Break"] = agent.iloc[:,5]
    agent["Total Talk Time"] = agent.iloc[:,6]

    agent["Total Meeting"] = (
        agent.iloc[:,20].apply(time_to_seconds)
        +
        agent.iloc[:,23].apply(time_to_seconds)
    ).apply(seconds_to_time)

    agent["Total Net Login"] = (
        agent["Total Login Time"].apply(time_to_seconds)
        -
        agent["Total Break"].apply(time_to_seconds)
    ).apply(seconds_to_time)


    # =========================
    # LOAD CDR FILE
    # =========================

    cdr = pd.read_excel(cdr_file, header=None)
    cdr = detect_header(cdr)

    cdr.columns = cdr.columns.str.strip()


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


    # =========================
    # COUNT LOGIC (NO .0)
    # =========================

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
            (cdr["Campaign"]=="CSRINBOUND")
        ]
        .groupby("Employee ID")
        .size()
        .reset_index(name="IB Mature")
    )


    # =========================
    # MERGE
    # =========================

    df = agent.merge(total_mature, on="Employee ID", how="left")

    df = df.merge(ib_mature, on="Employee ID", how="left")


    df["Total Mature"] = df["Total Mature"].fillna(0).astype(int)

    df["IB Mature"] = df["IB Mature"].fillna(0).astype(int)

    df["OB Mature"] = (
        df["Total Mature"]
        -
        df["IB Mature"]
    ).astype(int)


    # =========================
    # AHT
    # =========================

    df["AHT"] = (

        df["Total Talk Time"].apply(time_to_seconds)

        /

        df["Total Mature"].replace(0,1)

    ).apply(seconds_to_time)


    # =========================
    # FINAL
    # =========================

    final_df = df[

        [

        "Employee ID",
        "Agent Name",
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

        rows = final_df.to_dict("records")

    )


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run()
