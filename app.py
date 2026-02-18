from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# =========================
# TIME FORMAT FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        h, m, s = str(t).split(":")
        return int(h)*3600 + int(m)*60 + int(s)
    except:
        return 0


def seconds_to_time(sec):
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)
    return f"{h:02}:{m:02}:{s:02}"


# =========================
# MAIN PAGE
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

    agent_path = os.path.join(UPLOAD_FOLDER, agent_file.filename)
    cdr_path = os.path.join(UPLOAD_FOLDER, cdr_file.filename)

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)

    # =========================
    # READ AGENT FILE
    # =========================

    agent = pd.read_excel(agent_path, skiprows=2)

    agent.columns = agent.columns.str.strip()

    # Column Mapping based on your file
    agent["Employee ID"] = agent.iloc[:, 0].astype(str).str.strip()
    agent["Agent Name"] = agent.iloc[:, 2].astype(str).str.strip()

    login = agent.iloc[:, 3]
    lunch = agent.iloc[:, 19]
    tea = agent.iloc[:, 22]
    short = agent.iloc[:, 24]
    meeting = agent.iloc[:, 20]
    system = agent.iloc[:, 23]

    # =========================
    # CALCULATIONS
    # =========================

    total_break_sec = (
        lunch.apply(time_to_seconds) +
        tea.apply(time_to_seconds) +
        short.apply(time_to_seconds)
    )

    meeting_sec = (
        meeting.apply(time_to_seconds) +
        system.apply(time_to_seconds)
    )

    login_sec = login.apply(time_to_seconds)

    net_login_sec = login_sec - total_break_sec

    agent["Total Login"] = login
    agent["Total Break"] = total_break_sec.apply(seconds_to_time)
    agent["Total Meeting"] = meeting_sec.apply(seconds_to_time)
    agent["Total Net Login"] = net_login_sec.apply(seconds_to_time)


    # =========================
    # READ CDR FILE
    # =========================

    cdr = pd.read_excel(cdr_path)
    cdr.columns = cdr.columns.str.strip()

    cdr["Employee ID"] = cdr.iloc[:, 1].astype(str).str.strip()

    ib = cdr[cdr.iloc[:, 5] == "IB"]
    ob = cdr[cdr.iloc[:, 5] == "OB"]

    ib_count = ib.groupby("Employee ID").size()
    ob_count = ob.groupby("Employee ID").size()

    talk_time = cdr.groupby("Employee ID")[cdr.columns[7]].apply(
        lambda x: sum(time_to_seconds(i) for i in x)
    )


    # =========================
    # FINAL MERGE
    # =========================

    final = pd.DataFrame()

    final["Employee ID"] = agent["Employee ID"]
    final["Agent Name"] = agent["Agent Name"]
    final["Total Login"] = agent["Total Login"]
    final["Total Net Login"] = agent["Total Net Login"]
    final["Total Break"] = agent["Total Break"]
    final["Total Meeting"] = agent["Total Meeting"]

    final["IB Mature"] = final["Employee ID"].map(ib_count).fillna(0).astype(int)
    final["OB Mature"] = final["Employee ID"].map(ob_count).fillna(0).astype(int)

    final["Total Mature"] = final["IB Mature"] + final["OB Mature"]

    talk_sec = final["Employee ID"].map(talk_time).fillna(0)

    final["AHT"] = (
        talk_sec / final["Total Mature"].replace(0, 1)
    ).apply(seconds_to_time)


    # =========================
    # RESULT PAGE
    # =========================

    data = final.to_dict(orient="records")

    return render_template("result.html", data=data)


# =========================
# RUN APP
# =========================

if __name__ == "__main__":
    app.run()
