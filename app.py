from flask import Flask, render_template, request, send_file
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

# TIME TO SECONDS
def time_to_seconds(t):
    try:
        h, m, s = str(t).split(":")
        return int(h)*3600 + int(m)*60 + int(s)
    except:
        return 0

# SECONDS TO TIME
def seconds_to_time(sec):
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h:02}:{m:02}:{s:02}"

@app.route("/")
def index():
    return render_template("upload.html")

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    agent = pd.read_excel(agent_file)
    cdr = pd.read_excel(cdr_file)

    agent.columns = agent.columns.str.strip()
    cdr.columns = cdr.columns.str.strip()

    # REMOVE HEADER DUPLICATE ROWS
    agent = agent[agent["Employee ID"] != "Employee ID"]

    # TOTAL BREAK
    agent["Total Break"] = (
        agent["Lunch Break"].apply(time_to_seconds)
        + agent["Tea Break"].apply(time_to_seconds)
        + agent["Short Break"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # TOTAL MEETING
    agent["Total Meeting"] = (
        agent["Meeting"].apply(time_to_seconds)
        + agent["System Down"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # NET LOGIN
    agent["Total Net Login"] = (
        agent["Total Login"].apply(time_to_seconds)
        - agent["Total Break"].apply(time_to_seconds)
    ).apply(seconds_to_time)

    # ---- CDR PROCESS ----
    cdr["Employee ID"] = cdr["Employee ID"].astype(str)

    total_mature = cdr.groupby("Employee ID").size().reset_index(name="Total Mature")

    ib = cdr[cdr["Call Direction"] == "Inbound"].groupby("Employee ID").size().reset_index(name="IB Mature")

    ob = cdr[cdr["Call Direction"] == "Outbound"].groupby("Employee ID").size().reset_index(name="OB Mature")

    talk = cdr.groupby("Employee ID")["Talk Time"].apply(
        lambda x: sum(time_to_seconds(i) for i in x)
    ).reset_index(name="Talk")

    # MERGE
    df = agent.merge(total_mature,
