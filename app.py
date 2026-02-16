import pandas as pd
from flask import Flask, render_template, request
import os

app = Flask(__name__)

def time_to_seconds(t):
    try:
        return pd.to_timedelta(t).total_seconds()
    except:
        return 0

def seconds_to_time(s):
    return str(pd.to_timedelta(int(s), unit='s'))

@app.route('/')
def index():
    return render_template("upload.html")

@app.route('/generate', methods=['POST'])
def generate():

    agent_file = request.files['agent_file']
    cdr_file = request.files['cdr_file']

    if not agent_file or not cdr_file:
        return "Upload both files"

    agent = pd.read_excel(agent_file, header=2)
    cdr = pd.read_excel(cdr_file, header=1)

    agent = agent.fillna("00:00:00")

    # BREAK
    agent["Total Break Seconds"] = (
        agent["Lunch Break"].apply(time_to_seconds) +
        agent["Tea Break"].apply(time_to_seconds) +
        agent["Short Break"].apply(time_to_seconds)
    )

    # MEETING
    agent["Total Meeting Seconds"] = (
        agent["Meeting"].apply(time_to_seconds) +
        agent["System Down"].apply(time_to_seconds)
    )

    # LOGIN
    agent["Login Seconds"] = agent["Total Login Time"].apply(time_to_seconds)

    # NET LOGIN
    agent["Net Login Seconds"] = (
        agent["Login Seconds"] - agent["Total Break Seconds"]
    )

    # TALK TIME
    agent["Talk Seconds"] = agent["Total Talk Time"].apply(time_to_seconds)

    # CDR FILTER
    mature = cdr[
        (cdr["Call Status"].isin(["CALLMATURED", "TRANSFER"]))
    ]

    ib = mature[mature["CallType"] == "CSRINBOUND"]
    ob = mature[mature["CallType"] != "CSRINBOUND"]

    total_mature = mature.groupby("Username").size()
    ib_mature = ib.groupby("Username").size()
    ob_mature = ob.groupby("Username").size()

    result = []

    for index, row in agent.iterrows():

        emp = row["Employee ID"]
        name = row["Agent Full Name"]

        tm = total_mature.get(emp, 0)

        talk = row["Talk Seconds"]

        aht = seconds_to_time(talk/tm) if tm > 0 else "00:00:00"

        result.append({

            "Employee ID": emp,
            "Agent Name": name,
            "Total Login": seconds_to_time(row["Login Seconds"]),
            "Total Net Login": seconds_to_time(row["Net Login Seconds"]),
            "Total Break": seconds_to_time(row["Total Break Seconds"]),
            "Total Meeting": seconds_to_time(row["Total Meeting Seconds"]),
            "AHT": aht,
            "Total Mature": tm,
            "IB Mature": ib_mature.get(emp, 0),
            "OB Mature": ob_mature.get(emp, 0)

        })

    return render_template("result.html", data=result)

if __name__ == "__main__":
    app.run()
