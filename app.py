from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# convert HH:MM:SS to seconds
def time_to_seconds(t):

    try:
        if pd.isna(t):
            return 0

        parts = str(t).split(":")
        return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])

    except:
        return 0


# convert seconds to HH:MM:SS
def seconds_to_time(sec):

    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)

    return f"{h:02}:{m:02}:{s:02}"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    agent_path = os.path.join(UPLOAD_FOLDER, agent_file.filename)
    cdr_path = os.path.join(UPLOAD_FOLDER, cdr_file.filename)

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)

    agent = pd.read_excel(agent_path, header=2)
    agent.columns = agent.columns.str.strip()

    # convert time columns to seconds
    agent["login_sec"] = agent["Total Login Time"].apply(time_to_seconds)

    agent["break_sec"] = (
        agent.get("SHORTBREAK", "0:00:00").apply(time_to_seconds)
        + agent.get("TEABREAK", "0:00:00").apply(time_to_seconds)
        + agent.get("SYSTEMDOWN", "0:00:00").apply(time_to_seconds)
    )

    agent["meeting_sec"] = agent.get(
        "SYSTEMDOWN", "0:00:00"
    ).apply(time_to_seconds)

    agent["net_sec"] = agent["login_sec"] - agent["break_sec"]

    agent["talk_sec"] = agent.get(
        "Total Talk Time", "0:00:00"
    ).apply(time_to_seconds)

    agent["Total Call"] = agent.get("No Of Call", 0)

    agent["aht_sec"] = agent["talk_sec"] / agent["Total Call"].replace(0, 1)

    # convert back to time format
    agent["Total Login Time"] = agent["login_sec"].apply(seconds_to_time)

    agent["Total Break"] = agent["break_sec"].apply(seconds_to_time)

    agent["Total Meeting"] = agent["meeting_sec"].apply(seconds_to_time)

    agent["Total Net Login"] = agent["net_sec"].apply(seconds_to_time)

    agent["AHT"] = agent["aht_sec"].apply(seconds_to_time)

    # read CDR
    cdr = pd.read_excel(cdr_path, header=2)
    cdr.columns = cdr.columns.str.strip()

    ib = cdr[cdr.get("Call Type", "") == "INBOUND"].groupby(
        "User Full Name"
    ).size().reset_index(name="IB Mature")

    ob = cdr[cdr.get("Call Type", "") == "OUTBOUND"].groupby(
        "User Full Name"
    ).size().reset_index(name="OB Mature")

    final = agent.merge(
        ib,
        left_on="Agent Full Name",
        right_on="User Full Name",
        how="left"
    )

    final = final.merge(
        ob,
        left_on="Agent Full Name",
        right_on="User Full Name",
        how="left"
    )

    final.fillna(0, inplace=True)

    final = final[[
        "Agent Name",
        "Agent Full Name",
        "Total Login Time",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "AHT",
        "Total Call",
        "IB Mature",
        "OB Mature"
    ]]

    return render_template(
        "result.html",
        table=final.to_html(index=False)
    )


if __name__ == "__main__":
    app.run()
