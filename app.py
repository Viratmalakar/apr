from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


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

    # Read Agent Performance
    agent = pd.read_excel(agent_path, header=1)

    agent = agent[[
        "Agent Name",
        "Agent Full Name",
        "Total Login Time",
        "Total Break Duration",
        "SHORTBREAK",
        "SYSTEMDOWN",
        "Total Talk Time",
        "No Of Call"
    ]]

    agent["Total Break"] = agent["Total Break Duration"]

    agent["Total Meeting"] = agent["SHORTBREAK"] + agent["SYSTEMDOWN"]

    agent["Total Net Login"] = (
        agent["Total Login Time"] - agent["Total Break"]
    )

    agent["AHT"] = (
        agent["Total Talk Time"] / agent["No Of Call"]
    )

    agent.rename(columns={
        "No Of Call": "Total Call"
    }, inplace=True)

    # Read CDR
    cdr = pd.read_excel(cdr_path, header=1)

    ib = cdr[cdr["Call Type"] == "INBOUND"].groupby(
        "User Full Name"
    ).size().reset_index(name="IB Mature")

    ob = cdr[cdr["Call Type"] == "OUTBOUND"].groupby(
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
