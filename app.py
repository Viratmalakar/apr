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

    # FIX: correct header row
    agent = pd.read_excel(agent_path, header=2)

    # strip spaces from column names
    agent.columns = agent.columns.str.strip()

    # create required fields safely
    agent["Total Break"] = (
        agent.get("SHORTBREAK", 0)
        + agent.get("TEABREAK", 0)
        + agent.get("SYSTEMDOWN", 0)
    )

    agent["Total Meeting"] = agent.get("SYSTEMDOWN", 0)

    agent["Total Net Login"] = (
        agent.get("Total Login Time", 0)
        - agent["Total Break"]
    )

    agent["AHT"] = (
        agent.get("Total Talk Time", 0)
        / agent.get("No Of Call", 1)
    )

    agent["Total Call"] = agent.get("No Of Call", 0)

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

    # FINAL REQUIRED HEADERS ONLY
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
