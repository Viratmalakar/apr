from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

OUTPUT = "report.xlsx"


def format_time(val):
    try:
        val = pd.to_timedelta(val)
        total_seconds = int(val.total_seconds())
        h = total_seconds // 3600
        m = (total_seconds % 3600) // 60
        s = total_seconds % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


@app.route("/")
def index():
    return render_template("upload.html")


@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files.get("agent_file")
    cdr_file = request.files.get("cdr_file")

    if not agent_file or not cdr_file:
        return "Upload both files", 400

    # ==========================
    # READ AGENT PERFORMANCE
    # ==========================

    agent = pd.read_excel(agent_file, header=2)

    agent = agent.replace("-", "00:00:00")

    agent["Total Break"] = (
        pd.to_timedelta(agent["LUNCHBREAK"]) +
        pd.to_timedelta(agent["SHORTBREAK"]) +
        pd.to_timedelta(agent["TEABREAK"])
    )

    agent["Total Meeting"] = (
        pd.to_timedelta(agent["MEETING"]) +
        pd.to_timedelta(agent["SYSTEMDOWN"])
    )

    agent["Total Net Login"] = (
        pd.to_timedelta(agent["LOGIN"]) -
        agent["Total Break"]
    )

    agent["Total Break"] = agent["Total Break"].astype(str)
    agent["Total Meeting"] = agent["Total Meeting"].astype(str)
    agent["Total Net Login"] = agent["Total Net Login"].astype(str)

    agent_df = pd.DataFrame({
        "Employee ID": agent["EMPLOYEEID"],
        "Agent Name": agent["NAME"],
        "Total Login": agent["LOGIN"],
        "Total Net Login": agent["Total Net Login"],
        "Total Break": agent["Total Break"],
        "Total Meeting": agent["Total Meeting"]
    })

    # ==========================
    # READ CDR
    # ==========================

    cdr = pd.read_excel(cdr_file, header=1)

    cdr = cdr[
        (cdr["CALLSTATUS"].isin(["CALLMATURED", "TRANSFER"]))
    ]

    total_mature = cdr.groupby("EMPLOYEEID").size()

    ib = cdr[cdr["CALLTYPE"] == "CSRINBOUND"].groupby("EMPLOYEEID").size()

    ob = total_mature - ib

    talk = cdr.groupby("EMPLOYEEID")["TALKTIME"].sum()

    final = agent_df.copy()

    final["Total Mature"] = final["Employee ID"].map(total_mature).fillna(0).astype(int)
    final["IB Mature"] = final["Employee ID"].map(ib).fillna(0).astype(int)
    final["OB Mature"] = final["Employee ID"].map(ob).fillna(0).astype(int)

    final["Total Talk"] = final["Employee ID"].map(talk).fillna(pd.Timedelta(seconds=0))

    final["AHT"] = final.apply(
        lambda x:
        format_time(x["Total Talk"] / x["Total Mature"])
        if x["Total Mature"] > 0 else "00:00:00",
        axis=1
    )

    final = final.drop(columns=["Total Talk"])

    final.to_excel(OUTPUT, index=False)

    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")

    return render_template(
        "result.html",
        tables=final.values.tolist(),
        report_time=report_time
    )


@app.route("/download")
def download():
    return send_file(OUTPUT, as_attachment=True)


@app.route("/reset")
def reset():
    return render_template("upload.html")


if __name__ == "__main__":
    app.run()
