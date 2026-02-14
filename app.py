from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

agent_df = None
cdr_df = None


def clean_agent_report(path):

    df = pd.read_excel(path, header=1)

    df.columns = [
        "Report Date",
        "Agent ID",
        "Agent Name",
        "Total Login",
        "Total Calls",
        "Talk Time",
        "Wait Time",
        "ACW Time",
        "Aux Time",
        "Ringing Time",
        *df.columns[10:]
    ]

    df = df[["Agent ID", "Agent Name", "Total Login", "Total Calls", "Talk Time"]]

    return df


def clean_cdr_report(path):

    df = pd.read_excel(path, header=1)

    df.columns = [
        "Sno",
        "Agent ID",
        "Agent Name",
        "Account Code",
        "Call Date",
        "Queue",
        "Campaign",
        "Skill",
        "List Name",
        "UniqueId",
        *df.columns[10:]
    ]

    df["Connected"] = df["Call Status"].apply(
        lambda x: 1 if str(x).upper() == "CONNECTED" else 0
    )

    summary = df.groupby(
        ["Agent ID", "Agent Name"]
    ).agg(
        Total_CDR_Calls=("Agent ID", "count"),
        Connected_Calls=("Connected", "sum")
    ).reset_index()

    return summary


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload_agent", methods=["POST"])
def upload_agent():

    global agent_df

    file = request.files["file"]
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    agent_df = clean_agent_report(path)

    return render_template(
        "result.html",
        table=agent_df.to_html(index=False)
    )


@app.route("/upload_cdr", methods=["POST"])
def upload_cdr():

    global cdr_df, agent_df

    file = request.files["file"]
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    cdr_df = clean_cdr_report(path)

    if agent_df is not None:

        final = agent_df.merge(
            cdr_df,
            on=["Agent ID", "Agent Name"],
            how="left"
        )

        final.fillna(0, inplace=True)

        final["IVR HIT"] = (
            final["Connected_Calls"] / final["Total_CDR_Calls"]
        ) * 100

        return render_template(
            "result.html",
            table=final.to_html(index=False)
        )

    return render_template(
        "result.html",
        table=cdr_df.to_html(index=False)
    )


if __name__ == "__main__":
    app.run()
