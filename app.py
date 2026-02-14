from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# Home page
@app.route("/")
def index():
    return render_template("index.html")


# Agent Performance upload
@app.route("/upload_agent", methods=["POST"])
def upload_agent():

    file = request.files["file"]

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    df = pd.read_excel(filepath)

    df.replace("-", 0, inplace=True)
    df.fillna(0, inplace=True)

    try:
        df["Total Break"] = (
            df["Lunch Break"] +
            df["Tea Break"] +
            df["Short Break"]
        )

        df["Net Login"] = (
            df["Total Login"] -
            df["Total Break"]
        )
    except:
        pass

    table = df.to_html(index=False)

    return render_template("result.html", table=table)


# CDR upload
@app.route("/upload_cdr", methods=["POST"])
def upload_cdr():

    file = request.files["file"]

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    df = pd.read_excel(filepath)

    df.fillna(0, inplace=True)

    # Example calculation
    try:
        summary = df.groupby("Agent Name").size().reset_index(name="Total Calls")
        table = summary.to_html(index=False)
    except:
        table = df.to_html(index=False)

    return render_template("result.html", table=table)


if __name__ == "__main__":
    app.run()
