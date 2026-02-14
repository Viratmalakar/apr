from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        file = request.files["file"]

        if file.filename == "":
            return "No file selected"

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        df = pd.read_excel(filepath)

        df.fillna(0, inplace=True)

        # Example calculation
        if "Total Login" in df.columns and "Break" in df.columns:
            df["Net Login"] = df["Total Login"] - df["Break"]

        table = df.to_html(classes="table", index=False)

        return render_template("result.html", table=table)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
