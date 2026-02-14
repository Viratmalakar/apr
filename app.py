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

        # Excel read
        df = pd.read_excel(filepath)

        # Replace '-' with 0
        df.replace("-", 0, inplace=True)

        # Fill blank with 0
        df.fillna(0, inplace=True)

        # Example column calculations (adjust according to your file)
        try:
            df["Total Break"] = (
                df["Lunch Break"] +
                df["Tea Break"] +
                df["Short Break"]
            )

            df["Total Meeting"] = (
                df["Meeting"] +
                df["System Down"]
            )

            df["Net Login"] = (
                df["Total Login"] -
                df["Total Break"]
            )

        except:
            pass

        table = df.to_html(classes="table table-bordered", index=False)

        return render_template("result.html", table=table)

    return render_template("index.html")


if __name__ == "__main__":
    app.run()
