from flask import Flask, render_template, request
import pandas as pd
import os
import tempfile

app = Flask(__name__)

UPLOAD_FOLDER = tempfile.gettempdir()


# ===============================
# TIME FORMAT SAFE
# ===============================
def format_time(val):
    try:
        if pd.isna(val):
            return "00:00:00"

        if isinstance(val, str):
            if ":" in val:
                return val
            val = float(val)

        seconds = int(val)

        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60

        return f"{h:02}:{m:02}:{s:02}"

    except:
        return "00:00:00"


# ===============================
# MAIN PAGE
# ===============================
@app.route("/")
def index():
    return render_template("index.html")


# ===============================
# GENERATE DASHBOARD
# ===============================
@app.route("/generate", methods=["POST"])
def generate():

    try:

        ap_file = request.files["agent_performance"]
        cdr_file = request.files["cdr"]

        ap_path = os.path.join(UPLOAD_FOLDER, ap_file.filename)
        cdr_path = os.path.join(UPLOAD_FOLDER, cdr_file.filename)

        ap_file.save(ap_path)
        cdr_file.save(cdr_path)

        # ===============================
        # READ AGENT PERFORMANCE
        # Header row = row 3
        # ===============================

        ap = pd.read_excel(ap_path, header=2)

        ap = ap.replace("-", 0)

        # Employee ID
        ap["Employee ID"] = ap.iloc[:, 1].astype(str).str.strip()

        # Agent Name
        ap["Agent Name"] = ap.iloc[:, 2]

        # Time columns
        login = pd.to_timedelta(ap.iloc[:, 3].astype(str), errors="coerce").dt.total_seconds().fillna(0)

        lunch = pd.to_timedelta(ap.iloc[:, 19].astype(str), errors="coerce").dt.total_seconds().fillna(0)
        short = pd.to_timedelta(ap.iloc[:, 22].astype(str), errors="coerce").dt.total_seconds().fillna(0)
        tea = pd.to_timedelta(ap.iloc[:, 24].astype(str), errors="coerce").dt.total_seconds().fillna(0)

        meeting = pd.to_timedelta(ap.iloc[:, 20].astype(str), errors="coerce").dt.total_seconds().fillna(0)
        systemdown = pd.to_timedelta(ap.iloc[:, 23].astype(str), errors="coerce").dt.total_seconds().fillna(0)

        talk = pd.to_timedelta(ap.iloc[:, 5].astype(str), errors="coerce").dt.total_seconds().fillna(0)

        ap["Login"] = login
        ap["Break"] = lunch + short + tea
        ap["Meeting"] = meeting + systemdown
        ap["Net Login"] = login - ap["Break"]
        ap["Talk"] = talk

        ap_summary = ap[
            [
                "Employee ID",
                "Agent Name",
                "Login",
                "Net Login",
                "Break",
                "Meeting",
                "Talk",
            ]
        ]

        # ===============================
        # READ CDR
        # Header row = row 2
        # ===============================

        cdr = pd.read_excel(cdr_path, header=1)

        cdr["Employee ID"] = cdr.iloc[:, 1].astype(str).str.strip()

        status = cdr.iloc[:, 25].astype(str).str.upper()
        calltype = cdr.iloc[:, 6].astype(str).str.upper()

        cdr["Total Mature"] = status.isin(["CALLMATURED", "TRANSFER"]).astype(int)

        cdr["IB Mature"] = (
            (status.isin(["CALLMATURED", "TRANSFER"]))
            & (calltype == "CSRINBOUND")
        ).astype(int)

        mature_summary = cdr.groupby("Employee ID").agg(
            {
                "Total Mature": "sum",
                "IB Mature": "sum",
            }
        ).reset_index()

        mature_summary["OB Mature"] = (
            mature_summary["Total Mature"]
            - mature_summary["IB Mature"]
        )

        # ===============================
        # MERGE
        # ===============================

        final = pd.merge(
            ap_summary,
            mature_summary,
            on="Employee ID",
            how="left",
        )

        final = final.fillna(0)

        # ===============================
        # AHT
        # ===============================

        final["AHT"] = final.apply(
            lambda x:
            x["Talk"] / x["Total Mature"]
            if x["Total Mature"] > 0 else 0,
            axis=1
        )

        # ===============================
        # FORMAT TIME
        # ===============================

        final["Login"] = final["Login"].apply(format_time)
        final["Net Login"] = final["Net Login"].apply(format_time)
        final["Break"] = final["Break"].apply(format_time)
        final["Meeting"] = final["Meeting"].apply(format_time)
        final["AHT"] = final["AHT"].apply(format_time)

        # ===============================
        # REMOVE JUNK HEADER ROW
        # ===============================

        final = final[
            final["Employee ID"].str.lower() != "agent name"
        ]

        # ===============================
        # RENAME HEADERS
        # ===============================

        final = final.rename(
            columns={
                "Employee ID": "ID",
                "Agent Name": "Name",
                "Login": "Login",
                "Net Login": "Net Login",
                "Break": "Break",
                "Meeting": "Meeting",
                "AHT": "AHT",
                "Total Mature": "Total Mature",
                "IB Mature": "IB Mature",
                "OB Mature": "OB Mature",
            }
        )

        data = final.to_dict(orient="records")

        return render_template("result.html", data=data)

    except Exception as e:

        return str(e)


# ===============================
# RUN
# ===============================

if __name__ == "__main__":
    app.run(debug=True)
