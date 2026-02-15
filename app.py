from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
from datetime import datetime
import io

app = Flask(__name__)

# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        h, m, s = map(int, str(t).split(":"))
        return h*3600 + m*60 + s
    except:
        return 0


def seconds_to_time(sec):
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)
    return f"{h:02}:{m:02}:{s:02}"


# =========================
# AUTO HEADER DETECT
# =========================

def detect_header(df):
    for i in range(10):
        row = df.iloc[i].astype(str).str.lower().tolist()
        if any("agent" in x for x in row):
            df.columns = df.iloc[i]
            df = df[i+1:]
            break
    return df.reset_index(drop=True)


# =========================
# HOME PAGE
# =========================

@app.route("/")
def index():
    return render_template("index.html")


# =========================
# PROCESS FILE
# =========================

@app.route("/process", methods=["POST"])
def process():

    file = request.files["agent_file"]

    df = pd.read_excel(file, header=None)

    df = detect_header(df)

    df.columns = df.columns.str.strip()


    column_map = {}

    for col in df.columns:

        c = col.lower()

        if "employee" in c:
            column_map["Employee ID"] = col

        elif "agent name" in c or "full name" in c:
            column_map["Agent Name"] = col

        elif "login" in c and "total" in c:
            column_map["Total Login Time"] = col

        elif "net login" in c:
            column_map["Total Net Login"] = col

        elif "break" in c:
            column_map["Total Break"] = col

        elif "meeting" in c:
            column_map["Total Meeting"] = col

        elif "mature" in c and "total" in c:
            column_map["Total Mature"] = col

        elif "ib mature" in c:
            column_map["IB Mature"] = col

        elif "ob mature" in c:
            column_map["OB Mature"] = col

        elif "aht" in c:
            column_map["AHT"] = col


    final_df = pd.DataFrame()

    final_df["Employee ID"] = df[column_map["Employee ID"]].astype(str)
    final_df["Agent Name"] = df[column_map["Agent Name"]]
    final_df["Total Login Time"] = df[column_map["Total Login Time"]]
    final_df["Total Net Login"] = df[column_map["Total Net Login"]]
    final_df["Total Break"] = df[column_map["Total Break"]]
    final_df["Total Meeting"] = df[column_map["Total Meeting"]]
    final_df["Total Mature"] = df[column_map["Total Mature"]]
    final_df["IB Mature"] = df[column_map["IB Mature"]]
    final_df["OB Mature"] = df[column_map["OB Mature"]]
    final_df["AHT"] = df[column_map["AHT"]]

    final_df = final_df.fillna("00:00:00")


    # SORT TOP PERFORMER FIRST
    final_df["sort"] = final_df["Total Net Login"].apply(time_to_seconds)

    final_df = final_df.sort_values("sort", ascending=False)

    final_df = final_df.drop("sort", axis=1)


    # SUMMARY

    summary = {

        "login": seconds_to_time(
            final_df["Total Login Time"].apply(time_to_seconds).sum()
        ),

        "mature": int(
            pd.to_numeric(final_df["Total Mature"], errors="coerce").sum()
        ),

        "ib": int(
            pd.to_numeric(final_df["IB Mature"], errors="coerce").sum()
        ),

        "ob": int(
            pd.to_numeric(final_df["OB Mature"], errors="coerce").sum()
        ),

        "aht": seconds_to_time(
            final_df["AHT"].apply(time_to_seconds).mean()
        )

    }


    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")


    return render_template(

        "result.html",
        rows=final_df.to_dict("records"),
        summary=summary,
        report_time=report_time

    )


# =========================
# EXPORT
# =========================

@app.route("/export", methods=["POST"])
def export():

    data = request.json

    df = pd.DataFrame(data)

    output = io.BytesIO()

    df.to_excel(output, index=False)

    output.seek(0)

    return send_file(

        output,
        download_name="Agent_Report.xlsx",
        as_attachment=True

    )


if __name__ == "__main__":
    app.run(debug=True)
