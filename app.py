from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import os
import tempfile

app = Flask(__name__)

final_df = None


# -------------------------
# SAFE EXCEL READ FUNCTION
# -------------------------
def read_excel_safe(file, skip_rows):

    df = pd.read_excel(file, header=None)

    df = df.iloc[skip_rows:]

    df = df.reset_index(drop=True)

    return df


# -------------------------
# CLEAN TIME FUNCTION
# -------------------------
def clean_time(val):

    if pd.isna(val) or val == "-":
        return "00:00:00"

    return str(val)


# -------------------------
# CALCULATE REPORT
# -------------------------
def generate_report(agent_file, cdr_file):

    # Agent Performance
    agent = read_excel_safe(agent_file, 2)

    # CDR
    cdr = read_excel_safe(cdr_file, 1)

    # Agent columns
    agent.columns = range(agent.shape[1])
    cdr.columns = range(cdr.shape[1])

    # Replace -
    agent = agent.replace("-", 0)

    # Employee ID
    agent["Employee ID"] = agent[1].astype(str)

    # Agent Name
    agent["Agent Name"] = agent[2].astype(str)

    # Login
    agent["Total Login"] = agent[3]

    # Break
    agent["Total Break"] = (
        pd.to_timedelta(agent[19], errors="coerce").fillna(pd.Timedelta(0))
        + pd.to_timedelta(agent[22], errors="coerce").fillna(pd.Timedelta(0))
        + pd.to_timedelta(agent[24], errors="coerce").fillna(pd.Timedelta(0))
    )

    # Meeting
    agent["Total Meeting"] = (
        pd.to_timedelta(agent[20], errors="coerce").fillna(pd.Timedelta(0))
        + pd.to_timedelta(agent[23], errors="coerce").fillna(pd.Timedelta(0))
    )

    # Net Login
    agent["Total Net Login"] = (
        pd.to_timedelta(agent[3], errors="coerce").fillna(pd.Timedelta(0))
        - agent["Total Break"]
    )

    # Talk Time
    agent["Total Talk Time"] = pd.to_timedelta(agent[5], errors="coerce").fillna(pd.Timedelta(0))


    # -------------------------
    # CDR PROCESS
    # -------------------------

    cdr["Employee ID"] = cdr[1].astype(str)

    cdr["Call Status"] = cdr[25].astype(str)

    cdr["Campaign"] = cdr[6].astype(str)

    # Total Mature count
    total_mature = cdr[cdr["Call Status"].isin(["CALLMATURED", "TRANSFER"])]

    total_mature = total_mature.groupby("Employee ID").size()

    # IB Mature
    ib = cdr[
        (cdr["Call Status"].isin(["CALLMATURED", "TRANSFER"]))
        & (cdr["Campaign"] == "CSRINBOUND")
    ]

    ib = ib.groupby("Employee ID").size()

    # Merge
    result = agent[[
        "Employee ID",
        "Agent Name",
        "Total Login",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "Total Talk Time"
    ]].copy()

    result["Total Mature"] = result["Employee ID"].map(total_mature).fillna(0).astype(int)

    result["IB Mature"] = result["Employee ID"].map(ib).fillna(0).astype(int)

    result["OB Mature"] = result["Total Mature"] - result["IB Mature"]


    # AHT
    result["AHT"] = result.apply(
        lambda row:
        str(row["Total Talk Time"] / row["Total Mature"])
        if row["Total Mature"] > 0 else "00:00:00",
        axis=1
    )


    # FORMAT TIME
    for col in [
        "Total Login",
        "Total Net Login",
        "Total Break",
        "Total Meeting"
    ]:
        result[col] = result[col].astype(str)

    result["Total Talk Time"] = result["Total Talk Time"].astype(str)

    # FINAL HEADER ORDER
    result = result[[
        "Employee ID",
        "Agent Name",
        "Total Login",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "AHT",
        "Total Mature",
        "IB Mature",
        "OB Mature"
    ]]

    return result


# -------------------------
# ROUTES
# -------------------------

@app.route("/")
def index():

    return render_template("upload.html")


@app.route("/generate", methods=["POST"])
def generate():

    global final_df

    agent_file = request.files["agent_file"]
    cdr_file = request.files["cdr_file"]

    final_df = generate_report(agent_file, cdr_file)

    return render_template(
        "result.html",
        tables=final_df.to_dict(orient="records"),
        headers=final_df.columns
    )


# -------------------------
# SAFE DOWNLOAD (NO CRASH)
# -------------------------

@app.route("/download")
def download():

    global final_df

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

    final_df.to_excel(temp.name, index=False, engine="xlsxwriter")

    temp.close()

    return send_file(
        temp.name,
        as_attachment=True,
        download_name="Agent_Performance_Report.xlsx"
    )


@app.route("/reset")
def reset():

    global final_df

    final_df = None

    return redirect(url_for("index"))


# -------------------------
# RUN
# -------------------------

if __name__ == "__main__":
    app.run(debug=True)
