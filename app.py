from flask import Flask, render_template, request
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import tempfile

app = Flask(__name__)

# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        t = str(t)
        h, m, s = map(int, t.split(":"))
        return h*3600 + m*60 + s
    except:
        return 0


def seconds_to_time(sec):

    sec = int(sec)

    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60

    return f"{h:02}:{m:02}:{s:02}"


# =========================
# UNMERGE FUNCTION
# =========================

def unmerge_excel(file):

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp.close()

    wb = load_workbook(file)
    ws = wb.active

    merged_ranges = list(ws.merged_cells.ranges)

    for merged in merged_ranges:

        value = ws.cell(merged.min_row, merged.min_col).value

        ws.unmerge_cells(str(merged))

        for row in ws.iter_rows(
            min_row=merged.min_row,
            max_row=merged.max_row,
            min_col=merged.min_col,
            max_col=merged.max_col
        ):
            for cell in row:
                cell.value = value

    wb.save(temp.name)

    return temp.name


# =========================
# HOME
# =========================

@app.route("/")
def index():
    return render_template("index.html")


# =========================
# GENERATE REPORT
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    agent_file = request.files.get("agent_file")
    cdr_file = request.files.get("cdr_file")

    if not agent_file or not cdr_file:
        return "Upload both files"

    # =========================
    # LOAD AGENT REPORT
    # =========================

    clean_agent = unmerge_excel(agent_file)

    agent = pd.read_excel(
        clean_agent,
        skiprows=2
    )

    agent.columns = agent.columns.str.strip()

    agent["Employee ID"] = agent.iloc[:,1].astype(str).str.strip()
    agent["Agent Full Name"] = agent.iloc[:,1]

    agent["Total Login Time"] = agent.iloc[:,3]
    agent["Total Net Login"] = agent.iloc[:,4]
    agent["Total Break"] = agent.iloc[:,5]

    meeting_u = pd.to_timedelta(agent.iloc[:,20], errors="coerce").fillna(pd.Timedelta(0))
    meeting_x = pd.to_timedelta(agent.iloc[:,23], errors="coerce").fillna(pd.Timedelta(0))

    agent["Total Meeting"] = meeting_u + meeting_x

    agent["Total Talk Time"] = agent.iloc[:,5]


    # =========================
    # LOAD CDR REPORT
    # =========================

    clean_cdr = unmerge_excel(cdr_file)

    cdr = pd.read_excel(
        clean_cdr,
        skiprows=1
    )

    cdr.columns = cdr.columns.str.strip()

    cdr["Employee ID"] = cdr.iloc[:,1].astype(str).str.strip()

    cdr["CallMatured"] = pd.to_numeric(
        cdr.iloc[:,25],
        errors="coerce"
    ).fillna(0)

    cdr["Transfer"] = pd.to_numeric(
        cdr.iloc[:,26],
        errors="coerce"
    ).fillna(0)

    cdr["Campaign"] = cdr.iloc[:,6].astype(str).str.upper()

    cdr["is_mature"] = (
        (cdr["CallMatured"]==1) |
        (cdr["Transfer"]==1)
    ).astype(int)


    # =========================
    # CALCULATE COUNTS
    # =========================

    total_mature = (
        cdr.groupby("Employee ID")["is_mature"]
        .sum()
        .reset_index()
    )

    total_mature.rename(
        columns={"is_mature":"Total Mature"},
        inplace=True
    )


    ib_mature = (
        cdr[cdr["Campaign"]=="CSRINBOUND"]
        .groupby("Employee ID")["is_mature"]
        .sum()
        .reset_index()
    )

    ib_mature.rename(
        columns={"is_mature":"IB Mature"},
        inplace=True
    )


    cdr_summary = total_mature.merge(
        ib_mature,
        on="Employee ID",
        how="left"
    )

    cdr_summary["IB Mature"] = (
        cdr_summary["IB Mature"]
        .fillna(0)
        .astype(int)
    )

    cdr_summary["OB Mature"] = (
        cdr_summary["Total Mature"] -
        cdr_summary["IB Mature"]
    )


    # =========================
    # MERGE BOTH REPORTS
    # =========================

    final_df = agent.merge(
        cdr_summary,
        on="Employee ID",
        how="left"
    )

    final_df[["Total Mature","IB Mature","OB Mature"]] = (
        final_df[["Total Mature","IB Mature","OB Mature"]]
        .fillna(0)
        .astype(int)
    )


    # =========================
    # AHT
    # =========================

    def calc_aht(row):

        talk = time_to_seconds(row["Total Talk Time"])
        mature = row["Total Mature"]

        if mature == 0:
            return "00:00:00"

        return seconds_to_time(talk/mature)


    final_df["AHT"] = final_df.apply(calc_aht, axis=1)


    # =========================
    # FINAL COLUMNS
    # =========================

    final_df = final_df[[
        "Employee ID",
        "Agent Full Name",
        "Total Login Time",
        "Total Net Login",
        "Total Break",
        "Total Meeting",
        "Total Mature",
        "IB Mature",
        "OB Mature",
        "Total Talk Time",
        "AHT"
    ]]


    report_time = datetime.now().strftime("%d %b %Y %I:%M %p")


    return render_template(
        "result.html",
        rows=final_df.to_dict("records"),
        report_time=report_time
    )


# =========================
# RUN LOCAL
# =========================

if __name__ == "__main__":
    app.run(debug=True)
