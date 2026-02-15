from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(val):
    try:
        if pd.isna(val) or val == "-":
            return 0
        h, m, s = map(int, str(val).split(":"))
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

def unmerge_excel(path):

    from openpyxl import load_workbook

    wb = load_workbook(path)
    ws = wb.active

    merged = list(ws.merged_cells.ranges)

    for m in merged:
        value = ws.cell(m.min_row, m.min_col).value
        ws.unmerge_cells(str(m))

        for row in ws.iter_rows(
            min_row=m.min_row,
            max_row=m.max_row,
            min_col=m.min_col,
            max_col=m.max_col
        ):
            for cell in row:
                cell.value = value

    temp = path + "_unmerged.xlsx"
    wb.save(temp)

    return temp


# =========================
# MAIN ROUTE
# =========================

@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        agent_file = request.files["agent"]
        cdr_file = request.files["cdr"]

        agent_path = "agent.xlsx"
        cdr_path = "cdr.xlsx"

        agent_file.save(agent_path)
        cdr_file.save(cdr_path)

        # Unmerge
        agent_path = unmerge_excel(agent_path)
        cdr_path = unmerge_excel(cdr_path)

        # =========================
        # LOAD AGENT FILE
        # =========================

        agent = pd.read_excel(agent_path, header=None)

        # Ignore top 2 rows
        agent = agent.iloc[2:].reset_index(drop=True)

        # Replace "-" with 0
        agent = agent.replace("-", "00:00:00")

        # Column mapping
        agent["Employee ID"] = agent.iloc[:,1].astype(str).str.strip()
        agent["Agent Name"] = agent.iloc[:,2]

        agent["Login_sec"] = agent.iloc[:,3].apply(time_to_seconds)
        agent["Talk_sec"] = agent.iloc[:,5].apply(time_to_seconds)

        lunch = agent.iloc[:,19].apply(time_to_seconds)
        short = agent.iloc[:,22].apply(time_to_seconds)
        tea = agent.iloc[:,24].apply(time_to_seconds)

        meeting = agent.iloc[:,20].apply(time_to_seconds)
        system = agent.iloc[:,23].apply(time_to_seconds)

        agent["Break_sec"] = lunch + short + tea
        agent["Meeting_sec"] = meeting + system
        agent["NetLogin_sec"] = agent["Login_sec"] - agent["Break_sec"]

        agent_final = agent[[
            "Employee ID",
            "Agent Name",
            "Login_sec",
            "Break_sec",
            "Meeting_sec",
            "NetLogin_sec",
            "Talk_sec"
        ]]

        # =========================
        # LOAD CDR FILE
        # =========================

        cdr = pd.read_excel(cdr_path, header=None)

        # Ignore top 1 row
        cdr = cdr.iloc[1:].reset_index(drop=True)

        cdr["Employee ID"] = cdr.iloc[:,1].astype(str).str.strip()
        cdr["CallType"] = cdr.iloc[:,6].astype(str).str.upper()
        cdr["Status"] = cdr.iloc[:,25].astype(str).str.upper()

        # Mature flag
        cdr["Mature"] = cdr["Status"].isin(["CALLMATURED","TRANSFER"])

        # IB Mature flag
        cdr["IB"] = (cdr["CallType"] == "CSRINBOUND") & cdr["Mature"]

        # Count per agent
        total_mature = cdr.groupby("Employee ID")["Mature"].sum().astype(int)
        ib_mature = cdr.groupby("Employee ID")["IB"].sum().astype(int)

        cdr_final = pd.DataFrame({
            "Employee ID": total_mature.index,
            "Total Mature": total_mature.values,
            "IB Mature": ib_mature.reindex(total_mature.index, fill_value=0).values
        })

        cdr_final["OB Mature"] = (
            cdr_final["Total Mature"] - cdr_final["IB Mature"]
        ).astype(int)

        # =========================
        # MERGE BOTH
        # =========================

        final = agent_final.merge(
            cdr_final,
            on="Employee ID",
            how="left"
        ).fillna(0)

        # =========================
        # AHT Calculation
        # =========================

        final["AHT_sec"] = final.apply(
            lambda x:
            x["Talk_sec"] // x["Total Mature"]
            if x["Total Mature"] > 0 else 0,
            axis=1
        )

        # =========================
        # Convert to display
        # =========================

        final["Total Login Time"] = final["Login_sec"].apply(seconds_to_time)
        final["Total Break"] = final["Break_sec"].apply(seconds_to_time)
        final["Total Meeting"] = final["Meeting_sec"].apply(seconds_to_time)
        final["Total Net Login"] = final["NetLogin_sec"].apply(seconds_to_time)
        final["Total Talk Time"] = final["Talk_sec"].apply(seconds_to_time)
        final["AHT"] = final["AHT_sec"].apply(seconds_to_time)

        final = final[[
            "Employee ID",
            "Agent Name",
            "Total Login Time",
            "Total Break",
            "Total Meeting",
            "Total Net Login",
            "Total Talk Time",
            "Total Mature",
            "IB Mature",
            "OB Mature",
            "AHT"
        ]]

        # Convert to int (remove .0)
        final["Total Mature"] = final["Total Mature"].astype(int)
        final["IB Mature"] = final["IB Mature"].astype(int)
        final["OB Mature"] = final["OB Mature"].astype(int)

        return render_template(
            "result.html",
            data=final.to_dict("records")
        )

    return render_template("index.html")


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
