from flask import Flask, render_template, request
import pandas as pd
from datetime import datetime

app = Flask(__name__)

# =========================
# TIME FUNCTIONS
# =========================

def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        parts = str(t).split(":")
        if len(parts) != 3:
            return 0
        h, m, s = map(int, parts)
        return h*3600 + m*60 + s
    except:
        return 0


def seconds_to_time(sec):
    try:
        sec = int(sec)
        h = sec // 3600
        m = (sec % 3600) // 60
        s = sec % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


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

    try:

        agent_file = request.files['agent_file']
        cdr_file = request.files['cdr_file']

        # =========================
        # LOAD AGENT PERFORMANCE
        # =========================

        agent = pd.read_excel(agent_file, header=None)

        # ignore first 2 rows
        agent = agent.iloc[2:].reset_index(drop=True)

        # replace '-' with 0
        agent = agent.replace('-', 0)

        agent['Employee ID'] = agent.iloc[:,1].astype(str).str.strip()
        agent['Agent Name'] = agent.iloc[:,2].astype(str).str.strip()

        agent['Total Login Time'] = agent.iloc[:,3]
        agent['Total Talk Time'] = agent.iloc[:,5]

        # Break = T + W + Y
        break_sec = (
            agent.iloc[:,19].apply(time_to_seconds) +
            agent.iloc[:,22].apply(time_to_seconds) +
            agent.iloc[:,24].apply(time_to_seconds)
        )

        agent['Total Break'] = break_sec.apply(seconds_to_time)

        # Meeting = U + X
        meeting_sec = (
            agent.iloc[:,20].apply(time_to_seconds) +
            agent.iloc[:,23].apply(time_to_seconds)
        )

        agent['Total Meeting'] = meeting_sec.apply(seconds_to_time)

        # Net Login
        net_sec = agent.iloc[:,3].apply(time_to_seconds) - break_sec

        agent['Total Net Login'] = net_sec.apply(seconds_to_time)


        # =========================
        # LOAD CDR
        # =========================

        cdr = pd.read_excel(cdr_file, header=None)

        # ignore first row
        cdr = cdr.iloc[1:].reset_index(drop=True)

        cdr['Employee ID'] = cdr.iloc[:,1].astype(str).str.strip()
        cdr['Call Type'] = cdr.iloc[:,6].astype(str).str.strip()
        cdr['Call Status'] = cdr.iloc[:,25].astype(str).str.strip()

        # Total Mature
        total_mature = cdr[
            cdr['Call Status'].isin(['CALLMATURED','TRANSFER'])
        ].groupby('Employee ID').size().reset_index(name='Total Mature')

        # IB Mature
        ib_mature = cdr[
            (cdr['Call Status'].isin(['CALLMATURED','TRANSFER'])) &
            (cdr['Call Type'] == 'CSRINBOUND')
        ].groupby('Employee ID').size().reset_index(name='IB Mature')

        cdr_summary = pd.merge(total_mature, ib_mature,
                               on='Employee ID', how='left')

        cdr_summary['IB Mature'] = cdr_summary['IB Mature'].fillna(0)

        cdr_summary['OB Mature'] = (
            cdr_summary['Total Mature'] -
            cdr_summary['IB Mature']
        )


        # =========================
        # MERGE
        # =========================

        final = pd.merge(agent, cdr_summary,
                         on='Employee ID', how='left')

        final['Total Mature'] = final['Total Mature'].fillna(0).astype(int)
        final['IB Mature'] = final['IB Mature'].fillna(0).astype(int)
        final['OB Mature'] = final['OB Mature'].fillna(0).astype(int)


        # =========================
        # AHT
        # =========================

        def calc_aht(row):

            if row['Total Mature'] == 0:
                return "00:00:00"

            talk_sec = time_to_seconds(row['Total Talk Time'])

            return seconds_to_time(
                talk_sec / row['Total Mature']
            )

        final['AHT'] = final.apply(calc_aht, axis=1)


        # =========================
        # FINAL COLUMNS
        # =========================

        final = final[[
            'Employee ID',
            'Agent Name',
            'Total Login Time',
            'Total Net Login',
            'Total Break',
            'Total Meeting',
            'Total Mature',
            'IB Mature',
            'OB Mature',
            'Total Talk Time',
            'AHT'
        ]]

        report_time = datetime.now().strftime(
            "%d %b %Y %I:%M %p"
        )

        return render_template(
            "result.html",
            rows=final.to_dict("records"),
            report_time=report_time
        )

    except Exception as e:
        return f"Error: {str(e)}"


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run(debug=True)
