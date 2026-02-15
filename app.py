from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
import glob
from datetime import timedelta

# Auto downloader
try:
    from downloader import download_reports
    AUTO_AVAILABLE = True
except:
    AUTO_AVAILABLE = False


app = Flask(__name__)


# =========================
# TIME FORMAT FUNCTION
# =========================

def format_time(seconds):

    if pd.isna(seconds):
        return "00:00:00"

    seconds = int(seconds)

    return str(timedelta(seconds=seconds))


# =========================
# PROCESS REPORT FUNCTION
# =========================

def process_reports(agent_file, cdr_file):

    # =====================
    # AGENT PERFORMANCE
    # =====================

    agent = pd.read_excel(agent_file, header=2)

    agent.replace("-", 0, inplace=True)

    agent['Employee ID'] = agent.iloc[:,1].astype(str)
    agent['Agent Name'] = agent.iloc[:,2]

    login = pd.to_timedelta(agent.iloc[:,3].astype(str), errors='coerce').dt.total_seconds()

    lunch = pd.to_timedelta(agent.iloc[:,19].astype(str), errors='coerce').dt.total_seconds()
    tea = pd.to_timedelta(agent.iloc[:,24].astype(str), errors='coerce').dt.total_seconds()
    short = pd.to_timedelta(agent.iloc[:,22].astype(str), errors='coerce').dt.total_seconds()

    meeting = pd.to_timedelta(agent.iloc[:,20].astype(str), errors='coerce').dt.total_seconds()
    system = pd.to_timedelta(agent.iloc[:,23].astype(str), errors='coerce').dt.total_seconds()

    talk = pd.to_timedelta(agent.iloc[:,5].astype(str), errors='coerce').dt.total_seconds()

    agent['Total Login'] = login
    agent['Total Break'] = lunch + tea + short
    agent['Total Meeting'] = meeting + system
    agent['Net Login'] = login - agent['Total Break']
    agent['Total Talk'] = talk

    agent_final = agent[['Employee ID','Agent Name','Total Login','Net Login','Total Break','Total Meeting','Total Talk']]


    # =====================
    # CDR REPORT
    # =====================

    cdr = pd.read_excel(cdr_file, header=1)

    cdr['Employee ID'] = cdr.iloc[:,1].astype(str)

    cdr['CallType'] = cdr.iloc[:,6].astype(str)
    cdr['CallStatus'] = cdr.iloc[:,25].astype(str)

    matured = cdr[cdr['CallStatus'].isin(['CALLMATURED','TRANSFER'])]

    total_mature = matured.groupby('Employee ID').size()

    ib_mature = matured[matured['CallType']=="CSRINBOUND"].groupby('Employee ID').size()

    cdr_summary = pd.DataFrame({
        'Total Mature': total_mature,
        'IB Mature': ib_mature
    }).fillna(0)

    cdr_summary['OB Mature'] = cdr_summary['Total Mature'] - cdr_summary['IB Mature']

    cdr_summary.reset_index(inplace=True)


    # =====================
    # MERGE BOTH
    # =====================

    final = pd.merge(agent_final, cdr_summary, on="Employee ID", how="left")

    final.fillna(0, inplace=True)


    # =====================
    # AHT
    # =====================

    final['AHT'] = final.apply(
        lambda x: x['Total Talk']/x['Total Mature'] if x['Total Mature']>0 else 0,
        axis=1
    )


    # =====================
    # FORMAT TIME
    # =====================

    final['Total Login'] = final['Total Login'].apply(format_time)
    final['Net Login'] = final['Net Login'].apply(format_time)
    final['Total Break'] = final['Total Break'].apply(format_time)
    final['Total Meeting'] = final['Total Meeting'].apply(format_time)
    final['AHT'] = final['AHT'].apply(format_time)


    # =====================
    # FINAL HEADERS
    # =====================

    final = final[[
        'Employee ID',
        'Agent Name',
        'Total Login',
        'Net Login',
        'Total Break',
        'Total Meeting',
        'AHT',
        'Total Mature',
        'IB Mature',
        'OB Mature'
    ]]

    return final


# =========================
# HOME PAGE
# =========================

@app.route('/')
def index():
    return render_template("index.html")


# =========================
# MANUAL GENERATE
# =========================

@app.route('/generate', methods=['POST'])
def generate():

    agent_file = request.files['agent']
    cdr_file = request.files['cdr']

    if not agent_file or not cdr_file:
        return "Upload both files"

    agent_path = "agent.xlsx"
    cdr_path = "cdr.xlsx"

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)

    df = process_reports(agent_path, cdr_path)

    return render_template(
        "result.html",
        tables=[df.to_html(classes="data", index=False)]
    )


# =========================
# AUTO DOWNLOAD GENERATE
# =========================

@app.route('/auto_generate')
def auto_generate():

    if not AUTO_AVAILABLE:
        return "Auto download not available"

    download_reports()

    folder = os.path.expanduser("~/Downloads")

    agent_file = max(
        glob.glob(os.path.join(folder,"*Agent*.xlsx")),
        key=os.path.getctime
    )

    cdr_file = max(
        glob.glob(os.path.join(folder,"*CdrReport*.xlsx")),
        key=os.path.getctime
    )

    df = process_reports(agent_file, cdr_file)

    return render_template(
        "result.html",
        tables=[df.to_html(classes="data", index=False)]
    )


# =========================

if __name__ == "__main__":
    app.run(debug=True)
