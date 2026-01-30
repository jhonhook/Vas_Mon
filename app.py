from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
from fpdf import FPDF

app = Flask(__name__)
app.secret_key = "vas_secret_key"

# ---------------- Files ----------------
IN_PROGRESS_FILE = "Vas_in_progress.xlsx"
DONE_FILE = "Vas_Done.xlsx"

# Columns
COLUMNS_IN = ["Name", "Shift", "PLT ID", "Status", "Date", "In Time"]
COLUMNS_DONE = ["Name", "Shift", "PLT ID", "Date", "In Time", "Out Time", "Total Time"]

# Dummy users
USERS = ["Ali", "Ahmed", "Sara", "Usman"]

# Admin
ADMIN_USER = "admin"
ADMIN_PASS = "1234"

# ---------------- Initialize Files ----------------
def init_files():
    if not os.path.exists(IN_PROGRESS_FILE):
        pd.DataFrame(columns=COLUMNS_IN).to_excel(IN_PROGRESS_FILE, index=False)
    if not os.path.exists(DONE_FILE):
        pd.DataFrame(columns=COLUMNS_DONE).to_excel(DONE_FILE, index=False)
init_files()

def load(file):
    return pd.read_excel(file) if os.path.exists(file) else pd.DataFrame()

# ---------------- User Entry ----------------
@app.route("/", methods=["GET","POST"])
def index():
    df = load(IN_PROGRESS_FILE)
    active_ids = df["PLT ID"].astype(str).tolist() if not df.empty else []

    if request.method=="POST":
        name = request.form["name"]
        status = request.form["status"]
        now = datetime.now()

        if status=="In":
            # New In entry
            new_row = {
                "Name": name,
                "Shift": request.form["shift"],
                "PLT ID": request.form["plt_id_in"],
                "Status": "In",
                "Date": now.strftime("%Y-%m-%d"),
                "In Time": now.strftime("%H:%M:%S")
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            df.to_excel(IN_PROGRESS_FILE, index=False)
        else:
            # Out process
            plt_id = request.form["plt_id_out"]
            record = df[df["PLT ID"].astype(str)==plt_id]

            if record.empty:
                return "<script>alert('Error: Selected PLT ID not found in In-Progress!'); window.location.href='/';</script>"

            in_user = record.iloc[0]["Name"]
            if in_user != name:
                return f"<script>alert('Error: Only {in_user} can mark this PLT ID as Out!'); window.location.href='/';</script>"

            in_time_str = record.iloc[0]["In Time"]
            in_date_str = record.iloc[0]["Date"]
            out_time = now

            # ---------------- Correct Total Time Calculation ----------------
            in_datetime = datetime.strptime(f"{in_date_str} {in_time_str}", "%Y-%m-%d %H:%M:%S")
            if out_time < in_datetime:
                # Rare case: next day
                in_datetime -= timedelta(days=1)

            tdelta = out_time - in_datetime
            total_seconds = int(tdelta.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            total_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

            done_df = load(DONE_FILE)
            done_row = {
                "Name": in_user,
                "Shift": record.iloc[0]["Shift"],
                "PLT ID": plt_id,
                "Date": record.iloc[0]["Date"],
                "In Time": in_time_str,
                "Out Time": out_time.strftime("%H:%M:%S"),
                "Total Time": total_time
            }
            done_df = pd.concat([done_df, pd.DataFrame([done_row])], ignore_index=True)
            done_df.to_excel(DONE_FILE, index=False)

            # Remove from In-progress
            df = df[df["PLT ID"].astype(str)!=plt_id]
            df.to_excel(IN_PROGRESS_FILE, index=False)

        return redirect("/")

    return render_template("index.html", users=USERS, active_ids=active_ids)

# ---------------- Admin Login ----------------
@app.route("/admin", methods=["GET","POST"])
def admin_login():
    if request.method=="POST":
        if request.form["username"]==ADMIN_USER and request.form["password"]==ADMIN_PASS:
            session["admin"]=True
            return redirect("/dashboard")
    return render_template("admin_login.html")

# ---------------- Admin Dashboard ----------------
@app.route("/dashboard")
def dashboard():
    if not session.get("admin"):
        return redirect("/admin")

    in_df = load(IN_PROGRESS_FILE)
    done_df = load(DONE_FILE)

    if not in_df.empty:
        in_df['Date']=pd.to_datetime(in_df['Date']).dt.strftime("%Y-%m-%d")
        in_df['In Time']=in_df['In Time'].astype(str)
        in_df=in_df.sort_values(by=["Date","In Time"],ascending=False)

    if not done_df.empty:
        done_df['Date']=pd.to_datetime(done_df['Date']).dt.strftime("%Y-%m-%d")
        done_df['In Time']=done_df['In Time'].astype(str)
        done_df['Out Time']=done_df['Out Time'].astype(str)
        done_df['Total Time']=done_df['Total Time'].astype(str)
        done_df=done_df.sort_values(by=["Date","Out Time"],ascending=False)

    return render_template("admin_dashboard.html",
                           in_records=in_df.to_dict(orient="records"),
                           done_records=done_df.to_dict(orient="records"))

# ---------------- Styled Excel Download ----------------
@app.route("/download_styled/<file>")
def download_styled(file):
    if not session.get("admin"):
        return redirect("/admin")
    
    df = load(file)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook  = writer.book
        worksheet = writer.sheets["Sheet1"]

        # Header
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#667eea',
            'font_color':'white',
            'border':1
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Alternating row colors
        row_format1 = workbook.add_format({'bg_color': '#f2f2f2'})
        row_format2 = workbook.add_format({'bg_color': '#ffffff'})
        for i in range(1,len(df)+1):
            fmt = row_format1 if i%2==0 else row_format2
            worksheet.set_row(i, None, fmt)

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"Styled_{file}",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# ---------------- PDF Report ----------------
@app.route("/generate_report/pdf")
def generate_pdf():
    if not session.get("admin"):
        return redirect("/admin")
    df = load(DONE_FILE)

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial","B",16)
    pdf.cell(0,10,"VAS Completed Stock Report",0,1,'C')
    pdf.set_font("Arial","",12)
    pdf.cell(0,10,f"Date: {datetime.now().strftime('%Y-%m-%d')}",0,1,'C')
    pdf.ln(5)

    # Table header
    pdf.set_fill_color(102,126,234)
    pdf.set_text_color(255,255,255)
    pdf.set_font("Arial","B",11)
    col_widths=[20,25,25,20,25,25,25]
    headers=["Name","Shift","PLT ID","Date","In Time","Out Time","Total Time"]
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i],8,header,1,0,'C',1)
    pdf.ln()

    # Table rows
    pdf.set_fill_color(242,242,242)
    pdf.set_text_color(0,0,0)
    fill=True
    pdf.set_font("Arial","",10)
    for idx,row in df.iterrows():
        pdf.cell(col_widths[0],8,str(row['Name']),1,0,'C',fill)
        pdf.cell(col_widths[1],8,str(row['Shift']),1,0,'C',fill)
        pdf.cell(col_widths[2],8,str(row['PLT ID']),1,0,'C',fill)
        pdf.cell(col_widths[3],8,str(row['Date']),1,0,'C',fill)
        pdf.cell(col_widths[4],8,str(row['In Time']),1,0,'C',fill)
        pdf.cell(col_widths[5],8,str(row['Out Time']),1,0,'C',fill)
        pdf.cell(col_widths[6],8,str(row['Total Time']),1,0,'C',fill)
        pdf.ln()
        fill = not fill

    pdf_bytes = pdf.output(dest='S').encode('latin-1')
    return send_file(BytesIO(pdf_bytes),
                     as_attachment=True,
                     download_name="Completed_Report.pdf",
                     mimetype="application/pdf")

# ---------------- Logout ----------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/admin")

if __name__=="__main__":
    app.run(debug=True)
