from flask import Flask, request, render_template
import requests
import openpyxl
import os

app = Flask(__name__)

# UltraMsg API credentials
INSTANCE_KEY = "instance143653"
TOKEN = os.environ.get("8i34klnikkn43e2t")
API_URL = f"https://api.ultramsg.com/{INSTANCE_KEY}/messages/chat"

EXCEL_FILE = "Student_Data.xlsx"

# Create Excel file if not exists
if not os.path.exists(EXCEL_FILE):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Student Data"
    sheet.append(["Student Name", "Phone Number", "Course Name", "Parent Name", "Parent Contact"])
    workbook.save(EXCEL_FILE)

# Save to Excel
def save_to_excel(data):
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    sheet.append(data)
    workbook.save(EXCEL_FILE)

# Send WhatsApp message
def send_whatsapp_message(phone_number, message):
    # Delete if you tokens not added
    if not TOKEN:
        return {"error": "Token not configured"}
        
    if not phone_number.startswith("+91"):
        phone_number = "+91" + phone_number.lstrip("0")
    payload = {"to": phone_number, "body": message}
    try:
        response = requests.post(f"{API_URL}?token={TOKEN}", json=payload)
        return response.json()
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}

# Home Page
@app.route("/")
def home():
    return '''
    <html>
    <head><title>Visitor Management System</title></head>
    <body style="text-align:center; font-family:Arial; background:#f0f4f8; padding-top:50px;">
        <h1 style="color:#6a11cb;">Visitor Management System</h1>
        <a href="/visitor_form" style="padding:15px 25px; margin:10px; background:#6a11cb; color:white; border-radius:8px; text-decoration:none;">Individual Visitor Message</a>
        <a href="/bulk_message" style="padding:15px 25px; margin:10px; background:#6a11cb; color:white; border-radius:8px; text-decoration:none;">Bulk WhatsApp Sender</a>
    </body>
    </html>
    '''

# Visitor Form Page
@app.route("/visitor_form")
def visitor_form():
    return render_template("index.html")

# Handle Individual Form Submission
@app.route("/send_message", methods=["POST"])
def send_message():
    student_name = request.form.get("student_name", "").strip()
    student_number = request.form.get("student_number", "").strip()
    course_name = request.form.get("course_name", "").strip()
    parent_name = request.form.get("parent_name", "").strip()
    parent_contact = request.form.get("parent_contact", "").strip()

    if not all([student_name, student_number, course_name, parent_name, parent_contact]):
        return "Error: All fields are required."

    save_to_excel([student_name, student_number, course_name, parent_name, parent_contact])

    message = f"""
Hello {student_name},

Welcome to Vikrant Group Of Institutions, Indore.

Thank you very much for visiting the campus.

Vikrant Group has been into existence since more then 20 years now, with more than 10,000/- Alumni working in India and abroad. 

VGI is running UG & PG courses in following institutes:- 

Engineering
Management 
Nursing 
Pharmacy
Law

All courses are approved by UGC, AICTE, PCI, BCI and MPNRC and affiliated to state Govt. Universities.

For various Discount and Scholarship Schemes click on the link below:- 
https://www.vitm.edu.in/scholarship.html

Stay updated and connected by following our official Instagram page: [vikrant.indore]
https://www.instagram.com/vikrant.indore

We hope your visit is both enjoyable and inspiring. Feel free to reach out for more details.

Thanks"""
    result = send_whatsapp_message(student_number, message)

    if "error" in result:
        return f"Message not sent: {result['error']}"
    return f"✅ Message sent to {student_name} ({student_number})!"

# Bulk Message Page and Logic
@app.route("/bulk_message", methods=["GET", "POST"])
def bulk_message():
    if request.method == "POST":
        message = request.form.get("message")
        manual_input = request.form.get("manual_numbers", "")
        file = request.files.get("file")
        numbers = []

        # Read from Excel file
        if file and file.filename.endswith(".xlsx"):
            workbook = openpyxl.load_workbook(file)
            sheet = workbook.active
            # Find the "Phone Number" column index
            header = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            try:
                phone_idx = header.index("Phone Number")
            except ValueError:
                return "⚠️ 'Phone Number' column not found in uploaded file."
            for row in sheet.iter_rows(min_row=2, values_only=True):
                phone = row[phone_idx]
                if phone:
                    numbers.append(str(phone))

        # Read from manual input
        if manual_input.strip():
            for entry in manual_input.replace(',', '\n').split('\n'):
                num = entry.strip()
                if num:
                    numbers.append(num)

        if not numbers:
            return "⚠️ Please upload a file or enter numbers manually."

        failed = []
        for number in numbers:
            result = send_whatsapp_message(number, message)
            if "error" in result:
                failed.append(number)

        if failed:
            return f"❌ Failed for: {', '.join(failed)}"
        return "✅ Bulk messages sent successfully!"

    return render_template("bulk_message.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
