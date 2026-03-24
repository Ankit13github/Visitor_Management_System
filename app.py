
from flask import Flask, request, render_template, redirect, send_file, session, redirect, url_for
import pandas as pd
import io
import requests
import os
import logging
import psycopg2

app = Flask(__name__)

# ---------------- LOGGING ----------------
logging.basicConfig(level=logging.INFO)

# ---------------- DATABASE ----------------
DATABASE_URL = os.environ.get("DATABASE_URL")

def get_connection():
    return psycopg2.connect(DATABASE_URL)

# Create table
def create_table():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS visitors (
        id SERIAL PRIMARY KEY,
        student_name TEXT,
        student_number TEXT,
        course_name TEXT,
        parent_name TEXT,
        parent_contact TEXT
    );
    """)

    conn.commit()
    cur.close()
    conn.close()

create_table()

# ---------------- SAVE TO DB ----------------
def save_to_db(data):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    INSERT INTO visitors 
    (student_name, student_number, course_name, parent_name, parent_contact)
    VALUES (%s, %s, %s, %s, %s)
    """, data)

    conn.commit()
    cur.close()
    conn.close()

# ---------------- CHECK DUPLICATE ----------------
def is_duplicate(phone):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM visitors WHERE student_number = %s", (phone,))
    result = cur.fetchone()

    cur.close()
    conn.close()

    return result is not None

# ---------------- GET TOTAL ----------------
def get_total():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM visitors")
    total = cur.fetchone()[0]

    cur.close()
    conn.close()
    return total

# ---------------- GET ALL VISITORS ----------------
def get_all_visitors(filter_type=None):
    conn = get_connection()
    cur = conn.cursor()

    if filter_type == "today":
        cur.execute("""
        SELECT * FROM visitors 
        WHERE DATE(created_at) = CURRENT_DATE
        ORDER BY id DESC
        """)

    elif filter_type == "week":
        cur.execute("""
        SELECT * FROM visitors 
        WHERE created_at >= CURRENT_DATE - INTERVAL '7 days'
        ORDER BY id DESC
        """)

    else:
        cur.execute("SELECT * FROM visitors ORDER BY id DESC")
        
    rows = cur.fetchall()

    headers = ["ID", "Student Name", "Phone", "Course", "Parent", "Parent Contact"]

    cur.close()
    conn.close()

    return headers, rows

# ---------------- API CONFIG ----------------
INSTANCE_KEY = "instance143653"
TOKEN = os.environ.get("TOKEN")
API_URL = f"https://api.ultramsg.com/{INSTANCE_KEY}/messages/chat"

# ---------------- Admin Login ----------------
app.secret_key = "super_secret_key_123"

# ---------------- WHATSAPP FUNCTION ----------------
def send_whatsapp_message(phone_number, message):
    if not TOKEN:
        return {"error": "Token not configured"}

    if not phone_number.startswith("+91"):
        phone_number = "+91" + phone_number.lstrip("0")

    payload = {"to": phone_number, "body": message}

    try:
        response = requests.post(f"{API_URL}?token={TOKEN}", json=payload, timeout=10)
        return response.json()
    except Exception as e:
        return {"error": str(e)}

# ---------------- HOME ----------------
@app.route("/")
def home():
    return render_template("home.html")

# ---------------- VISITOR FORM ----------------
@app.route("/visitor_form")
def visitor_form():
    return render_template("index.html")

# ---------------- SEND MESSAGE ----------------
@app.route("/send_message", methods=["POST"])
def send_message():
    student_name = request.form.get("student_name", "").strip()
    student_number = request.form.get("student_number", "").strip()
    course_name = request.form.get("course_name", "").strip()
    parent_name = request.form.get("parent_name", "").strip()
    parent_contact = request.form.get("parent_contact", "").strip()

    if not all([student_name, student_number, course_name, parent_name, parent_contact]):
        return "❌ All fields are required."

    if len(student_number) < 10:
        return "❌ Invalid phone number."

    if is_duplicate(student_number):
        return "⚠️ Visitor already exists."

    # Save to DB
    save_to_db([student_name, student_number, course_name, parent_name, parent_contact])
    logging.info(f"New visitor added: {student_name}")

    message = f"""
Hello {student_name},

Welcome to Vikrant Group Of Institutions, Indore.

Thank you for visiting our campus.

Courses: Engineering, Management, Nursing, Pharmacy, Law

Scholarship:
https://www.vitm.edu.in/scholarship.html

Instagram:
https://www.instagram.com/vikrant.indore

Thanks
"""

    result = send_whatsapp_message(student_number, message)

    if "error" in result:
        return f"⚠️ Saved but message failed: {result['error']}"

    return redirect("/visitor_form?success=1")

# ---------------- BULK MESSAGE ----------------
@app.route("/bulk_message", methods=["GET", "POST"])
def bulk_message():
    if not session.get("admin"):
        return redirect("/login")
        
    if request.method == "POST":
        message = request.form.get("message")
        manual_input = request.form.get("manual_numbers", "")

        numbers = []

        if manual_input.strip():
            for num in manual_input.replace(',', '\n').split('\n'):
                if num.strip():
                    numbers.append(num.strip())

        if not numbers:
            return "⚠️ No numbers found."

        failed = []

        for number in numbers:
            result = send_whatsapp_message(number, message)
            if "error" in result:
                failed.append(number)

        if failed:
            return f"❌ Failed for: {', '.join(failed)}"

        return "✅ Bulk messages sent successfully!"

    return render_template("bulk_message.html")

# ---------------- DASHBOARD ----------------
@app.route("/dashboard")
def dashboard():
    if not session.get("admin"):
        return redirect("/login")
        
    total = get_total()
    return render_template("dashboard.html", total=total)

# ---------------- VIEW VISITORS ----------------
@app.route("/view_visitors")
def view_visitors():
    if not session.get("admin"):
        return redirect("/login")

    filter_type = request.args.get("filter")
    headers, rows = get_all_visitors(filter_type)
    return render_template("view.html", headers=headers, rows=rows)

# ---------------- DOWNLOAD FILE ----------------
@app.route("/download")
def download():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM visitors ORDER BY id DESC")
    rows = cur.fetchall()

    headers = ["ID", "Student Name", "Phone", "Course", "Parent", "Parent Contact"]

    cur.close()
    conn.close()

    # Convert to DataFrame
    df = pd.DataFrame(rows, columns=headers)

    # Save to Excel in memory
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="visitors.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ---------------- DELETE VISITOR FUNCTION ----------------
@app.route("/delete/<int:id>")
def delete_visitor(id):
    if not session.get("admin"):
        return redirect("/login")
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("DELETE FROM visitors WHERE id = %s", (id,))

    conn.commit()
    cur.close()
    conn.close()

    return redirect("/view_visitors")

# ---------------- EDIT VISITOR  ----------------
@app.route("/edit/<int:id>", methods=["GET", "POST"])
def edit_visitor(id):
    if not session.get("admin"):
        return redirect("/login")
    
    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":
        student_name = request.form.get("student_name")
        student_number = request.form.get("student_number")
        course_name = request.form.get("course_name")
        parent_name = request.form.get("parent_name")
        parent_contact = request.form.get("parent_contact")

        cur.execute("""
        UPDATE visitors
        SET student_name=%s, student_number=%s, course_name=%s,
            parent_name=%s, parent_contact=%s
        WHERE id=%s
        """, (student_name, student_number, course_name, parent_name, parent_contact, id))

        conn.commit()
        cur.close()
        conn.close()

        return redirect("/view_visitors")

    # GET request (load data)
    cur.execute("SELECT * FROM visitors WHERE id = %s", (id,))
    visitor = cur.fetchone()

    cur.close()
    conn.close()

    return render_template("edit.html", visitor=visitor)

# ---------------- Login Method----------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")

        if username == "admin" and password == "1234":
            session["admin"] = True
            return redirect("/dashboard")
        else:
            return "❌ Invalid credentials"

    return render_template("login.html")

# ---------------- Logout Config ----------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# ---------------- ERROR ----------------
@app.errorhandler(404)
def not_found(e):
    return "<h1>Page Not Found</h1>", 404

# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
