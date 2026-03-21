import streamlit as st
import requests
import openpyxl
import os

# ---------------- API CONFIG ----------------
INSTANCE_KEY = "instance143653"
TOKEN = "your_token_here"   # replace later with secrets
API_URL = f"https://api.ultramsg.com/{INSTANCE_KEY}/messages/chat"

EXCEL_FILE = "Student_Data.xlsx"

# ---------------- EXCEL SETUP ----------------
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Student Data"
    sheet.append(["Student Name", "Phone Number", "Course Name", "Parent Name", "Parent Contact"])
    wb.save(EXCEL_FILE)

def save_to_excel(data):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active
    sheet.append(data)
    wb.save(EXCEL_FILE)

# ---------------- WHATSAPP FUNCTION ----------------
def send_whatsapp_message(phone_number, message):
    if not phone_number.startswith("+91"):
        phone_number = "+91" + phone_number.lstrip("0")

    payload = {"to": phone_number, "body": message}

    try:
        response = requests.post(f"{API_URL}?token={TOKEN}", json=payload)
        return response.json()
    except Exception as e:
        return {"error": str(e)}

# ---------------- UI ----------------
st.set_page_config(page_title="Visitor Management System", layout="centered")

st.title("Visitor Management System")

menu = st.sidebar.radio("Menu", ["Visitor Entry", "Bulk Message"])

# =====================================================
# 👤 VISITOR ENTRY
# =====================================================
if menu == "Visitor Entry":

    st.header("Visitor Registration Form")

    student_name = st.text_input("Student Name")
    student_number = st.text_input("Student Contact Number")
    course_name = st.text_input("Course Name")
    parent_name = st.text_input("Parent Name")
    parent_contact = st.text_input("Parent Contact Number")

    if st.button("Submit"):
        if not all([student_name, student_number, course_name, parent_name, parent_contact]):
            st.error("All fields are required")
        else:
            save_to_excel([student_name, student_number, course_name, parent_name, parent_contact])

            message = f"""
Hello {student_name},

Welcome to Vikrant Group Of Institutions, Indore.

Thank you for visiting our campus.

Courses available:
Engineering, Management, Nursing, Pharmacy, Law

Scholarship:
https://www.vitm.edu.in/scholarship.html

Instagram:
https://www.instagram.com/vikrant.indore

Thanks
"""

            result = send_whatsapp_message(student_number, message)

            if "error" in result:
                st.error("Message failed")
            else:
                st.success("Visitor registered & message sent!")

# =====================================================
# 📤 BULK MESSAGE
# =====================================================
elif menu == "Bulk Message":

    st.header("Bulk WhatsApp Sender")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    manual_numbers = st.text_area("Or Enter Numbers (comma or line separated)")
    message = st.text_area("Enter Message")

    if st.button("Send Messages"):

        numbers = []

        # Excel Input
        if uploaded_file:
            wb = openpyxl.load_workbook(uploaded_file)
            sheet = wb.active

            header = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]

            if "Phone Number" not in header:
                st.error("'Phone Number' column missing")
            else:
                idx = header.index("Phone Number")

                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if row[idx]:
                        numbers.append(str(row[idx]))

        # Manual Input
        if manual_numbers:
            for num in manual_numbers.replace(',', '\n').split('\n'):
                if num.strip():
                    numbers.append(num.strip())

        if not numbers:
            st.error("No numbers found")
        else:
            for number in numbers:
                send_whatsapp_message(number, message)

            st.success("Bulk messages sent successfully!")
