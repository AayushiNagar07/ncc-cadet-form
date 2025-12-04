import streamlit as st
from openpyxl import load_workbook
import os

FILE_NAME = "Nominal Roll of B Cert Exam 2026.xlsx"

# Load Excel workbook
if os.path.exists(FILE_NAME):
    wb = load_workbook(FILE_NAME)
else:
    st.error("Excel file not found. Upload it beside the app.py file.")
    st.stop()

st.title("NCC – Cadet Nominal Roll Form (Certificate 'B' Exam 2026)")

st.write("Fill all details carefully. This form will be submitted to ANO & HQ.")

# ------- FORM UI --------

with st.form("cadet_form"):
    ser_no = st.text_input("Ser No")
    regtl_no = st.text_input("Regtl No")
    rank = st.text_input("Rank")
    cadet_name = st.text_input("Name of Cadet")
    father_name = st.text_input("Father Name")
    dob = st.date_input("Date of Birth")
    enroll_date = st.date_input("Date of Enrollment")
    discharge_date = st.date_input("Date of Discharge")
    camp_details = st.text_area("Details of Camps Attended")
    attendance_1 = st.text_input("Parade Attendance (I Yr) %")
    attendance_2 = st.text_input("Parade Attendance (II Yr) %")
    photo = st.file_uploader("Upload Cadet Photo (White background)", type=["jpg","jpeg","png"])

    submitted = st.form_submit_button("Submit Details")

# ------- SAVE TO EXCEL --------

if submitted:
    if cadet_name.strip() == "":
        st.error("Cadet Name is required!")
    else:
        # Create a separate sheet per cadet
        if cadet_name not in wb.sheetnames:
            ws = wb.create_sheet(title=cadet_name)
            ws.append(["Field", "Details"])
        else:
            ws = wb[cadet_name]

        ws.append(["Ser No", ser_no])
        ws.append(["Regtl No", regtl_no])
        ws.append(["Rank", rank])
        ws.append(["Name of Cadet", cadet_name])
        ws.append(["Father Name", father_name])
        ws.append(["Date of Birth", str(dob)])
        ws.append(["Enrollment Date", str(enroll_date)])
        ws.append(["Date of Discharge", str(discharge_date)])
        ws.append(["Camp Details", camp_details])
        ws.append(["Attendance I Yr", attendance_1])
        ws.append(["Attendance II Yr", attendance_2])

        # Save photo file
        if photo:
            photo_path = f"photos/{cadet_name}.jpg"
            os.makedirs("photos", exist_ok=True)
            with open(photo_path, "wb") as f:
                f.write(photo.read())
            ws.append(["Photo Path", photo_path])
        else:
            ws.append(["Photo Path", "Not Uploaded"])

        ws.append([])

        wb.save(FILE_NAME)
        st.success("Cadet details submitted successfully!")
        st.balloons()

