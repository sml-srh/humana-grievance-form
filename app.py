import streamlit as st
import pandas as pd
from fpdf import FPDF
import os
import zipfile
from io import BytesIO
from datetime import date

# PDF Generation Function
def generate_pdfs(df):
    class PDF(FPDF):
        def header(self):
            self.set_font("Arial", "B", 12)
            self.cell(0, 10, "Humana Grievance Form", align="C", ln=True)

        def add_response(self, header, response):
            self.set_font("Arial", size=12)
            for question, answer in zip(header, response):
                self.multi_cell(0, 10, f"{question}: {answer}", border=0)
                self.ln(2)  # Add spacing after each response

    header = df.columns.tolist()
    pdf_files = []

    for index, row in df.iterrows():
        if not row.isnull().all():  # Skip rows without submissions
            pdf = PDF()
            pdf.add_page()
            pdf.set_left_margin(10)
            pdf.set_right_margin(10)
            response = row.tolist()
            pdf.add_response(header, response)

            # Save PDF content as bytes using dest='S'
            pdf_content = pdf.output(dest='S').encode('latin1')
            pdf_buffer = BytesIO(pdf_content)
            pdf_files.append((f"Response_{index + 1}.pdf", pdf_buffer))
    
    return pdf_files

# Streamlit App Interface
st.title("Humana Grievance Form")

uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df = df.iloc[:, 5:]  # Ignore the first 5 columns

    # Generate PDFs
    pdf_files = generate_pdfs(df)

    # Create an in-memory ZIP file with the current date in the name
    today = date.today().strftime("%Y-%m-%d")
    zip_file_name = f"Humana Grievance Forms_{today}.zip"
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for file_name, pdf_buffer in pdf_files:
            zip_file.writestr(file_name, pdf_buffer.getvalue())
    zip_buffer.seek(0)

    # Provide download link for the ZIP file
    st.success("PDFs have been generated successfully!")
    st.download_button(
        label="Download All PDFs as ZIP",
        data=zip_buffer,
        file_name=zip_file_name,
        mime="application/zip"
    )
