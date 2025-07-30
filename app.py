import streamlit as st
from docx import Document
from io import BytesIO
import ast
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


st.title("📄 Dictionary to DOCX / PDF Converter")

st.markdown("Paste a list of dictionaries like this:")
lec = st.text_input('Lecture')
da = datetime.today().date()
st.write(f'Today: {da}')
dict_input = st.text_area("📋 Paste your dictionary here:", height=250)

if st.button("📄 Generate DOCX and PDF"):
    try:
        # Safely evaluate string to Python object
        data = ast.literal_eval(dict_input)

        if not isinstance(data, list) or not all(isinstance(i, dict) for i in data):
            st.error("Input must be a list of dictionaries.")
        else:
            # --- Generate DOCX ---
            doc = Document()
            doc.add_heading(f"Attendance Report", 0)
            doc.add_paragraph(f'Date: {da}\nSubject: {lec}')
            
            table = doc.add_table(rows=1, cols=len(data[0]))
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, key in enumerate(data[0].keys()):
                hdr_cells[i].text = key

            for row in data:
                row_cells = table.add_row().cells
                for i, val in enumerate(row.values()):
                    row_cells[i].text = str(val)

            doc_buffer = BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)

            # --- Generate PDF ---
            pdf_buffer = BytesIO()
            p = canvas.Canvas(pdf_buffer, pagesize=A4)
            width, height = A4
            p.setFont("Helvetica-Bold", 14)
            p.drawString(100, height - 50, f"Attendance Report")
            p.setFont("Helvetica", 12)
            p.drawString(100, height - 70, f"Date: {da}")
            p.drawString(100, height - 90, f"Subject: {lec}")

            y = height - 120
            headers = list(data[0].keys())
            p.setFont("Helvetica-Bold", 11)
            for i, h in enumerate(headers):
                p.drawString(50 + i*100, y, str(h))

            y -= 20
            p.setFont("Helvetica", 10)
            for row in data:
                for i, val in enumerate(row.values()):
                    p.drawString(50 + i*100, y, str(val))
                y -= 20
                if y < 50:  # Create new page if space runs out
                    p.showPage()
                    y = height - 50

            p.save()
            pdf_buffer.seek(0)

            # --- Download Buttons ---
            st.success("✅ Files generated successfully!")

            st.download_button(
                label="📥 Download DOCX",
                data=doc_buffer,
                file_name="attendance_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            st.download_button(
                label="📥 Download PDF",
                data=pdf_buffer,
                file_name="attendance_report.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error: {e}")
