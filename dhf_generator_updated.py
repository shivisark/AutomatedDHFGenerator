import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (SimpleDocTemplate, Paragraph, PageBreak,
                                LongTable, TableStyle)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

# === Paths ===
base_dir = os.path.dirname(__file__)
output_dir = os.path.join(base_dir, "output_reports")
os.makedirs(output_dir, exist_ok=True)

# File paths
device_overview_file = os.path.join(base_dir, "device_overview.txt")
user_needs_file = os.path.join(base_dir, "user_needs.csv")
design_inputs_file = os.path.join(base_dir, "design_inputs.csv")
design_outputs_file = os.path.join(base_dir, "design_outputs.csv")
verification_methods_file = os.path.join(base_dir, "verification_methods.csv")
risk_table_file = os.path.join(base_dir, "risk_table.csv")
change_log_file = os.path.join(base_dir, "change_log.csv")


dhf_pdf_output = os.path.join(output_dir, "DHF_Report.pdf")
traceability_matrix_output = os.path.join(output_dir, "traceability_matrix.csv")

# === Styles ===
styles = getSampleStyleSheet()
styleN = styles['Normal']
styleH = styles['Heading1']
styleCell = ParagraphStyle(name='CellStyle', parent=styleN, fontSize=7, leading=9, spaceAfter=2, wordWrap='CJK', splitLongWords=True, allowWidows=1, allowOrphans=1)

# === Helper Functions ===
def create_paragraph(text):
    return Paragraph(text.replace('\n', '<br/>'), styleN)

def create_table(df):
    from reportlab.platypus import Paragraph
    wrapped_data = [[Paragraph(f"<b>{' '.join(part.capitalize() for part in col.replace('_', ' ').split())}</b>", ParagraphStyle(name='HeaderCell', parent=styleCell, alignment=1, wordWrap='LTR', splitLongWords=False)) for col in df.columns]]
    for row in df.astype(str).values.tolist():
        wrapped_row = [Paragraph(cell.replace("<", "&lt;").replace(">", "&gt;"), styleCell) for cell in row]
        wrapped_data.append(wrapped_row)

    total_width = 7.3 * inch
    num_cols = len(df.columns)
    col_width = total_width / num_cols
    if num_cols > 6:
        col_width -= 0.05 * inch
    table = LongTable(wrapped_data, repeatRows=1, colWidths=[col_width] * num_cols)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    return table

# === Load Data ===
with open(device_overview_file, 'r') as f:
    device_overview = f.read()

user_needs = pd.read_csv(user_needs_file)
design_inputs = pd.read_csv(design_inputs_file)
design_outputs = pd.read_csv(design_outputs_file)
verification_methods = pd.read_csv(verification_methods_file)
risk_table = pd.read_csv(risk_table_file)
change_log = pd.read_csv(change_log_file)


# === Create Traceability Matrix ===
trace_matrix = user_needs.rename(columns={"Need ID": "Need_ID", "Description": "User_Need_Description"}) \
    .merge(design_inputs.rename(columns={"Input ID": "Input_ID", "Description": "Input_Description", "Linked Need ID": "Need_ID"}), on="Need_ID", how="left") \
    .merge(design_outputs.rename(columns={"Output ID": "Output_ID", "Description": "Output_Description", "Linked Input ID": "Input_ID"}), on="Input_ID", how="left") \
    .merge(verification_methods.rename(columns={"Input ID": "Input_ID", "Verification Method": "Verification_Method"}), on="Input_ID", how="left")

trace_matrix.to_csv(traceability_matrix_output, index=False)

# === Generate PDF ===
doc = SimpleDocTemplate(dhf_pdf_output, pagesize=A4, leftMargin=0.4*inch, rightMargin=0.4*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
elements = []

# Add each section to the PDF
elements.append(Paragraph("Device Overview", styleH))
elements.append(create_paragraph(device_overview))
elements.append(PageBreak())

elements.append(Paragraph("User Needs", styleH))
elements.append(create_table(user_needs))
elements.append(PageBreak())

elements.append(Paragraph("Design Inputs", styleH))
elements.append(create_table(design_inputs))
elements.append(PageBreak())

elements.append(Paragraph("Design Outputs", styleH))
elements.append(create_table(design_outputs))
elements.append(PageBreak())

elements.append(Paragraph("Verification Methods", styleH))
elements.append(create_table(verification_methods))
elements.append(PageBreak())

elements.append(Paragraph("Risk Management", styleH))
elements.append(create_table(risk_table))
elements.append(PageBreak())

elements.append(Paragraph("Change Log", styleH))
elements.append(create_table(change_log))
elements.append(PageBreak())

elements.append(Paragraph("Traceability Matrix", styleH))
elements.append(create_table(trace_matrix))


# Build PDF
doc.build(elements)
print("DHF report and traceability matrix generated successfully.")
