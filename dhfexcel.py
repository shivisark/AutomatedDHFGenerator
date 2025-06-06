import os
import pandas as pd
import xlsxwriter

# === Paths ===
base_dir = os.path.dirname(__file__)
output_dir = os.path.join(base_dir, "output_reports")
os.makedirs(output_dir, exist_ok=True)

# Input file paths
user_needs_file = os.path.join(base_dir, "user_needs.csv")
design_inputs_file = os.path.join(base_dir, "design_inputs.csv")
design_outputs_file = os.path.join(base_dir, "design_outputs.csv")
verification_methods_file = os.path.join(base_dir, "verification_methods.csv")
change_log_file = os.path.join(base_dir, "change_log.csv")
traceability_matrix_output = os.path.join(output_dir, "traceability_matrix.csv")
risk_table_file = os.path.join(base_dir, "risk_table.csv")

# Excel output path
excel_output_file = os.path.join(output_dir, "DHF_Report_Enhanced.xlsx")

# === Excel Writer ===
with pd.ExcelWriter(excel_output_file, engine='xlsxwriter') as writer:
    def write_sheet(file_path, sheet_name):
        if os.path.exists(file_path):
            df = pd.read_csv(file_path)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Wrap text in all columns
            worksheet = writer.sheets[sheet_name]
            workbook = writer.book
            wrap_format = workbook.add_format({'text_wrap': True})
            for col_num in range(len(df.columns)):
                worksheet.set_column(col_num, col_num, 20, wrap_format)

    # Write all sections to Excel
    write_sheet(user_needs_file, 'User Needs')
    write_sheet(design_inputs_file, 'Design Inputs')
    write_sheet(design_outputs_file, 'Design Outputs')
    write_sheet(verification_methods_file, 'Verification Methods')
    write_sheet(change_log_file, 'Change Log')
    write_sheet(traceability_matrix_output, 'Traceability Matrix')

    # === Risk Management Sheet with Python-based RPN Calculation ===
    if os.path.exists(risk_table_file):
        risk_df = pd.read_csv(risk_table_file)

        # Calculate RPN in Python
        if all(col in risk_df.columns for col in ['Severity', 'Occurrence', 'Detection']):
            risk_df['RPN'] = (
                pd.to_numeric(risk_df['Severity'], errors='coerce') *
                pd.to_numeric(risk_df['Occurrence'], errors='coerce') *
                pd.to_numeric(risk_df['Detection'], errors='coerce')
            )

        sheet_name = 'Risk Management'
        risk_df.to_excel(writer, sheet_name=sheet_name, index=False)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Apply wrap text to all columns
        wrap_format = workbook.add_format({'text_wrap': True})
        for col_num in range(len(risk_df.columns)):
            worksheet.set_column(col_num, col_num, 20, wrap_format)

        # Apply conditional formatting on RPN column
        if 'RPN' in risk_df.columns:
            rpn_col_idx = risk_df.columns.get_loc('RPN')
            rpn_col_letter = chr(65 + rpn_col_idx)
            worksheet.conditional_format(f'{rpn_col_letter}2:{rpn_col_letter}{len(risk_df)+1}', {
                'type': '3_color_scale',
                'min_color': "#63BE7B",
                'mid_color': "#FFEB84",
                'max_color': "#F8696B"
            })

            # Create a bar chart of RPNs
            chart = workbook.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'RPN',
                'categories': f"='{sheet_name}'!$A$2:$A${len(risk_df)+1}",
                'values':     f"='{sheet_name}'!${rpn_col_letter}$2:${rpn_col_letter}${len(risk_df)+1}",
                'fill': {'color': '#5B9BD5'}
            })
            chart.set_title({'name': 'Risk Priority Numbers'})
            chart.set_x_axis({'name': 'Failure Mode'})
            chart.set_y_axis({'name': 'RPN Value'})
            worksheet.insert_chart(f'{chr(65 + rpn_col_idx + 2)}2', chart)
