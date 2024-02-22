import pandas as pd
import openpyxl

# Read the Excel file into a pandas DataFrame
def read_excel_file(file_path):
    df = pd.read_excel(file_path)
    return df

# Convert specified exam columns to numeric type
def convert_to_numeric(df, exam_columns):
    for col in exam_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

# Find the top scorers based on each exam column
def find_top_scorers(df, exam_columns, top_count):
    top_scorers = {}
    for col in exam_columns:
        top_scorers[col] = df.nlargest(top_count, col, 'all')
    return top_scorers

# Create DataFrame for each exam
def create_output_dfs(top_scorers):
    output_dfs = {}
    for exam, top_scorer_df in top_scorers.items():
        # Check if the exam column exists in the DataFrame
        if exam in top_scorer_df.columns:
            # Filter out rows with a value of 100 in the specified exam columns
            if exam != 'exam18':
                top_scorer_df = top_scorer_df[top_scorer_df[exam] != 100]
            else:
                top_scorer_df = top_scorer_df[top_scorer_df[exam] != 775]
            
            output_dfs[exam] = pd.DataFrame({
                'Sr. No.': range(1, len(top_scorer_df) + 1),
                'Roll No': top_scorer_df['ROLLNO'].tolist(),
                'Name': top_scorer_df['NAME'].tolist(),
                'Marks': top_scorer_df[exam].fillna(0).tolist()
            })
    return output_dfs






# Save the DataFrames to a new Excel file with formatting
def save_to_excel(output_dfs, output_file_path):
    start_row = 6  # Starting row for writing the output tables

    exam_mapping = {
        'exam3': 'Subject 1',
        'exam6': 'Subject 2',
        'exam9': 'Subject 3',
        'exam12': 'Subject 4',
        'exam15': 'Subject 5',
        'exam18': 'TOTAL'
    }

    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        for exam, output_df in output_dfs.items():
            # Map exam column names to new headings
            exam_heading = exam_mapping.get(exam, exam)

            output_df.to_excel(writer, index=False, sheet_name='Output', startrow=start_row, startcol=1)

            # Update formatting for the new data
            worksheet = writer.sheets['Output']
            for row in worksheet.iter_rows(min_row=start_row + 1, max_row=start_row + len(output_df), min_col=2, max_col=2):
                for cell in row:
                    cell.alignment = openpyxl.styles.Alignment(wrapText=True, vertical='center', horizontal='center')

            # Add a bold header with the new heading
            header_cell = worksheet.cell(row=start_row - 1, column=4, value=f"{exam_heading} Toppers")
            header_cell.font = openpyxl.styles.Font(bold=True)
            header_cell.alignment = openpyxl.styles.Alignment(horizontal='center')

            start_row += len(output_df) + 2  # Move to the next section for the next exam
            start_row += 3  # Add a gap of 3 rows

        # Auto-adjust column widths based on content for all tables
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

        # Save the workbook
        writer.book.save(output_file_path)
