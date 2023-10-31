import os
import re
import csv
from docx import Document

def extract_info_from_docx(docx_path):
    doc = Document(docx_path)
    student_name = ""
    student_id = ""
    companies = []
    in_company_column = False

    for table in doc.tables:
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                
                if cell_text:  # Only process non-empty cells
                    print(f"Reading cell: {cell_text}")  # Debugging line

                    if "Student Name + Student ID" in cell_text:
                        name_id_match = re.search(r"([A-Za-z ]+)(?:\s*(2005\d{5}))?", cell_text.replace("Student Name + Student ID", ""))
                        if name_id_match:
                            student_name = name_id_match.group(1).strip()
                            student_id = name_id_match.group(2) if name_id_match.group(2) else "N/A"

                    if "Company and Location of Role" in cell_text:
                        in_company_column = True  # Flag to indicate we are in the correct column
                        continue  # Skip the header itself

                    if in_company_column and i == 0:  # Only add companies from the first column after the flag is set
                        companies.append(cell_text)

    application_count = len(companies)
    return student_name, student_id, companies, application_count

if __name__ == "__main__":
    print("Script started")  # Debugging line
    output_data = []
    failed_files = []

    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    print(f"Changed working directory to {script_dir}")  # Debugging line

    for filename in os.listdir('.'):
        print(f"Checking file: {filename}")  # Debugging line
        if filename.endswith('.pdf') or filename.endswith('.docx'):
            print(f"Reading {filename}")  # Debugging line
            try:
                if filename.endswith('.pdf'):
                    student_name, student_id, companies, application_count = extract_info_from_pdf(filename)
                else:
                    student_name, student_id, companies, application_count = extract_info_from_docx(filename)
                
                output_data.append({
                    'Student Name': student_name,
                    'Student ID': student_id,
                    'Companies': ', '.join(companies),
                    'Application Count': application_count,
                    'File Name': filename
                })
            except Exception as e:
                print(f"Failed to read {filename}: {e}")
                failed_files.append(filename)

    print("Writing to output.csv")  # Debugging line
    with open("output.csv", "w", newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Student Name', 'Student ID', 'Companies', 'Application Count', 'File Name']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for data in output_data:
            writer.writerow(data)

    print("Data written to output.csv")

    if failed_files:
        with open("log.txt", "w") as log_file:
            log_file.write("Failed to read the following files:\n")
            for file in failed_files:
                log_file.write(f"{file}\n")

        print("Log written to log.txt")
