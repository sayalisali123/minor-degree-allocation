import pandas as pd
import re
import os
import glob
import pdfplumber


# --- 0. CONFIGURATION ---
MAX_MARKS = 800.0


# --- 1. PDF READER ---
def read_pdf_text(pdf_path):
    """Reads all text content from a PDF file using pdfplumber."""
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text(layout=True) + "\n"
        return text.strip()
    except Exception as e:
        print(f"ERROR reading {pdf_path}: {e}")
        return None


# --- 2. ROBUST DATA EXTRACTION LOGIC ---
def extract_student_data_final(student_text):
    """Extracts student personal info and subject-wise marks."""
    all_subjects_data = []

    # Extract personal info
    name_match = re.search(r"Name\s*:\s*([A-Z\s]+)", student_text, re.IGNORECASE)
    name = name_match.group(1).strip() if name_match else "UNKNOWN_NAME"

    prn_match = re.search(r"University PRN\s*:\s*(\d+)", student_text)
    prn = prn_match.group(1).strip() if prn_match else "UNKNOWN_PRN"

    # Subject block pattern
    subject_block_pattern = re.compile(
        r'(\d{5})\s+([^0-9\n]+?)(.*?)(?=\s*\d{5}\s+[^0-9\n]+?|Sem\s*-\s*6|\Z)',
        re.DOTALL
    )

    # Mark pattern
    simplified_mark_pattern = re.compile(
        r'(CIE|TW|ESEx|PR)\s*\(\d+\)\s*(\d+)\s*PASS'
    )

    for block_match in subject_block_pattern.finditer(student_text):
        subject_name = block_match.group(2).strip()
        mark_block = block_match.group(3)

        if "Paper / Subject Name" in subject_name or "Category Marks" in subject_name or not subject_name:
            continue

        current_subject_data = {
            'Name': name,
            'PRN': prn,
            'Subject': subject_name,
            'CIE': 'N/A',
            'TW': 'N/A',
            'ESEx': 'N/A',
            'PR': 'N/A',
        }

        for mark_match in simplified_mark_pattern.finditer(mark_block):
            category = mark_match.group(1).strip()
            mark = mark_match.group(2).strip()
            current_subject_data[category] = mark

        marks_to_sum = pd.Series({k: v for k, v in current_subject_data.items() if k in ['CIE', 'TW', 'ESEx', 'PR']})
        marks_to_sum = pd.to_numeric(marks_to_sum, errors='coerce')
        current_subject_data['Subject Total Marks'] = marks_to_sum.sum(skipna=True)

        if current_subject_data['Subject Total Marks'] > 0:
            all_subjects_data.append(current_subject_data)

    return all_subjects_data


# --- 3. MAIN PROCESSING AND RANKING FUNCTION ---
def make_result_analysis_and_rank(folder_path, max_marks):
    """Processes all PDF files, aggregates marks, and calculates final rank/percentage."""
    pdf_files = glob.glob(os.path.join(folder_path, '*.pdf'))
    if not pdf_files:
        print(f"Error: No PDF files found in {folder_path}.")
        return None, None

    print(f"Found {len(pdf_files)} PDF files to process...")
    all_students_subject_data = []

    for pdf_path in pdf_files:
        print(f"\nProcessing: {os.path.basename(pdf_path)}...")
        student_text = read_pdf_text(pdf_path)
        if student_text:
            extracted_data = extract_student_data_final(student_text)
            all_students_subject_data.extend(extracted_data)

    if not all_students_subject_data:
        print("Error: No student data was successfully extracted.")
        return None, None

    df_subject_wise = pd.DataFrame(all_students_subject_data)
    mark_columns = ['CIE', 'TW', 'ESEx', 'PR']
    for col in mark_columns:
        df_subject_wise[col] = pd.to_numeric(df_subject_wise[col], errors='coerce').fillna('N/A').apply(
            lambda x: str(int(x)) if isinstance(x, (float, int)) and x == int(x) else str(x)
        )

    # Aggregate total marks per student
    overall_results = df_subject_wise.groupby(['Name', 'PRN'])['Subject Total Marks'].sum().reset_index()
    overall_results.rename(columns={'Subject Total Marks': 'Grand Total Marks'}, inplace=True)
    overall_results['Percentage'] = (overall_results['Grand Total Marks'] / max_marks) * 100

    # Sort and rank
    final_ranking = overall_results.sort_values(by='Grand Total Marks', ascending=False).reset_index(drop=True)
    final_ranking.index = final_ranking.index + 1
    final_ranking.insert(0, 'Rank', final_ranking.index)
    final_ranking['Grand Total Marks'] = final_ranking['Grand Total Marks'].astype(int)
    final_ranking['Percentage'] = final_ranking['Percentage'].map('{:.2f}%'.format)

    return df_subject_wise, final_ranking[['Rank', 'Name', 'PRN', 'Grand Total Marks', 'Percentage']]


# --- 4. COLUMN-WISE OUTPUT FUNCTION (WITH ALL STUDENTS) ---
def make_column_wise_result_analysis_from_pdf(folder_path, max_marks):
    """Process all PDF files in the folder and return column-wise subject-wise DataFrame with all students."""
    pdf_files = glob.glob(os.path.join(folder_path, '*.pdf'))
    if not pdf_files:
        print(f"Error: No PDF files found in {folder_path}.")
        return None

    all_students_data = []

    for pdf_path in pdf_files:
        student_text = read_pdf_text(pdf_path)
        if student_text:
            extracted_data = extract_student_data_final(student_text)
            all_students_data.extend(extracted_data)

    if not all_students_data:
        print("Error: No student data was successfully extracted.")
        return None

    df_subject_wise = pd.DataFrame(all_students_data)
    mark_columns = ['CIE', 'TW', 'ESEx', 'PR']

    for col in mark_columns:
        df_subject_wise[col] = pd.to_numeric(df_subject_wise[col], errors='coerce').fillna('N/A').apply(
            lambda x: str(int(x)) if isinstance(x, (float, int)) and x == int(x) else str(x)
        )

    # Pivot the data to get one row per student with all subject marks as columns
    pivot_df = df_subject_wise.pivot_table(
        index=['Name', 'PRN'],
        columns=['Subject'],
        values=mark_columns,
        aggfunc='first'
    )

    # Flatten the multi-level columns
    pivot_df.columns = [f"{subject}_{mark_col}" for mark_col, subject in pivot_df.columns]
    
    # Reset index to make Name and PRN regular columns
    result_df = pivot_df.reset_index()

    return result_df


# --- 5. EXECUTION BLOCK ---
if __name__ == '__main__':
    print("This module is intended to be used by the web application.")
    print("Call make_result_analysis_and_rank(folder_path, MAX_MARKS) for row-wise output.")
    print("Call make_column_wise_result_analysis_from_pdf(folder_path, MAX_MARKS) for column-wise output with all students.")