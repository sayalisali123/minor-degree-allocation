import re
import pandas as pd

def extract_student_data(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Pattern to capture student data blocks
    pattern = re.compile(
        r"^\s*(\d+)\s+(\d{10,12})\s+\(.*?\)\s+(.*?)\s+((?:--|\*\?|\w+)[\s\S]+?)\s+(PASS|FAIL|ATKT|fail|pass|Fail|Fail ATKT\(\d+\)|FAIL ATKT\(\d+\))",
        re.MULTILINE | re.IGNORECASE
    )

    data = []
    for match in pattern.finditer(content):
        prn = match.group(2).strip()
        result = match.group(5).upper()
        # Extract numerical marks from the marks block
        marks_block = match.group(4)
        marks = re.findall(r"\b\d+\b", marks_block)
        total_marks = 0
        if marks:
            total_marks = max(map(int, marks))
        data.append({
            "PRN No": prn,
            "Total Marks": total_marks,
            "Status": result
        })

    df = pd.DataFrame(data)
    return df

# The file handling is now done in the Flask application
if __name__ == "__main__":
    print("This module is now used through the web interface")