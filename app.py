from flask import Flask, render_template, request, send_file, jsonify, session, redirect, url_for
import os
from werkzeug.utils import secure_filename
import pandas as pd
import resultss
import io
from functools import wraps
from datetime import timedelta
from minor_degree_allocation import MinorAllocationSystem


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['SECRET_KEY'] = 'your-secret-key-change-this-in-production'  # Change this!
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)


# Create uploads folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


# Simple user database (replace with real database in production)
USERS = {
    'faculty': {
        'f001': {'password': 'faculty123', 'name': 'siddhi'},
        'f002': {'password': 'faculty456', 'name': 'prajyot'},
    },
    'student': {
        'S12345': {'password': 'student123', 'name': 'krutika'},
        'S12346': {'password': 'student456', 'name': 'sayali'},
        'S12347': {'password': 'student789', 'name': 'sakshi'}
    }
}


# Decorator to check login
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


# Decorator to check faculty role
def faculty_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session or session.get('role') != 'faculty':
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        role = request.form.get('role')
        userid = request.form.get('userid')
        password = request.form.get('password')

        if userid in USERS['faculty'] and USERS['faculty'][userid]['password'] == password:
            session.permanent = True
            session['user_id'] = userid
            session['role'] = 'faculty'
            session['name'] = USERS['faculty'][userid]['name']
            return redirect(url_for('home'))
        else:
            return render_template('login.html', error='Invalid credentials')
    
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/')
def home():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    return render_template('index.html')


@app.route('/result-analysis', methods=['GET', 'POST'])
@login_required
def result_analysis():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file uploaded'
        file = request.files['file']
        if file.filename == '':
            return 'No file selected'
        
        if file:
            filename = secure_filename(file.filename) if file.filename else 'result_analysis.xlsx'
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Process all PDF files in the upload folder for row-wise and column-wise tables
            df_subject_wise, df_ranking = resultss.make_result_analysis_and_rank(app.config['UPLOAD_FOLDER'], resultss.MAX_MARKS)
            df_column_wise = resultss.make_column_wise_result_analysis_from_pdf(app.config['UPLOAD_FOLDER'], resultss.MAX_MARKS)

            # Check if processing was successful
            if df_subject_wise is None or df_ranking is None or df_column_wise is None:
                return render_template('result_analysis.html', 
                                     error="No data could be extracted from the uploaded file. Please check the file format.")

            # Convert DataFrames to HTML table with Bootstrap classes
            subject_table = df_column_wise.to_html(classes=['table', 'table-striped', 'table-hover'], 
                                                index=False, escape=False)
            ranking_table = df_ranking.to_html(classes=['table', 'table-striped', 'table-hover'],
                                            index=False, escape=False)

            # Store data in session for separate display route
            session['subject_table'] = subject_table
            session['ranking_table'] = ranking_table
            session['filename'] = filename
            return redirect(url_for('results'))

    return render_template('result_analysis.html')


@app.route('/results')
@login_required
def results():
    """Display the result analysis data"""
    subject_table = session.get('subject_table')
    ranking_table = session.get('ranking_table')
    filename = session.get('filename')
    
    if not subject_table or not ranking_table:
        return redirect(url_for('result_analysis'))
    
    return render_template('result_display.html', 
                         subject_table=subject_table,
                         ranking_table=ranking_table,
                         filename=filename)


@app.route('/download-result-analysis')
@login_required
def download_result_analysis():
    # Process all PDF files in the upload folder for the column-wise result
    df_subject_wise, df_ranking = resultss.make_result_analysis_and_rank(app.config['UPLOAD_FOLDER'], resultss.MAX_MARKS)
    df_column_wise = resultss.make_column_wise_result_analysis_from_pdf(app.config['UPLOAD_FOLDER'], resultss.MAX_MARKS)

    # Check if processing was successful
    if df_subject_wise is None or df_ranking is None or df_column_wise is None:
        return 'No data available for download. Please upload analysis data first.', 404

    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_column_wise.to_excel(writer, sheet_name='Subject Details', index=False)
        df_ranking.to_excel(writer, sheet_name='Overall Ranking', index=False)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='result_analysis.xlsx'
    )


# ===== MINOR DEGREE ALLOCATION ROUTES =====

@app.route('/faculty/minor-allocation', methods=['GET', 'POST'])
@faculty_required
def faculty_minor_allocation():
    if request.method == 'POST':
        if 'preferences_file' not in request.files or 'marks_file' not in request.files:
            return render_template('mdm_allocation.html', error='Both preferences and marks files are required')
        
        pref_file = request.files['preferences_file']
        marks_file = request.files['marks_file']
        
        if pref_file.filename == '' or marks_file.filename == '':
            return render_template('mdm_allocation.html', error='Both files must be selected')
        
        if pref_file and marks_file:
            try:
                pref_filename = secure_filename(pref_file.filename)
                marks_filename = secure_filename(marks_file.filename)
                pref_filepath = os.path.join(app.config['UPLOAD_FOLDER'], pref_filename)
                marks_filepath = os.path.join(app.config['UPLOAD_FOLDER'], marks_filename)
                
                pref_file.save(pref_filepath)
                marks_file.save(marks_filepath)
                
                allocator = MinorAllocationSystem(
                    preferences_file=pref_filepath,
                    marks_file=marks_filepath,
                    max_marks=1600
                )
                allocator.allocate_students()
                
                # Prepare allocation data for template
                allocation_data = []
                for alloc in allocator.allocations:
                    allocation_data.append({
                        'prn': alloc['PRN'],
                        'name': alloc['Name'],
                        'marks': alloc['Marks'],
                        'percentage': alloc['Percentage'],
                        'current_dept': alloc['Current_Dept'],
                        'minor_degree': alloc['Allocated_Minor'],
                        'preference_used': alloc['Preference_Used'],
                        'status': alloc['Status']
                    })
                
                # Prepare waiting list data
                waiting_list_data = []
                for wait in allocator.waiting_list:
                    waiting_list_data.append({
                        'prn': wait['PRN'],
                        'name': wait['Name'],
                        'marks': wait['Marks'],
                        'percentage': wait['Percentage'],
                        'current_dept': wait['Current_Dept'],
                        'reason': wait['Reason'],
                        'status': wait['Status']
                    })
                
                # Calculate summary statistics
                summary_stats = {
                    'total_processed': len(allocation_data) + len(waiting_list_data),
                    'allocated': len(allocation_data),
                    'waitlisted': len(waiting_list_data),
                    'seats': allocator.available_seats
                }
                
                return render_template('mdm_allocation.html',
                                     allocation_table=True,
                                     allocation_data=allocation_data,
                                     waiting_list_data=waiting_list_data,
                                     summary_stats=summary_stats,
                                     success_message=f'Successfully processed {summary_stats["total_processed"]} students')
            
            except Exception as e:
                return render_template('mdm_allocation.html', error=f'Error processing files: {str(e)}')
    
    return render_template('mdm_allocation.html')


@app.route('/download-minor-allocation')
@faculty_required
def download_minor_allocation():
    # Get latest Excel files from uploads
    xlsx_files = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.endswith('.xlsx')]
    if len(xlsx_files) < 2:
        return 'No data available for download. Upload both preferences and marks files first.', 404
    
    # Get the two most recently modified files
    sorted_files = sorted(xlsx_files, key=lambda x: os.path.getctime(os.path.join(app.config['UPLOAD_FOLDER'], x)), reverse=True)
    pref_file = marks_file = None
    
    for f in sorted_files:
        if 'pref' in f.lower() or 'preference' in f.lower():
            pref_file = f
        elif 'mark' in f.lower():
            marks_file = f
        if pref_file and marks_file:
            break
    
    if not pref_file or not marks_file:
        # Try to use last two files
        pref_file = sorted_files[0]
        marks_file = sorted_files[1] if len(sorted_files) > 1 else sorted_files[0]
    
    try:
        pref_filepath = os.path.join(app.config['UPLOAD_FOLDER'], pref_file)
        marks_filepath = os.path.join(app.config['UPLOAD_FOLDER'], marks_file)
        
        # Create allocator and run
        allocator = MinorAllocationSystem(
            preferences_file=pref_filepath,
            marks_file=marks_filepath,
            max_marks=1500
        )
        allocator.allocate_students()
        
        # Generate report to file
        output_filename = 'Minor_Degree_Allocation_Result.xlsx'
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        allocator.generate_report(output_path)
        
        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='minor_degree_allocation.xlsx'
        )
    except Exception as e:
        return f'Error generating download: {str(e)}', 500


if __name__ == '__main__':
    app.run(debug=True)