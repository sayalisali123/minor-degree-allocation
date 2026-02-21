import pandas as pd

class MinorAllocationSystem:
    
    def __init__(self, preferences_file, marks_file, max_marks=1600):
        
       
        try:
         
            self.preferences_df = pd.read_excel(preferences_file)
            self.marks_df = pd.read_excel(marks_file)

            # Strip leading/trailing spaces from all column names to handle inconsistencies
            self.preferences_df.columns = self.preferences_df.columns.str.strip()
            self.marks_df.columns = self.marks_df.columns.str.strip()

        except Exception as e:
            raise ValueError(f"Failed to read file: {e}")
        
        # --- 2. Configuration and Criteria ---
        self.max_marks = max_marks

        self.targets = {
            'CSE-AIML': 60,
            'CSE': 120,
            'ETC': 60,
            'MECH': 60,
            'CIVIL': 60
        }

        self.criteria = {
            'CSE-AIML': 70,
            'CSE': 60,
            'ETC': 54,
            'MECH': 50,
            'CIVIL': 40
        }

        self.available_seats = self.targets.copy()
        self.allocations = []
        self.waiting_list = []
        
        # --- 3. Column Standardization (to ensure merge compatibility) ---
        self.name_column = self._detect_name_column()
        self._standardize_columns()

    def _detect_and_rename_column(self, df, aliases, target_name):
        """Helper to find a column by alias and rename it to the target_name."""
        found_col = None
        for alias in aliases:
            if alias in df.columns:
                found_col = alias
                break
        if found_col and found_col != target_name:
            df.rename(columns={found_col: target_name}, inplace=True)
            return target_name
        return found_col

    def _detect_name_column(self):
        """Detects the student name column."""
        possible_name_headers = [
            'FULL NAME OF THE STUDENT IN CAPITAL (STARTING WITH SURNAME)',
            'FULL NAME OF THE STUDENT',
            'FULL NAME OF THE STUDENT ',
            'Full Name',
            'Name',
            'Student Name',
            'Student Full Name',
            'STUDENT NAME',
            'NAME',
            'FULL_NAME',
            'STUDENT_FULL_NAME'
        ]
        for col in possible_name_headers:
            if col in self.preferences_df.columns:
                return col
        raise ValueError(f"Student name column not found in preferences file. Available columns: {list(self.preferences_df.columns)}")

    def _standardize_columns(self):
        """Standardizes PRN and Total Marks columns."""
        prn_aliases = ['PRN', 'PRN No', 'PRN No.', 'prn', 'Roll', 'Roll No']
        self._detect_and_rename_column(self.marks_df, prn_aliases, 'PRN')
        self._detect_and_rename_column(self.preferences_df, prn_aliases, 'PRN')
        self.marks_df['PRN'] = self.marks_df['PRN'].astype(str)
        self.preferences_df['PRN'] = self.preferences_df['PRN'].astype(str)
        
        marks_aliases = ['Total Marks', 'Total', 'Marks', 'TOTAL']
        self._detect_and_rename_column(self.marks_df, marks_aliases, 'Total Marks')

        # Ensure 'EXISTING DEPARTMENT' column exists in preferences
        dept_aliases = ['EXISTING DEPARTMENT', 'Existing Department', 'Department', 'DEPT']
        self._detect_and_rename_column(self.preferences_df, dept_aliases, 'EXISTING DEPARTMENT')

    def calculate_percentage(self, marks):
        """Calculates the percentage based on Max Marks (1600)."""
        return (marks / self.max_marks) * 100

    def _extract_preference_number(self, value):
        """Extracts the integer preference number (1, 2, 3, etc.) from the preference column."""
        if isinstance(value, (str, int, float)) and pd.notna(value):
            value = str(value).upper().strip()
            if 'PREFERENCE' in value:
                try:
                    parts = value.split('PREFERENCE')
                    if len(parts) > 1:
                        return int(parts[1].strip())
                except:
                    return 0
        return 0

    def get_student_preferences(self, row):
        """Extracts student preferences into a sorted dictionary, ignoring 'PREFERENCE 0'."""
        mapping = {
            'ETC': 'ETC', 'MECH': 'MECH', 'CIVIL': 'CIVIL',
            'CSE (AIML)': 'CSE-AIML', 'MDM Preference Choices [CSE (AIML)]': 'CSE-AIML', 
            'CSE': 'CSE', 'MDM Preference Choices [CSE]': 'CSE',
            ' CSE': 'CSE',  # Handle stripped spaces if needed
            ' CSE (AIML)': 'CSE-AIML'
        }
        prefs = {}
        for col, minor_code in mapping.items():
            if col in self.preferences_df.columns:
                pref_value = row.get(col)
                pref_num = self._extract_preference_number(pref_value)
                # Only preferences > 0 (1, 2, 3, 4) are considered. 
                # This ensures self-department preferences (marked 0) are correctly ignored.
                if pref_num > 0:
                    prefs[pref_num] = minor_code
        return prefs

    def allocate_students(self):
        """Performs the main allocation process."""
        
        # --- 1. Data Merge and Merit Sorting ---
        merged = self.preferences_df.merge(
            self.marks_df[['PRN', 'Total Marks']], on='PRN', how='inner'
        )
        
        merged['Total Marks'] = pd.to_numeric(merged['Total Marks'], errors='coerce')
        merged = merged.dropna(subset=['Total Marks'])
        
        merged['Percentage'] = merged['Total Marks'].apply(self.calculate_percentage)
        merged = merged.sort_values('Total Marks', ascending=False).reset_index(drop=True)
        
        # --- 2. Allocation Loop (Merit > Preference) ---
        for _, row in merged.iterrows():
            prn = row['PRN']
            name = row[self.name_column]
            dept = row['EXISTING DEPARTMENT']
            marks = row['Total Marks']
            perc = row['Percentage']

            prefs = self.get_student_preferences(row)
            allocated = False

            # Check preferences in priority order (1st, 2nd, 3rd...)
            for pref_num in sorted(prefs.keys()):
                minor = prefs[pref_num]
                
                if minor not in self.criteria: continue
                threshold = self.criteria[minor]
                
                # Check 1: Merit (Percentage >= Criteria) AND Check 2: Availability (Seat > 0)
                if perc >= threshold and self.available_seats.get(minor, 0) > 0:
                    self.allocations.append({
                        'PRN': prn, 'Name': name, 'Current_Dept': dept, 
                        'Marks': marks, 'Percentage': f"{perc:.2f}",
                        'Allocated_Minor': minor, 'Preference_Used': f"PREFERENCE {pref_num}",
                        'Status': 'ALLOCATED'
                    })
                    self.available_seats[minor] -= 1
                    allocated = True
                    break # Stop checking preferences for this student

            # --- 3. Waiting List Assignment ---
            if not allocated:
                self.waiting_list.append({
                    'PRN': prn, 'Name': name, 'Current_Dept': dept, 
                    'Marks': marks, 'Percentage': f"{perc:.2f}",
                    'Allocated_Minor': 'WAITLIST', 
                    'Reason': self._determine_waitlist_reason(perc, prefs),
                    'Status': 'WAITING'
                })

    def _determine_waitlist_reason(self, percentage, preferences):
        """Determines the specific reason a student was not allocated."""
        for pref_num in sorted(preferences.keys()):
            minor = preferences[pref_num]
            if minor in self.criteria:
                threshold = self.criteria[minor]
                
                if percentage < threshold: return f"Failed {minor} criteria ({threshold}%)"
                if self.available_seats.get(minor, 0) == 0: return f"Seats full for {minor}"
                    
        return "No valid preference or below all required criteria"

    def generate_report(self, output_file='Minor_Degree_Allocation_Result_Final.xlsx'):
        """Generates the final Excel report with three sheets."""
        # ... (Report generation logic remains the same as in the previous output) ...
        df_alloc = pd.DataFrame(self.allocations)
        df_wait = pd.DataFrame(self.waiting_list)
        
        cut_offs = {}
        if not df_alloc.empty:
            df_alloc['Percentage_Float'] = df_alloc['Percentage'].astype(float)
            cut_off_df = df_alloc.groupby('Allocated_Minor')['Percentage_Float'].min().reset_index()
            cut_offs = cut_off_df.set_index('Allocated_Minor')['Percentage_Float'].to_dict()

        summary_data = []
        for minor, target in self.targets.items():
            allocated = target - self.available_seats[minor]
            cut_off_val = cut_offs.get(minor)
            cut_off_str = f"{cut_off_val:.2f}%" if cut_off_val is not None else 'N/A'

            summary_data.append({
                'Minor Degree': minor, 'Target Seats': target, 
                'Required %': f"{self.criteria[minor]}%", 'Allocated': allocated, 
                'Available': self.available_seats[minor],
                'Occupancy %': round((allocated / target) * 100 if target > 0 else 0, 2),
                'Cut-Off %': cut_off_str
            })
        df_summary = pd.DataFrame(summary_data)

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_alloc.to_excel(writer, sheet_name='Allocations', index=False)
            df_wait.to_excel(writer, sheet_name='Waiting_List', index=False)
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"Report successfully generated and saved to '{output_file}'")
    
    def print_summary(self):
        """Prints a concise summary of the allocation results to the console."""
        
        # 1. Print overall stats
        print(f"\n========== MINOR DEGREE ALLOCATION SUMMARY (Max Marks: {self.max_marks}) ==========")
        total_alloc = len(self.allocations)
        total_wait = len(self.waiting_list)
        print(f"Total Allocated Students : {total_alloc}")
        print(f"Total Waiting Students   : {total_wait}\n")
        
        df_alloc_temp = pd.DataFrame(self.allocations)
        
        # 2. Calculate Cut-Offs
        cut_offs = {}
        if not df_alloc_temp.empty:
            df_alloc_temp['Percentage_Float'] = df_alloc_temp['Percentage'].astype(float)
            cut_off_df = df_alloc_temp.groupby('Allocated_Minor')['Percentage_Float'].min().reset_index()
            cut_offs = cut_off_df.set_index('Allocated_Minor')['Percentage_Float'].to_dict()
        
        # 3. Print detailed summary table
        print(f"{'Minor':<12} {'Target':<6} {'Required':<10} {'Allocated':<9} {'Remaining':<9} {'Cut-Off':<9}")
        print("-" * 65)
        
        for minor, target in self.targets.items():
            allocated = target - self.available_seats[minor]
            remaining = self.available_seats[minor]
            
            cut_off_val = cut_offs.get(minor)
            # Format cut-off to 1 decimal place or 'N/A'
            cut_off = f"{cut_off_val:.1f}%" if cut_off_val is not None else 'N/A'
            required_perc = f"{self.criteria[minor]}%"
            
            print(f"{minor:<12} {target:<6} {required_perc:<10} {allocated:<9} {remaining:<9} {cut_off:<9}")
        print("=========================================================================================\n")


# Example Execution Block (to be run locally)
if __name__ == "__main__":
    
  
    preferences_file = 'preferences.xlsx - Sheet1.csv'
    marks_file = 'student_marks_output.xlsx - Sheet1.csv'
    output_file = 'Minor_Degree_Allocation_Result_Final.xlsx'

    try:
        system = MinorAllocationSystem(preferences_file, marks_file, max_marks=1600)
        system.allocate_students()
        system.generate_report(output_file)
        system.print_summary()

    except Exception as e:
        print(f"An error occurred during execution: {e}")