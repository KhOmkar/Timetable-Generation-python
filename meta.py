import pandas as pd
import re
import os

def extract_course_teacher_data(excel_path):
    # Read the Excel file
    xls = pd.ExcelFile(excel_path)
    
    # Initialize an empty list to store all records
    all_data = []
    
    # Process each sheet (division)
    for sheet_name in xls.sheet_names:
        print(f"Processing sheet: {sheet_name}")
        
        # Read the sheet
        df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=33, usecols=[0, 1, 2])
        
        # Convert column names to string to handle any non-string column names
        df.columns = df.columns.astype(str)
        
        # Look for rows where the first column contains "Course Code"
        header_rows = df[df.iloc[:, 0].astype(str).str.contains("Course Code", case=False, na=False)].index
        
        if len(header_rows) == 0:
            # If no "Course Code" found, try looking for rows with course codes matching patterns
            potential_rows = df[df.iloc[:, 0].astype(str).str.match(r'^\d{7}|^CS\d{3}', na=False)].index
            if len(potential_rows) > 0:
                # Process these rows directly
                course_data = df.iloc[potential_rows]
                course_col = 0
                name_col = 1
                teacher_col = 2
                # Make sure we have at least 3 columns
                if len(df.columns) >= 3:
                    for _, row in course_data.iterrows():
                        process_row(row, sheet_name, course_col, name_col, teacher_col, all_data)
            continue
            
        for header_row in header_rows:
            # Find the column indices for our data
            header = df.iloc[header_row]
            
            course_col = None
            name_col = None
            teacher_col = None
            
            for i, col_name in enumerate(header):
                col_name = str(col_name).lower()
                if "course code" in col_name:
                    course_col = i
                elif "course name" in col_name:
                    name_col = i
                elif "teacher" in col_name:
                    teacher_col = i
            
            if course_col is not None and name_col is not None and teacher_col is not None:
                # Process rows below this header
                data_rows = df.iloc[header_row+1:].reset_index(drop=True)
                
                for _, row in data_rows.iterrows():
                    course_code = str(row.iloc[course_col]).strip() if pd.notna(row.iloc[course_col]) else ""
                    # Stop at the first empty row
                    if not course_code or course_code == "nan":
                        continue
                    
                    # Use direct indices rather than column names
                    process_row(row, sheet_name, course_col, name_col, teacher_col, all_data)
    
    # Create a DataFrame from all collected data
    result_df = pd.DataFrame(all_data)
    return result_df

def process_row(row, sheet_name, course_col, name_col, teacher_col, all_data):
    course_code = str(row.iloc[course_col]).strip() if pd.notna(row.iloc[course_col]) else ""
    course_full = str(row.iloc[name_col]).strip() if pd.notna(row.iloc[name_col]) else ""
    teachers_text = str(row.iloc[teacher_col]).strip() if pd.notna(row.iloc[teacher_col]) else ""
    
    # Skip empty rows
    if course_code == "" or course_code == "nan" or course_full == "" or course_full == "nan" or teachers_text == "" or teachers_text == "nan":
        return
    
    # Check for course codes that match our expected pattern
    if not re.match(r'^\d{6,7}|^CS\d{3}', course_code):
        return
    
    # Extract course short form
    course_short_match = re.search(r'\(([^)]+)\)', course_full)
    course_short = course_short_match.group(1) if course_short_match else ""
    course_full = re.sub(r'\s*\([^)]*\)', '', course_full).strip()
    
    # Handle multiple teachers in one cell (split by newlines or commas)
    teachers = []
    if '\n' in teachers_text:
        teachers = [t.strip() for t in teachers_text.split('\n') if t.strip()]
    else:
        teachers = [teachers_text]
    
    for teacher in teachers:
        # Extract teacher full name and initials
        teacher_short_match = re.search(r'\(([^)]+)\)', teacher)
        teacher_short = teacher_short_match.group(1) if teacher_short_match else ""
        teacher_full = re.sub(r'\s*\([^)]*\)', '', teacher).strip()
        
        # Add to our data collection
        all_data.append({
            'Division': sheet_name,
            'Course_Code': course_code,
            'Course_Name_Full': course_full,
            'Course_Name_Short': course_short,
            'Teacher_Name_Full': teacher_full,
            'Teacher_Name_Short': teacher_short
        })

def main():
    # Path to your Excel file
    excel_path = input("Enter the path to your Excel file: ")
    
    if not os.path.exists(excel_path):
        print(f"Error: File '{excel_path}' not found.")
        return
    
    try:
        # Extract the data
        result = extract_course_teacher_data(excel_path)
        
        if result.empty:
            print("No course-teacher data found in the Excel file.")
            return
        
        # Define output path
        output_path = os.path.splitext(excel_path)[0] + "_course_teacher_data.csv"
        
        # Save to CSV
        result.to_csv(output_path, index=False)
        print(f"Data successfully extracted and saved to {output_path}")
        
        # Display summary
        print("\nSummary:")
        print(f"Total records: {len(result)}")
        print(f"Unique divisions: {result['Division'].nunique()}")
        print(f"Unique courses: {result['Course_Name_Full'].nunique()}")
        print(f"Unique teachers: {result['Teacher_Name_Full'].nunique()}")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()