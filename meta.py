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
        df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=33, usecols=[0, 1, 3, 5],header=None)
        
        # Clean the DataFrame
        df = df.dropna()  # Drop rows with NaN values
        df = df.reset_index(drop = True) # Reset index after dropping rows
        print(df)  

        for col_idx in range(len(df.columns)):  

            df = df.iloc[1:]   

            # Process each row in this section
            for _, row in df.iterrows():  #
                if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]) and pd.notna(row.iloc[2]):
                    course_code = str(row.iloc[0]).strip()   #
                    
                    # Extract course full name and short form
                    course_full = str(row.iloc[1]).strip()
                    course_short_match = re.search(r'\(([^)]+)\)', course_full)
                    course_short = course_short_match.group(1) if course_short_match else ""
                    course_full = re.sub(r'\s*\([^)]*\)', '', course_full).strip()
                    
                    # Handle multiple teachers in one cell
                    teachers_text = str(row.iloc[2])
                    teachers = [t.strip() for t in teachers_text.split('\n') if t.strip()]
                    
                    # Extract classroom information
                    classroom = str(row.iloc[3]).strip()

                    for teacher in teachers:
                        # Extract teacher full name and initials
                        teacher_short_match = re.search(r'\(([^)]+)\)', teacher)
                        teacher_short = teacher_short_match.group(1) if teacher_short_match else ""
                        teacher_full = re.sub(r'\s*\([^)]*\)', '', teacher).strip()
                        
                        # Add to our data collection
                        all_data.append({
                            'Division': sheet_name,
                            'Teacher_Initials': teacher_short,
                            'Course_Initials': course_short,
                            'Course_Code': course_code,
                            'Course_Name': course_full,
                            'Teacher_Name': teacher_full,
                            'Classroom': classroom
                        })
            # Move to next section in the sheet
            break
    
    # Create a DataFrame from all collected data
    result_df = pd.DataFrame(all_data)
    return result_df

def main():
    # Path to your Excel file
    excel_path = "D:\\Classwise 24 25 Sem I.xlsm"
    
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
        output_path = "C:\\Users\\omkar\\Downloads\\timetable\\meta_info_4.csv"
        
        # Save to CSV
        result.to_csv(output_path, index=False)
        print(f"Data successfully extracted and saved to {output_path}")
        
        # Display summary
        print("\nSummary:")
        print(f"Total records: {len(result)}")
        print(f"Unique divisions: {result['Division'].nunique()}")
        print(f"Unique courses: {result['Course_Name'].nunique()}")
        print(f"Unique teachers: {result['Teacher_Name'].nunique()}")
        print(f"Unique classrooms: {result['Classroom'].nunique()}")
        
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()