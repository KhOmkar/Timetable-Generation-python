import pandas as pd
import re
from openpyxl import load_workbook

def read_timetable(file_path):
    # We'll read both the metadata and timetable sections
    # Metadata is around rows 34-42 based on your image
    metadata_df = pd.read_excel(file_path, skiprows=33, nrows=9)
    
    # Main timetable starts after row 6
    timetable_df = pd.read_excel(file_path, skiprows=6, nrows=25)
    
    return metadata_df, timetable_df

def find_teacher_courses(metadata_df, teacher_short="SBK"):
    # This function finds all courses taught by SBK from the metadata
    teacher_courses = []
    
    # Check theory courses (left side of metadata)
    theory_courses = metadata_df.iloc[:, 1:4]  # Columns B-D
    for _, row in theory_courses.iterrows():
        if pd.notna(row[1]) and pd.notna(row[2]):  # Check if course and teacher exist
            teacher_info = str(row[2]).strip()
            if f"({teacher_short})" in teacher_info:
                course_info = str(row[1]).strip()
                # Extract course short form from within parentheses
                course_match = re.search(r'\((.*?)\)', course_info)
                if course_match:
                    teacher_courses.append(course_match.group(1))
    
    # Check lab courses (right side of metadata)
    lab_courses = metadata_df.iloc[:, 5:8]  # Columns F-H
    for _, row in lab_courses.iterrows():
        if pd.notna(row[1]) and pd.notna(row[2]):
            teacher_info = str(row[2]).strip()
            if f"({teacher_short})" in teacher_info:
                course_info = str(row[1]).strip()
                course_match = re.search(r'\((.*?)\)', course_info)
                if course_match:
                    teacher_courses.append(course_match.group(1))
    
    return teacher_courses

def extract_teacher_schedule(timetable_df, teacher_courses, teacher_short="SBK"):
    # Initialize dictionary to store the teacher's schedule
    schedule = {}
    time_slots = timetable_df.columns[1:]  # All columns except first (which contains days)
    
    # Process each row (day) in the timetable
    for index, row in timetable_df.iterrows():
        day = row[timetable_df.columns[0]]  # Get the day from first column
        
        # Skip if day is not valid
        if not isinstance(day, str):
            continue
            
        day = day.strip()
        if day not in ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT']:
            continue
        
        # Initialize dictionary for this day
        schedule[day] = {}
        
        # Check each time slot
        for time_slot in time_slots:
            cell_content = row[time_slot]
            
            # Skip empty cells
            if pd.isna(cell_content) or not isinstance(cell_content, str):
                continue
            
            # Check if this cell contains our teacher's short form or any of their courses
            if any(course in cell_content for course in teacher_courses) or teacher_short in cell_content:
                # Split cell content into lines
                lines = cell_content.split('\n')
                
                # Extract relevant information
                subject = ''
                classroom = ''
                division = ''
                
                # Process the lines to extract information
                for line in lines:
                    if any(course in line for course in teacher_courses):
                        subject = line.strip()
                    elif 'H2' in line or 'H3' in line:  # Assuming classroom format
                        classroom = line.strip()
                    elif any(div in line for div in ['C1', 'C2', 'C3', 'C4']):
                        division = line.strip()
                
                # Store the information in our schedule
                if subject or classroom:
                    schedule[day][str(time_slot)] = {
                        'Subject': subject,
                        'Classroom': classroom,
                        'Division': division
                    }
    
    return schedule

def save_teacher_schedule(schedule, output_file, teacher_short="SBK"):
    # Create a DataFrame from the schedule
    # First, get all unique time slots
    all_time_slots = set()
    for day_schedule in schedule.values():
        all_time_slots.update(day_schedule.keys())
    all_time_slots = sorted(list(all_time_slots))
    
    # Create empty DataFrame
    df = pd.DataFrame(index=['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'], columns=all_time_slots)
    
    # Fill in the schedule
    for day in schedule:
        for time_slot in schedule[day]:
            entry = schedule[day][time_slot]
            cell_content = f"{entry['Subject']}\n{entry['Classroom']}"
            if entry['Division']:
                cell_content = f"{entry['Division']}\n{cell_content}"
            df.loc[day, time_slot] = cell_content
    
    # Fill NaN values with empty string
    df = df.fillna('')
    
    # Save to Excel
    df.to_excel(output_file, sheet_name=f'Schedule_{teacher_short}')

def main():
    # File paths
    file_path = "D:\\Classwise 24 25 Sem I 05.xlsm"
    output_file = "C:\\Users\\omkar\\Downloads\\timetable\\SBK_.xlsx"
    
    # Read both metadata and timetable
    metadata_df, timetable_df = read_timetable(file_path)
    print(metadata_df)
    # Find all courses taught by SBK
    teacher_courses = find_teacher_courses(metadata_df)
    print(f"Courses taught by SBK: {teacher_courses}")
    
    # Extract SBK's schedule
    schedule = extract_teacher_schedule(timetable_df, teacher_courses)
    
    # Save the schedule
    save_teacher_schedule(schedule, output_file)
    
    print(f"Schedule for SBK has been generated and saved to {output_file}")

if __name__ == "__main__":
    main()