# works okish for theoretical classes but need to work on the practical classes data



import pandas as pd
import re

def read_timetable(file_path):
    # Read the timetable from an Excel file, skipping initial rows
    timetable_df = pd.read_excel(file_path, skiprows=6)  # Adjust based on your structure
    return timetable_df

def extract_classroom_info(timetable_df):
    classroom_data = {}
    time_slots = timetable_df.columns[1:]  # The time slot headers (excluding 'Day')

    # Iterate through rows (days) and columns (time slots)
    for index, row in timetable_df.iterrows():
        day = row[timetable_df.columns[0]]  # Extract the day
        for col_idx, col in enumerate(time_slots):
            cell_content = row[col]
            if pd.notna(cell_content) and isinstance(cell_content, str):  # Check if cell is not empty
                time_slot = col  # Corresponding time slot
                # Split by new lines for theoretical classes or hyphen for practical classes
                if '-' in cell_content:
                    entries = cell_content.split(' - ')
                    for entry in entries:
                        lines = entry.split('\n')
                        if len(lines) >= 3:  # Ensure we have enough lines
                            sub_division = lines[0].strip()
                            subject_info = re.search(r'(\w+)\s*\((.*?)\)', lines[1])  # Subject and Teacher
                            classroom_info = re.search(r'\((.*?)\)', lines[2])  # Classroom number

                            if subject_info and classroom_info:
                                subject_short_form = subject_info.group(1).strip()
                                teacher_short_name = subject_info.group(2).strip()
                                classroom_number = classroom_info.group(1).strip()

                                # Store data in dictionary
                                if classroom_number not in classroom_data:
                                    classroom_data[classroom_number] = []
                                
                                classroom_data[classroom_number].append({
                                    'Day': day,
                                    'Time': time_slot,
                                    'Sub-Division': sub_division,
                                    'Subject': subject_short_form,
                                    'Teacher': teacher_short_name
                                })
                else:
                    lines = cell_content.split('\n')
                    if len(lines) >= 3:  # Ensure we have enough lines
                        subject_short_form = lines[0].strip()
                        teacher_short_name = lines[1].strip()
                        classroom_number = lines[2].strip()

                        # Store data in dictionary
                        if classroom_number not in classroom_data:
                            classroom_data[classroom_number] = []
                        
                        classroom_data[classroom_number].append({
                            'Day': day,
                            'Time': time_slot,
                            'Sub-Division': '',
                            'Subject': subject_short_form,
                            'Teacher': teacher_short_name
                        })

    return classroom_data

def save_classroom_data(classroom_data, output_file):
    with pd.ExcelWriter(output_file) as writer:
        for classroom_number, entries in classroom_data.items():
            df_entries = pd.DataFrame(entries)
            df_entries.to_excel(writer, sheet_name=classroom_number[:30], index=False)  # Limit sheet name length

def main():
    file_path = "D:\\Downloads\\Timetable\\omkar\\Classwise 24 25 Sem I 05.xlsm"  # Input file path (Excel format)
    output_file = "C:\\Users\\91774\\Downloads\\excel\\Classroom_Timetables1.xlsx"  # Output file path with a valid file name and extension
    
    timetable_df = read_timetable(file_path)
    
    classroom_data = extract_classroom_info(timetable_df)
    
    save_classroom_data(classroom_data, output_file)
    
    print(f"Classroom information has been extracted and saved to {output_file}.")

if __name__ == "__main__":
    main()
