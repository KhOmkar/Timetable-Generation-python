import pandas as pd
import re
from openpyxl import load_workbook



def read_timetable(file_path):
    timetable_df = pd.read_excel(file_path, skiprows=6, nrows=25)
    return timetable_df

def extract_schedule_info(timetable_df):
    classroom_data = {}
    faculty_data = {}
    time_slots = timetable_df.columns[1:]
    all_days = timetable_df[timetable_df.columns[0]].unique()

    for index, row in timetable_df.iterrows():
        day = row[timetable_df.columns[0]]
        for col_idx, col in enumerate(time_slots):
            cell_content = row[col]
            if pd.notna(cell_content) and isinstance(cell_content, str):
                time_slot = col
                
                # Handle both practical and theoretical classes
                if '-' in cell_content:
                    entries = cell_content.split(' - ')
                    for entry in entries:
                        lines = entry.split('\n')
                        if len(lines) >= 3:
                            sub_division = lines[0].strip()
                            subject_info = re.search(r'(\w+)\s*\((.*?)\)', lines[1])
                            classroom_info = re.search(r'\((.*?)\)', lines[2])

                            if subject_info and classroom_info:
                                subject = subject_info.group(1).strip()
                                teacher = subject_info.group(2).strip()
                                classroom = classroom_info.group(1).strip()

                                # Store classroom data
                                if classroom not in classroom_data:
                                    classroom_data[classroom] = {}
                                if day not in classroom_data[classroom]:
                                    classroom_data[classroom][day] = {}
                                classroom_data[classroom][day][time_slot] = {
                                    'Sub-Division': sub_division,
                                    'Subject': subject,
                                    'Teacher': teacher
                                }

                                # Store faculty data
                                if teacher not in faculty_data:
                                    faculty_data[teacher] = {}
                                if day not in faculty_data[teacher]:
                                    faculty_data[teacher][day] = {}
                                faculty_data[teacher][day][time_slot] = {
                                    'Sub-Division': sub_division,
                                    'Subject': subject,
                                    'Classroom': classroom
                                }
                else:
                    lines = cell_content.split('\n')
                    if len(lines) >= 3:
                        subject = lines[0].strip()
                        teacher = lines[1].strip()
                        classroom = lines[2].strip()

                        # Store classroom data
                        if classroom not in classroom_data:
                            classroom_data[classroom] = {}
                        if day not in classroom_data[classroom]:
                            classroom_data[classroom][day] = {}
                        classroom_data[classroom][day][time_slot] = {
                            'Sub-Division': '',
                            'Subject': subject,
                            'Teacher': teacher
                        }

                        # Store faculty data
                        if teacher not in faculty_data:
                            faculty_data[teacher] = {}
                        if day not in faculty_data[teacher]:
                            faculty_data[teacher][day] = {}
                        faculty_data[teacher][day][time_slot] = {
                            'Sub-Division': '',
                            'Subject': subject,
                            'Classroom': classroom
                        }

    return classroom_data, faculty_data, all_days, time_slots

def save_schedules_to_excel(classroom_data, faculty_data, output_file, all_days, time_slots):
    with pd.ExcelWriter(output_file) as writer:
        # Save classroom schedules
        for classroom, timetable in classroom_data.items():
            df = pd.DataFrame(index=all_days, columns=time_slots)
            for day in all_days:
                if day in timetable:
                    for time, details in timetable[day].items():
                        df.loc[day, time] = f"{details['Subject']} ({details['Teacher']})" \
                                          + (f"\n{details['Sub-Division']}" if details['Sub-Division'] else "")
                else:
                    df.loc[day] = [""] * len(time_slots)
            df.to_excel(writer, sheet_name=f"CR_{classroom[:28]}")

        # Save faculty schedules
        for faculty, timetable in faculty_data.items():
            df = pd.DataFrame(index=all_days, columns=time_slots)
            for day in all_days:
                if day in timetable:
                    for time, details in timetable[day].items():
                        df.loc[day, time] = f"{details['Subject']}\n{details['Classroom']}" \
                                          + (f"\n{details['Sub-Division']}" if details['Sub-Division'] else "")
                else:
                    df.loc[day] = [""] * len(time_slots)
            df.to_excel(writer, sheet_name=f"FAC_{faculty[:28]}")

            

def main():
    file_path = "D:\\Classwise 24 25 Sem I 05.xlsm"
    output_file = "C:\\Users\\omkar\\Downloads\\timetable\\time.xlsx"

    timetable_df = read_timetable(file_path)
    classroom_data, faculty_data, all_days, time_slots = extract_schedule_info(timetable_df)
    save_schedules_to_excel(classroom_data, faculty_data, output_file, all_days, time_slots)

    print(f"Classroom and Faculty timetables have been generated and saved to {output_file}.")

if __name__ == "__main__":
    main()