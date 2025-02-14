import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

def create_timetable_structure():
    time_slots = [
        '8:30 to 9:25', '9:25 to 10:20', '10:20 to 10:30', '10:30 to 11:25',
        '11:25 to 12:20', '12:20 to 13:15', '13:15 to 14:10', '14:10 to 15:05',
        '15:05 to 15:10', '15:10 to 16:00', '16:00 to 16:50', '16:50 to 16:55',
        '16:55 to 17:45', '17:45 to 18:25'
    ]
    days = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT']
    df = pd.DataFrame(index=days, columns=time_slots)
    return df.fillna('')

def extract_teacher_schedule(input_file, teacher_short="MNV"):
    raw_timetable = pd.read_excel(input_file, skiprows=6, nrows=25)
    teacher_schedule = create_timetable_structure()
    time_slots = teacher_schedule.columns.tolist()
    
    for index, row in raw_timetable.iterrows():
        day = row.iloc[0]
        if not isinstance(day, str) or day.strip() not in teacher_schedule.index:
            continue
        day = day.strip()
        
        for col_idx, time_slot in enumerate(time_slots):
            current_cell = row.iloc[col_idx + 1]
            if pd.isna(current_cell) or not isinstance(current_cell, str):
                continue
            
            if teacher_short in current_cell:
                teacher_schedule.at[day, time_slot] = current_cell.strip()
    
    return teacher_schedule

def save_teacher_schedule(schedule_df, output_file, teacher_short="MNV"):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        schedule_df.to_excel(writer, sheet_name=f'Schedule_{teacher_short}')
        workbook = writer.book
        worksheet = writer.sheets[f'Schedule_{teacher_short}']
        
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)
        
        for row in worksheet.rows:
            max_lines = max([str(cell.value).count('\n') + 1 if cell.value else 1 for cell in row])
            worksheet.row_dimensions[row[0].row].height = max_lines * 15

def main():
    input_file = "D:\\Classwise 24 25 Sem I 05.xlsm"
    output_file = "C:\\Users\\omkar\\Downloads\\timetable\\mnv6.xlsx"
    
    try:
        print(f"Extracting schedule for teacher MNV...")
        teacher_schedule = extract_teacher_schedule(input_file, "MNV")
        print("Saving schedule to Excel...")
        save_teacher_schedule(teacher_schedule, output_file)
        print(f"Schedule has been generated and saved to {output_file}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    
if __name__ == "__main__":
    main()
