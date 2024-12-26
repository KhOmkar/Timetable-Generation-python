import pandas as pd
import os

def generate_faculty_and_classroom_timetables(student_timetable_path, output_directory):
    """
    Generate faculty and classroom timetables from a student timetable.

    Parameters:
        student_timetable_path (str): Path to the student timetable Excel file.
        output_directory (str): Directory where the generated timetables will be saved.

    Raises:
        ValueError: If the input file does not contain required columns.
    """
    # Ensure the output directory exists
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Load the student timetable from Excel
    student_timetable = pd.read_excel(student_timetable_path)

    # Validate required columns
    required_columns = {'Day', 'Time', 'Subject', 'Faculty', 'Classroom'}
    if not required_columns.issubset(student_timetable.columns):
        raise ValueError(f"The input Excel file must contain the following columns: {required_columns}")

    # Generate faculty timetable
    faculty_timetable = student_timetable.groupby(['Faculty', 'Day', 'Time'])\
        .apply(lambda x: ', '.join(x['Subject'])).reset_index()
    faculty_timetable.columns = ['Faculty', 'Day', 'Time', 'Subjects']

    # Generate classroom timetable
    classroom_timetable = student_timetable.groupby(['Classroom', 'Day', 'Time'])\
        .apply(lambda x: ', '.join(x['Subject'])).reset_index()
    classroom_timetable.columns = ['Classroom', 'Day', 'Time', 'Subjects']

    # Save the generated timetables to Excel files
    faculty_timetable_path = os.path.join(output_directory, "faculty_timetable.xlsx")
    classroom_timetable_path = os.path.join(output_directory, "classroom_timetable.xlsx")

    with pd.ExcelWriter(faculty_timetable_path) as writer:
        faculty_timetable.to_excel(writer, index=False, sheet_name='Faculty Timetable')

    with pd.ExcelWriter(classroom_timetable_path) as writer:
        classroom_timetable.to_excel(writer, index=False, sheet_name='Classroom Timetable')

    print(f"Faculty timetable saved to: {faculty_timetable_path}")
    print(f"Classroom timetable saved to: {classroom_timetable_path}")

# Example usage:
if _name_ == "_main_":
    # Provide the path to the student timetable Excel file and an output directory.
    student_timetable_path = r"C:\\Users\\Karan Doifode\\Desktop\\timetable\\classroom.xlsm"  # Replace with your file path
    output_directory = r"output"  # Replace with your desired output directory

    generate_faculty_and_classroom_timetables(student_timetable_path, output_directory)