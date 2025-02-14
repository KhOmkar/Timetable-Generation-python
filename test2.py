import pandas as pd

def read_metadata(file_path):
    # Read the first few rows to get metadata
    metadata = pd.read_excel(file_path, nrows=5)  # Adjust nrows as needed
    school_info = metadata.iloc[0].to_dict()  # Convert first row to dictionary
    return school_info

def read_timetable(file_path):
    # Load the timetable from an Excel file, skipping initial rows
    timetable_df = pd.read_excel(file_path, skiprows=5)  # Adjust skiprows as needed
    return timetable_df

def process_timetable(timetable_df):
    processed_timetable = {}

    # Iterate through rows (days) and columns (time slots)
    for index, row in timetable_df.iterrows():
        day = row['Day']  # Assuming 'Day' is a column in your DataFrame
        
        for col in timetable_df.columns[1:]:  # Skip 'Day' column
            cell_content = row[col]
            if pd.notna(cell_content):  # Check if cell is not empty
                components = cell_content.split(',')  # Assuming components are comma-separated
                
                entry = {
                    'sub_division': components[0].strip(),
                    'subject': components[1].strip(),
                    'faculty': components[2].strip(),
                    'venue': components[3].strip(),
                    'class_number': components[4].strip()
                }
                
                if day not in processed_timetable:
                    processed_timetable[day] = {}
                processed_timetable[day][col] = entry

                # Handle practical classes spanning two columns
                if len(components) > 5:  # Adjust condition based on your data structure
                    next_col = timetable_df.columns[timetable_df.columns.get_loc(col) + 1]
                    if pd.notna(row[next_col]):
                        practical_content = f"{cell_content}, {row[next_col]}"
                        practical_components = practical_content.split(',')
                        practical_entry = {
                            'sub_division': practical_components[0].strip(),
                            'subject': practical_components[1].strip(),
                            'faculty': practical_components[2].strip(),
                            'venue': practical_components[3].strip(),
                            'class_number': practical_components[4].strip()
                        }
                        processed_timetable[day][next_col] = practical_entry

    return processed_timetable

def extract_short_forms(file_path):
    short_forms_start_row = 6  # Adjust based on where short forms start
    short_forms_df = pd.read_excel(file_path, skiprows=short_forms_start_row)
    
    short_forms_dict = {row['Short Form']: row['Full Form'] for index, row in short_forms_df.iterrows()}
    return short_forms_dict

def save_processed_data(processed_timetable, short_forms_dict, output_file):
    with pd.ExcelWriter(output_file) as writer:
        for day, entries in processed_timetable.items():
            df_entries = pd.DataFrame.from_dict(entries, orient='index')
            df_entries.to_excel(writer, sheet_name=day)

        # Save short forms in a separate sheet
        short_forms_df = pd.DataFrame(list(short_forms_dict.items()), columns=['Short Form', 'Full Form'])
        short_forms_df.to_excel(writer, sheet_name='Short Forms', index=False)

def main():
    file_path = "C:\\Users\\91774\\Downloads\\time\\student\\SY D.xlsx"  # Input file path
    output_file = "C:\\Users\\91774\\Downloads\\time"  # Output file path
    
    school_info = read_metadata(file_path)
    print("School Information:", school_info)
    
    timetable_df = read_timetable(file_path)
    
    processed_timetable = process_timetable(timetable_df)
    
    short_forms_dict = extract_short_forms(file_path)
    
    save_processed_data(processed_timetable, short_forms_dict, output_file)
    
    print("Timetable processing complete. Output saved to", output_file)

if __name__ == "__main__":
    main()
