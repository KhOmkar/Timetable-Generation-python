import pandas as pd

def read_timetable(file_path):
    # Read the timetable data from the text file
    with open(file_path, 'r') as file:
        data = file.readlines()
    
    # Extract metadata from the first few lines
    metadata = {
        "Class": data[1].strip(),
        "Coordinator": data[2].strip(),
        "Division": data[3].strip(),
        "Academic Year": data[4].strip(),
        "Theory Classes": int(data[5].strip().split(":")[1]),
        "Practical Classes": int(data[6].strip().split(":")[1]),
        "Total Classes": int(data[7].strip().split(":")[1]),
    }
    
    # Extract timetable starting from a specific line
    timetable_start_index = 10  # Adjust based on your data structure
    timetable_data = data[timetable_start_index:]
                                
    # Create a DataFrame for the timetable
    timetable_lines = [line.strip().split('"') for line in timetable_data if line.strip()]
    timetable_df = pd.DataFrame(timetable_lines)
    
    # Set column names based on time slots (assuming first row contains time slots)
    time_slots = ["Day"] + [f"Slot {i+1}" for i in range(len(timetable_df.columns) - 1)]
    timetable_df.columns = time_slots
    
    return metadata, timetable_df

def process_timetable(timetable_df):
    processed_timetable = {}

    # Iterate through rows (days) and columns (time slots)
    for index, row in timetable_df.iterrows():
        day = row['Day']
        
        for col in timetable_df.columns[1:]:  # Skip 'Day' column
            cell_content = row[col]
            if pd.notna(cell_content) and cell_content.strip():  # Check if cell is not empty
                components = cell_content.split(',')
                entry = {
                    'Details': components,
                }
                
                if day not in processed_timetable:
                    processed_timetable[day] = {}
                processed_timetable[day][col] = entry

    return processed_timetable

def main():
    file_path = "C:\\Users\\91774\\Downloads\\time\\Classwise 24 25 Sem I.xlsm"  # Input file path
    metadata, timetable_df = read_timetable(file_path)
    
    print("Metadata:", metadata)
    
    processed_timetable = process_timetable(timetable_df)

    print("\nProcessed Timetable:")
    for day, entries in processed_timetable.items():
        print(f"\n{day}:")
        for slot, details in entries.items():
            print(f"  {slot}: {details['Details']}")

if __name__ == "__main__":
    main()
