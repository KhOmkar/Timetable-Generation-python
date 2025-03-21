# Timetable Generator

---
# Faculty Timetable Generator
---
# Classroom Timetable Generator

A Python utility to generate consolidated classroom schedules from multi-division timetables stored in Excel workbooks.

## Description

This tool processes an Excel workbook containing multiple division timetables and creates a comprehensive schedule for a specific classroom, combining all classes scheduled across different divisions.

## Features

- Multi-sheet processing (handles multiple divisions)
- Classroom-specific schedule extraction
- differentiation betwwen practical class and theoritical class
- Schedule conflict handling
- Formatted Excel output
- Division tracking
- Error handling and validation

## Requirements

- Python 3.x
- Required packages:
  ```
  pandas
  openpyxl
  ```

## Installation

1. Clone the repository:
   ```bash
   https://github.com/KhOmkar/Timetable-Generation-python
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Update the file paths in `main()`:
   ```python
   input_file = "path/to/your/input.xlsm"
   output_file = "path/to/output/classroom.xlsx"
   classroom = "XYZ"  # Target classroom
   ```

2. Run the script:
   ```bash
   python timetable_generator.py
   ```

## Input File Format

- Excel workbook (.xlsm/.xlsx)
- Multiple sheets (one per division)
- Division info in cell N3
- Timetable starts from row 7
- Cell format: "Subject Faculty Classroom"
- Contains the metadata below the schedule info (Teacher name, course_id, course name, classrooms and venues and the short codes assigned to them)

## Output Format

Generated Excel file includes:
- Combined schedule for specified classroom
- Format:
  ```
  SUB101 #classroom shortform
  ABC #Professor short code
  (A) #division A
  ---
  SUB202
  PQR #prof. short code
  (B) #division B
  ```

## Code Structure

```
timetable_generator.py
├── class TimetableGenerator
│   ├── __init__()
│   ├── create_timetable_structure()
│   ├── process_all_sheets()
│   └── save_classroom_schedule()
└── main()
```

## Example

```python
generator = TimetableGenerator()
combined_schedule = generator.process_all_sheets("input.xlsm", "XYZ")
generator.save_classroom_schedule(combined_schedule, "output.xlsx", "XYZ")
```

## Error Handling

The program handles:
- Missing files
- Invalid formats
- Processing errors
- Output errors


## Authors

* **Omkar Khilare** - *Initial work*


## Support

For support, email itsomkar.dev@gmail.com or create an issue in the repository.

## Roadmap

- [ ] Command line interface
- [ ] Multiple classroom processing
- [ ] Custom time slot configuration
- [ ] PDF output option
- [ ] Web interface

## Project Status

Under active development
