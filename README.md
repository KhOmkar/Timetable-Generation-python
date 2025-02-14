
# Classroom Timetable Generator

A Python utility to generate consolidated classroom schedules from multi-division timetables stored in Excel workbooks.

## Description

This tool processes an Excel workbook containing multiple division timetables and creates a comprehensive schedule for a specific classroom, combining all classes scheduled across different divisions.

## Features

- Multi-sheet processing (handles multiple divisions)
- Classroom-specific schedule extraction
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
   git clone https://github.com/yourusername/classroom-timetable-generator.git
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
   classroom = "H303"  # Target classroom
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

## Output Format

Generated Excel file includes:
- Combined schedule for specified classroom
- Format:
  ```
  SUBJECT101
  Prof. John Doe
  (Division A)
  ---
  SUBJECT202
  Prof. Jane Smith
  (Division B)
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
combined_schedule = generator.process_all_sheets("input.xlsm", "H303")
generator.save_classroom_schedule(combined_schedule, "output.xlsx", "H303")
```

## Error Handling

The program handles:
- Missing files
- Invalid formats
- Processing errors
- Output errors

## Contributing

1. Fork the repository
2. Create feature branch
3. Commit changes
4. Push to branch
5. Create Pull Request

## License

This project is licensed under the MIT License - see the LICENSE.md file for details

## Authors

* **Your Name** - *Initial work*

## Acknowledgments

* College/University Name
* Department Name
* Any other contributors

## Support

For support, email your.email@example.com or create an issue in the repository.

## Roadmap

- [ ] Command line interface
- [ ] Multiple classroom processing
- [ ] Custom time slot configuration
- [ ] PDF output option
- [ ] Web interface

## Project Status

Under active development