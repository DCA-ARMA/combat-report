# main.py

from gpt_integration import save_improved_text_to_file, create_combat_report_from_file
from grades import collect_grades
from datetime import datetime

if __name__ == "__main__":
    # Step 1: Read 'input.txt', extract date, and enhance text
    try:
        with open("input.txt", "r", encoding="utf-8") as file:
            input_text = file.read()

        # Extract the date from the input text
        date = None
        lines = input_text.splitlines()
        if lines:
            # Assume the date is in the first line
            first_line = lines[0]
            # Try to extract the date from the first line
            # Assuming the date is before the first '|'
            if '|' in first_line:
                date_part = first_line.split('|')[0].strip()
                try:
                    # Parse the date string
                    parsed_date = datetime.strptime(date_part, '%b %d, %Y')
                    # Format the date as desired, e.g., "DD/MM/YYYY"
                    date = parsed_date.strftime('%d/%m/%Y')
                except ValueError:
                    # If parsing fails, leave date as the original string
                    date = date_part
            else:
                date = ""  # Or set to current date

        # If date is not found, set it to an empty string
        if not date:
            date = ""

        # Save the improved text to 'middle.txt'
        save_improved_text_to_file(input_text)

    except FileNotFoundError:
        print("Error: 'input.txt' file not found.")
        date = ""

    # Step 2: Collect grades from the user
    grades_data = collect_grades()

    # Step 3: Generate combat report document with grades
    create_combat_report_from_file(date=date, grades_data=grades_data)
