# main.py

from gpt_integration import save_improved_text_to_file, create_combat_report_from_file
from grades import collect_grades
from datetime import datetime

if __name__ == "__main__":
    # Step 1: Read 'input.txt', extract date, and enhance text
    try:
        with open("input.txt", "r", encoding="utf-8") as file:
            input_text = file.read()

        # Extract the date from the input text or set current date
        date = datetime.now().strftime('%d/%m/%Y')

        # Set manager name and other details
        manager_name = "יואב סמיפור"  # Replace with actual manager name if needed
        force_name = "כוח האימון"    # Replace with the actual force name
        location = "מיקום האימון"    # Replace with the actual location

        # Save the improved text to 'middle.txt'
        save_improved_text_to_file(input_text, date, manager_name, force_name, location)

    except FileNotFoundError:
        print("Error: 'input.txt' file not found.")
        date = datetime.now().strftime('%d/%m/%Y')

    # Step 2: Collect grades from the user
    grades_data = collect_grades()

    # Step 3: Generate combat report document with grades
    create_combat_report_from_file(date=date, grades_data=grades_data, signature=manager_name)
