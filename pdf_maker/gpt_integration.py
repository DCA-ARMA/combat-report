# gpt_integration.py

import os
import openai
import re
from dotenv import load_dotenv
from templates import generate_word_document

# Load environment variables from .env file
load_dotenv()

# Set up OpenAI API key from environment variable
openai.api_key = os.getenv("OPENAI_API_KEY")

def improve_text(text):
    """
    Sends the input text to the GPT-4 model to make it more impressive and returns the raw response.
    """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an assistant that improves text by making it more impressive and organizes it into four parts: "
                        "Introduction, Scenario 1, Scenario 2, and Summary, all in Hebrew. Ensure the Summary is written in a positive tone "
                        "and enhance the overall Hebrew language quality. "
                        "Please ensure that each section starts with a heading on its own line, like 'מבוא', 'תרחיש 1', 'תרחיש 2', and 'סיכום', "
                        "without any additional formatting or symbols."
                    )
                },
                {
                    "role": "user",
                    "content": f"Please improve this text and divide it into four parts as instructed: {text}"
                }
            ]
        )
        raw_text = response.choices[0].message['content'].strip()
        print("Raw LLM response:\n", raw_text)  # Print the raw response for debugging

        # Return raw LLM response directly
        return raw_text

    except Exception as e:
        print(f"Error occurred while communicating with LLM: {e}")
        return ""

def parse_to_sections(text):
    """
    Splits the text into a dictionary with four sections: Introduction, Scenario 1, Scenario 2, and Summary.
    Recognizes both English and Hebrew headers, even if they have additional text after the header.
    """
    # Remove unwanted formatting characters
    text = text.replace('**', '').replace('*', '').strip()

    # Define default empty sections
    sections = {"Introduction": "", "Scenario 1": "", "Scenario 2": "", "Summary": ""}

    # Define the possible section names in both English and Hebrew
    section_names = {
        "Introduction": ["מבוא", "Introduction"],
        "Scenario 1": ["תרחיש 1", "תרחיש ראשון", "Scenario 1", "Scenario One", "Scenario I"],
        "Scenario 2": ["תרחיש 2", "תרחיש שני", "Scenario 2", "Scenario Two", "Scenario II"],
        "Summary": ["סיכום ונקודות למחשבה", "סיכום", "תקציר", "Summary"]
    }

    # Prepare regex patterns for each section
    patterns = {}
    for section, names in section_names.items():
        # Create a regex pattern that matches any of the names, possibly with additional text
        name_pattern = r'|'.join(re.escape(name) for name in names)
        # Pattern to match the section heading
        patterns[section] = r'^\s*(%s)[^\n]*\n(.*?)(?=(^\s*(%s)[^\n]*\n|\Z))' % (
            name_pattern,
            '|'.join(re.escape(n) for n in sum(section_names.values(), []))
        )

    # Preprocess the text to ensure consistent line breaks
    text = text.replace('\r\n', '\n')

    # Initialize the sections dictionary
    extracted_sections = {}

    for section, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL | re.MULTILINE | re.IGNORECASE)
        if match:
            content = match.group(2).strip()
            # Remove '#' and '*' from content
            content = content.replace('#', '').replace('*', '')
            extracted_sections[section] = content
        else:
            print(f"Warning: Could not find section '{section}' in the text. It may be missing or formatted differently.")
            extracted_sections[section] = ''

    return extracted_sections

def save_improved_text_to_file(input_text):
    """
    Enhances the input text and saves the raw response to 'middle.txt'.
    """
    # Get the raw response from the LLM
    raw_text = improve_text(input_text)

    # Save the raw text to middle.txt
    with open("middle.txt", "w", encoding="utf-8") as file:
        file.write(raw_text)
    print("Enhanced text saved to 'middle.txt'")

def create_combat_report_from_file(file_path="middle.txt", date="", grades_data=None):
    """
    Reads the improved text from 'middle.txt', parses it into sections,
    and generates the combat report Word document with grades.
    """
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            improved_text = file.read()
    except FileNotFoundError:
        print(f"Error: '{file_path}' file not found.")
        return

    # Parse the text into sections
    sections = parse_to_sections(improved_text)

    # Delete the old combat report Word document if it exists
    doc_output_path = "combat_report.docx"
    if os.path.exists(doc_output_path):
        try:
            os.remove(doc_output_path)
            print(f"Deleted existing file '{doc_output_path}'")
        except Exception as e:
            print(f"Error deleting '{doc_output_path}': {e}")
            print("Please close the document if it is open in another program and try again.")
            return

    # Generate Combat Report Word Document with grades
    generate_word_document(
        sections,
        output_path=doc_output_path,
        date=date,
        signature="יואב סמיפור",
        title="אימון בסימולטור DCA",
        grades_data=grades_data
    )
    print(f"Combat report generated and saved as '{doc_output_path}'")
