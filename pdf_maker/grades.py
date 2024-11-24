# grades.py

def collect_grades():
    """
    Collects grades and comments for each item from the user.
    Returns a dictionary with the grades, comments, and calculated averages.
    """
    grades_data = {}

    # Define the grading structure
    parts = {
        "פיקוד ושליטה": {
            "1.1 גיבוש תמונת מצב": None,
            "1.2 ניהול הכוח, חלוקת גזרות, הזרמת כוחות": None,
            "1.3 מיקום המפקד המאפשר שליטה בכוח (לא להישאב לרובאות)": None
        },
        "עבודת קשר": {
            "2.1 נדב\"ר בסיסי - עלייה לפי פורמט": None,
            "2.2 אסרטיביות ופיקוד בדגש על שליטה בכוח ומניעת קשקשת ברשת": None,
            "2.3 דיווחים והכרזות בדגש על סיווג ואיפיון האירוע": None,
            "2.4 וידוא קבלה בעיקר בציון ידיעות חשובות": None
        },
        "מבצעיות | עקרונות לחימה": {
            "3.1 שימוש בשפה משותפת": None,
            "3.2 פכת\"ט | רואה, מעריך, ממליץ - דיווחים קצרים ומדוייקים": None,
            "3.3 מתן מענה לאירועים, קבלת החלטות נכונות": None,
            "3.4 שימוש במעטפת - כוחות חבירים, תצפיות וכו'": None,
            "3.5 הזדהות, חבירה וסגירת מעגלים - בדגש על מניעת דו\"צים": None
        }
    }

    print("Please enter grades between 1 and 10 for each item.\n")

    total_parts = 0
    total_parts_score = 0.0

    for part_name, items in parts.items():
        print(f"\n{part_name}:")
        part_total = 0.0
        num_items = len(items)
        for item_name in items:
            while True:
                try:
                    grade = float(input(f"Enter grade for {item_name}: "))
                    if 1 <= grade <= 10:
                        break
                    else:
                        print("Grade must be between 1 and 10.")
                except ValueError:
                    print("Invalid input. Please enter a number between 1 and 10.")
            items[item_name] = grade
            part_total += grade
        # Collect comment for the part
        comment = input(f"Enter comment for {part_name} (in Hebrew): ")
        # Calculate average for the part
        part_average = round(part_total / num_items, 2)
        grades_data[part_name] = {
            "items": items,
            "average": part_average,
            "comment": comment
        }
        total_parts_score += part_average
        total_parts += 1

    # Calculate final grade
    final_grade = round(total_parts_score / total_parts, 2)
    grades_data["final_grade"] = final_grade

    return grades_data
