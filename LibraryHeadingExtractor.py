import openpyxl

def extract_task_headings_from_dict(action_dict):
    return [key for key in action_dict.keys()]

def write_to_excel(task_headings, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Task Headings"

    # Write the task headings to the first column
    for idx, heading in enumerate(task_headings, start=1):
        sheet[f'A{idx}'] = heading

    workbook.save(output_file)

# Import the action_examples dictionary from Microsoft_Teams.py
try:
    from MicrosoftTeams import action_examples
except ImportError:
    raise ImportError("Unable to import the action_examples dictionary from Microsoft_Teams.py. Ensure the file and dictionary are correctly named and structured.")

# Extract task headings
task_headings = extract_task_headings_from_dict(action_examples)

# Write to Excel
output_file = 'task_headings.xlsx'
write_to_excel(task_headings, output_file)

print(f'Task headings written to {output_file}')
