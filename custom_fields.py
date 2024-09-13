import win32com.client
from win32com.client import constants

# Dispatch the MSProject Application
project = win32com.client.Dispatch("MSProject.Application")
project.Visible = True  # Ensure MS Project is visible

# Access the active project
active_project = project.ActiveProject

# Ensure there's an active project loaded
if active_project is None:
    print("No active project found.")
else:
    # Define the PjValueListItem constant for value retrieval
    pjValueListValue = constants.pjValueListValue

    # Define the list of task-related custom field constants (using pjCustomTask)
    custom_field_constants = [
        constants.pjCustomTaskText1,  # Custom Task Text1
        constants.pjCustomTaskText2,  # Custom Task Text2
        constants.pjCustomTaskOutlineCode1,  # Custom Task Outline Code 1
        constants.pjCustomTaskNumber1,  # Custom Task Number1
        # Add more constants as needed
        # constants.pjCustomTaskText3, constants.pjCustomTaskOutlineCode2, etc.
    ]

    # Iterate through all custom fields
    for custom_field_constant in custom_field_constants:
        # Get the custom field name
        custom_field_name = project.CustomFieldGetName(custom_field_constant)
        if custom_field_name:
            print(f"Custom Field Name: {custom_field_name}")
        else:
            print("Custom Field Name: Unnamed Field")

        # Get the formula for the custom field, if any
        custom_field_formula = project.CustomFieldGetFormula(custom_field_constant)
        if custom_field_formula:
            print(f"Custom Field Formula for {custom_field_name}: {custom_field_formula}")
        else:
            print(f"No formula for {custom_field_name}")

        # Get lookup list items for custom field
        lookup_items = []
        try:
            index = 1
            while True:
                # Retrieve each lookup list item by index using the corrected call
                lookup_item = project.CustomFieldValueListGetItem(custom_field_constant, pjValueListValue, index)
                if lookup_item:
                    lookup_items.append(lookup_item)
                index += 1
        except Exception as e:
            # Ignore the "1107" error as it indicates no more lookup items
            if   "1101"  not in str(e):
                print(f"Unexpected error retrieving lookup items for {custom_field_name}: {e}")

        if lookup_items:
            print(f"Lookup items for {custom_field_name}: {', '.join(lookup_items)}")
        else:
            print(f"No lookup items for {custom_field_name}")
        
        print("-" * 50)
