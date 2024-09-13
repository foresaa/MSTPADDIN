import win32com.client
from utils import convert_duration_to_days, convert_to_string

# Extract task, resource, and custom fields
def extract_fields_from_msp(file_path):
    ms_project = win32com.client.Dispatch("MSProject.Application")
    ms_project.FileOpen(file_path)
    ms_project.Visible = True
    project = ms_project.ActiveProject  

    if project is None:
        raise Exception("No active project found.")

    hours_per_day = project.HoursPerDay  # Get HoursPerDay from the project

    # Extract Task and Resource Information
    tasks_data = []
    resources_data = []
    custom_fields_data = []

    # Extract task data
    for task in project.Tasks:
        if task is not None:
            task_id = task.ID  # Extract Task ID
            days_duration = convert_duration_to_days(task.Duration, hours_per_day)
            task_data = {
                'TaskID': task_id,
                'Name': task.Name,
                'Start': convert_to_string(task.Start),
                'Finish': convert_to_string(task.Finish),
                'Duration': days_duration,
                'Status': task.Status,
                'Summary': task.Summary,
                'Critical': task.Critical,
                'Stop': convert_to_string(task.Stop),
                'Resume': convert_to_string(task.Resume)
            }
            tasks_data.append(task_data)

            # Extract all custom fields (Text, Number, Cost, Date, Flag, Duration, Start, Finish, OutlineCode)
            for i in range(1, 31):
                custom_fields_data.append({
                    'TaskID': task.ID,
                    'FieldType': 'Text',
                    'FieldName': f'Text{i}',
                    'FieldValue': getattr(task, f'Text{i}', None)
                })
            for i in range(1, 21):
                custom_fields_data.append({
                    'TaskID': task_id,
                    'FieldType': 'Number',
                    'FieldName': f'Number{i}',
                    'FieldValue': getattr(task, f'Number{i}', None)
                })
            for i in range(1, 11):
                custom_fields_data.append({
                    'TaskID': task_id,
                    'FieldType': 'Start',
                    'FieldName': f'Start{i}',
                    'FieldValue': convert_to_string(getattr(task, f'Start{i}', None))
                })
                custom_fields_data.append({
                    'TaskID': task_id,
                    'FieldType': 'Finish',
                    'FieldName': f'Finish{i}',
                    'FieldValue': convert_to_string(getattr(task, f'Finish{i}', None))
                })
                custom_fields_data.append({
                    'TaskID': task_id,
                    'FieldType': 'OutlineCode',
                    'FieldName': f'OutlineCode{i}',
                    'FieldValue': getattr(task, f'OutlineCode{i}', None)
                })

    # Sort custom fields by Task ID
    custom_fields_data = sorted(custom_fields_data, key=lambda x: x['TaskID'])

    # Extract resource data
    for resource in project.Resources:
        if resource is not None:
            resource_data = {
                'Name': resource.Name,
                'MaxUnits': resource.MaxUnits,
                'Cost': resource.Cost,
                'Group': resource.Group
            }
            resources_data.append(resource_data)

    ms_project.FileClose(False)  # Close the project without saving
    ms_project.Quit()  # Quit Microsoft Project

    return tasks_data, resources_data, custom_fields_data
