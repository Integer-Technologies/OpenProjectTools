import requests
import json
import xml.etree.ElementTree as ET
import numpy as np
from datetime import datetime

# ------------------------------------------------------
# Before running the script:
#   + Copy the URL of the "WORK PACKAGES" page in the menu on the left. 
#   + Paste the URL into the PROJECT_LINK variable below on line 12. 

PROJECT_LINK = "https://op.integer-tech.com/projects/integer-example-project/work_packages"

#   + In Open Project, click on your profile circle on the top right. 
#   + Click on "My Account" and go to "Access tokens" in the menu on the left. 
#   + Generate an API token (it will look like a long string of random characters) and copy it from the pop-up. 
#   + Paste the API token into the API_KEY variable on line 19. 

API_KEY = "API KEY HERE"

#   + Run the script. 
#   + The resultant XML file will match the name of the project and will be generated in the same folder as the script. 
#   + To open in MS Project, open MS Project, click "Open", and select the XML file. 
#   + You may have to change the file selection via the drop down on the bottom right to "All Files". 
# ------------------------------------------------------

# Modify link to access API
api_link = PROJECT_LINK.removeprefix('https://op.integer-tech.com/')
response = requests.request('GET', 'https://op.integer-tech.com/api/v3/' + api_link + '?pageSize=1000', auth=('apikey', API_KEY))

name = api_link.removeprefix('projects/').removesuffix('/work_packages')

# Load API return into JSON
jason = json.loads(response.text)

# Grab project data from JSON
data = jason["_embedded"]["elements"]

# Create list for storing tasks
taskList = []

# Populate lists with relevant data with each object in "elements"
for entry in data:
    id = entry["id"]
    subject = entry["subject"]
    relation = entry["_links"]["parent"]["title"]
    percent = entry["percentageDone"]

    # Check if startDate exists, if not, only append end date
    if "startDate" in entry.keys():

        # Check if startDate and dueDate contain information
        if entry["startDate"] is None:
            startDate = "n/a"
        else:
            startDate = datetime.strptime(entry["startDate"], '%Y-%m-%d')

        if entry["dueDate"] is None:
            endDate = "n/a"
        else:
            endDate = datetime.strptime(entry["dueDate"], '%Y-%m-%d')

        # Check if duration exists
        if entry["startDate"] is None and entry["dueDate"] is None:
            duration = "n/a"
        else:
            duration = int(entry["duration"].strip("PD"))

    else:
        startDate = "n/a"
        endDate = datetime.strptime(entry["date"], '%Y-%m-%d')

    # Create unsorted task list
    taskList.append([id, subject, relation, percent, startDate, endDate, duration])

# Create a list for sorting tasks for the heirarchy
sTask = []

# Get root tasks (janky way to do this)
while None in [i[2] for i in taskList]:
    for i, val in enumerate(taskList):

        # If task is parentless
        if val[2] is None:

            # Assign hierarchy
            val.append(1)

            # Add to sorted tasks and remove from unsorted tasks
            sTask.append(val)
            taskList.pop(i)

# Loop over unsorted list until empty
while taskList:
    for i, val in enumerate(taskList):

        # Get current sorted parents
        parents = [child[1] for child in sTask]

        # If current task's parent is in the sorted task list
        if val[2] in parents:

            # Get position of parent
            position = parents.index(val[2])

            # Increment hierarchy from parent
            val.append(sTask[position][7] + 1)

            # Insert child after parent and remove from unsorted list
            sTask.insert(position + 1, val)
            taskList.pop(i)   

# Begin building xml file
root = ET.Element("Project")
root.set('xmlns', 'http://schemas.microsoft.com/project')

# Create headers for MS Project schema
ET.SubElement(root, "SaveVersion").text = "14"
ET.SubElement(root, "BuildNumber").text = "16.0.17328.20282"
ET.SubElement(root, "Name").text = f"{name}.xml"
ET.SubElement(root, "Title").text = str(name)

tasks = ET.SubElement(root, 'Tasks')

# Create a sanitized list of dates
sanStart, sanEnd = [], []

for i in range(len(sTask)):
    if isinstance(sTask[i][4], datetime) is True:
        sanStart.append(sTask[i][4])

    if isinstance(sTask[i][5], datetime) is True:
        sanEnd.append(sTask[i][5])

# Build overview task
task = ET.SubElement(tasks, "Task")
ET.SubElement(task, "UID").text = "0"
ET.SubElement(task, "Name").text = name
ET.SubElement(task, "Start").text = datetime.strftime(min(sanStart), '%Y-%m-%dT%H:%M:%S')
ET.SubElement(task, "Finish").text = datetime.strftime(max(sanEnd), '%Y-%m-%dT%H:%M:%S')
ET.SubElement(task, "Duration").text = f"PT{int(max(sanEnd) - min(sanStart)).removesuffix(" days, 0:00:00") * 8}H0M0S"

# Build individual tasks
for i in range(len(sTask)):
    
    task = ET.SubElement(tasks, "Task")
    ET.SubElement(task, "UID").text = str(i + 1)
    ET.SubElement(task, "Name").text = sTask[i][1]
    ET.SubElement(task, "Manual").text = "1"
    ET.SubElement(task, "OutlineLevel").text = str(sTask[i][7])
    ET.SubElement(task, "Priority").text = "500"

    # Format datetime object into a string and append duration
    # This is disgusting...
    if isinstance(sTask[i][4], datetime) is True:
        ET.SubElement(task, "Start").text = datetime.strftime(sTask[i][4], '%Y-%m-%dT%H:%M:%S')
    if isinstance(sTask[i][5], datetime) is True:
        ET.SubElement(task, "Finish").text = datetime.strftime(sTask[i][5], '%Y-%m-%dT%H:%M:%S')
    if isinstance(sTask[i][6], int) is True:
        ET.SubElement(task, "Duration").text = f"PT{sTask[i][6] * 8}H0M0S"
    if isinstance(sTask[i][4], datetime) is True:
        ET.SubElement(task, "ManualStart").text = datetime.strftime(sTask[i][4], '%Y-%m-%dT%H:%M:%S')
    if isinstance(sTask[i][5], datetime) is True:
        ET.SubElement(task, "ManualFinish").text = datetime.strftime(sTask[i][5], '%Y-%m-%dT%H:%M:%S')
    if isinstance(sTask[i][6], int) is True:
        ET.SubElement(task, "ManualDuration").text = f"PT{sTask[i][6] * 8}H0M0S"
    
    ET.SubElement(task, "ManualDuration").text = "PT24H0M0S"
    ET.SubElement(task, "DurationFormat").text = "7"
    ET.SubElement(task, "FreeformDurationFormat").text = "7"
    ET.SubElement(task, "IsSubproject").text = "0"
    ET.SubElement(task, "PercentComplete").text = "0"
    ET.SubElement(task, "PercentWorkComplete").text = str(sTask[i][3])

    # Calculate remaining duration in work hours
    if isinstance(sTask[i][6], int) is True:
        ET.SubElement(task, "RemainingDuration").text = f"PT{sTask[i][6] * 8}H0M0S"

# Write to XML file
output = ET.ElementTree(root)
output.write(f"{name}.xml")