import requests
import json
import xml.etree.ElementTree as ET
from datetime import datetime

# ------------------------------------------------------
# Before running the script:
#   + Copy the URL of the project in the Open Project web page. 
#   + Paste the URL into the PROJECT_LINK variable below on line 11. 

PROJECT_LINK = "https://op.integer-tech.com/projects/integer-example-project/work_packages"

#   + In Open Project web page, click on your profile circle on the top right. 
#   + Click on "My Account" and go to "Access tokens" in the menu on the left. 
#   + Generate an API token (it will look like a long string of random characters) and copy it from the pop-up. 
#   + Paste the API token into the API_KEY variable on line 18. 

API_KEY = "API KEY HERE"

#   + Run the script. 
#   + The resultant XML file will match the name of the project and will be generated in the same folder as the script. 
#   + To open in MS Project, open MS Project, click "Open", and select the XML file. 
#   + You may have to change the file selection via the drop down on the bottom right to "All Files". 
# ------------------------------------------------------

# Modify link to get project name by removing the domain name and cleaning up the path
projectName = PROJECT_LINK.split('/')[4]

response = requests.request('GET', f"https://op.integer-tech.com/api/v3/projects/{projectName}/work_packages", auth=('apikey', API_KEY))

# Check if the API responds
if response.status_code != requests.codes.ok:
    raise SystemExit("Bad API response! Check if the API key is correct.")

# Load API return into JSON
jason = json.loads(response.text)

# Grab page size
pSize = jason["total"]

# Access API again with the correct page size
response = requests.request('GET', f"https://op.integer-tech.com/api/v3/projects/{projectName}/work_packages?pageSize={pSize}", auth=('apikey', API_KEY))

# Load new API response into JSON
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

    # Check if startDate exists, if not, only append end date; only really applicable to milestones
    if "startDate" in entry.keys():

        # Check if startDate and dueDate contain information, if not, exit with missing information
        if entry["startDate"] is None:
            raise SystemExit(f"Uh oh! Missing start date on Name: {subject}, ID: {id}. Please provide a valid start date.")
        else:
            startDate = datetime.strptime(entry["startDate"], '%Y-%m-%d')

        if entry["dueDate"] is None:
            raise SystemExit(f"Uh oh! Missing due date on Name: {subject}, ID: {id}. Please provide a valid due date.")
        else:
            endDate = datetime.strptime(entry["dueDate"], '%Y-%m-%d')

        duration = int(entry["duration"].strip("PD"))

    else:
        startDate = "n/a"
        endDate = datetime.strptime(entry["date"], '%Y-%m-%d')
        duration = 0

    # Create unsorted task list
    taskList.append([id, subject, relation, percent, startDate, endDate, duration])

# Create a list for sorting tasks for the heirarchy
sTask = []

# Get root tasks
for i, val in enumerate(reversed(taskList)):

    # If task is parentless
    if val[2] is None:

        # Assign hierarchy
        val.append(1)

        # Insert at the beginning of sorted tasks and remove from unsorted tasks
        sTask.insert(0, val)
        taskList.remove(val)

# Loop over unsorted list until empty
while taskList:
    for i, val in enumerate(reversed(taskList)):

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
            taskList.remove(val)

# Begin building XML file
root = ET.Element("Project")
root.set('xmlns', 'http://schemas.microsoft.com/project')

# Create headers for MS Project schema
ET.SubElement(root, "Name").text = f"{projectName}.xml"
ET.SubElement(root, "Title").text = str(projectName)

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
ET.SubElement(task, "Name").text = projectName
ET.SubElement(task, "Start").text = datetime.strftime(min(sanStart), '%Y-%m-%dT%H:%M:%S')
ET.SubElement(task, "Finish").text = datetime.strftime(max(sanEnd), '%Y-%m-%dT%H:%M:%S')
ET.SubElement(task, "Duration").text = f"PT{int(str(max(sanEnd) - min(sanStart)).removesuffix(" days, 0:00:00")) * 8}H0M0S"

# Build individual tasks
for i in range(len(sTask)):
    task = ET.SubElement(tasks, "Task")

    # Save OP ID for task relationships
    uid = ET.SubElement(task, "UID")
    uid.set('opID', str(sTask[i][0]))
    uid.text = str(i + 1)

    ET.SubElement(task, "Name").text = sTask[i][1]
    ET.SubElement(task, "Manual").text = "1"
    ET.SubElement(task, "OutlineLevel").text = str(sTask[i][7])
    ET.SubElement(task, "Priority").text = "500"

    # Format datetime object into a string and append duration
    if isinstance(sTask[i][4], datetime) is True:
        ET.SubElement(task, "Start").text = datetime.strftime(sTask[i][4], '%Y-%m-%dT%H:%M:%S')
        ET.SubElement(task, "ManualStart").text = datetime.strftime(sTask[i][4], '%Y-%m-%dT%H:%M:%S')
    
    ET.SubElement(task, "Finish").text = datetime.strftime(sTask[i][5], '%Y-%m-%dT%H:%M:%S')
    ET.SubElement(task, "Duration").text = f"PT{sTask[i][6] * 8}H0M0S"
    ET.SubElement(task, "ManualFinish").text = datetime.strftime(sTask[i][5], '%Y-%m-%dT%H:%M:%S')
    ET.SubElement(task, "ManualDuration").text = f"PT{sTask[i][6] * 8}H0M0S"
    ET.SubElement(task, "DurationFormat").text = "7"
    ET.SubElement(task, "FreeformDurationFormat").text = "7"
    ET.SubElement(task, "PercentComplete").text = "0"
    ET.SubElement(task, "PercentWorkComplete").text = str(sTask[i][3])
    ET.SubElement(task, "RemainingDuration").text = f"PT{sTask[i][6] * 8}H0M0S"

# Build task relationships
for t in root.iter("Task"):

    # Grab Open Project's ID
    opID = t.find('UID').get('opID')

    # Get API response for relations
    relationResponse = requests.request('GET', f"https://op.integer-tech.com/api/v3/work_packages/{opID}/relations", auth=('apikey', API_KEY))
    relationJson = json.loads(relationResponse.text)

    # Check if response is not an error and contains any relations
    if relationJson["_type"] != "Error":
        if relationJson["total"] != 0:
            
            # Grab individual relations for each task
            for entry in relationJson["_embedded"]["elements"]:

                # Set target and origin task
                origin = int(entry["_links"]["from"]["href"].lstrip("/api/v3/work_packages"))
                target = int(entry["_links"]["to"]["href"].lstrip("/api/v3/work_packages"))

                # Only add relation if the origin is the current task; this prevents repeat relations
                if origin == int(opID):
                    
                    # Find the target's ID using Open Project ID
                    pred = root.find(f".//UID[@opID='{target}']")

                    # Build predecessor tag and assign predecessor
                    PredLink = ET.SubElement(t, "PredecessorLink")
                    ET.SubElement(PredLink, "PredecessorUID").text = pred.text

# Write to XML file
output = ET.ElementTree(root)
ET.indent(output, space = '\t', level = 0)
output.write(f"{projectName}.xml")