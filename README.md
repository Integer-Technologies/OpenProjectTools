# Description
OP2MSP is a Python script to convert Open Project projects to Microsoft Project XML files. Exporting a project requires access to said project in Open Project. 

### Features
- Converts Open Project projects into Microsoft Project XMl files.
- Retains parent/child relations along with follows/precedes relationships.
- Lightweight script using mostly standard Python libraries. 

# Requirements
OP2MSP requires Python >=3.8. 

Libraries used:
- [requests](https://pypi.org/project/requests/) - A simple HTTP library for API requests. 
- [json](https://docs.python.org/3/library/json.html) - JSON parsing library. 
- [ElementTree](https://docs.python.org/3/library/xml.etree.elementtree.html) - An XML library used for parsing and creating XML data. 
- [datetime](https://docs.python.org/3/library/datetime.html) - A library used for easy formating and manipulation of dates. 

The export script requires the user to have access to the Open Project project. 

# Installation
Download the OP2MSP.py file. A text editor will be required for inserting the API key and project link into the script. Instructions are below but also within the script itself. 

### Usage
1. Copy the URL of the "WORK PACKAGES" page in the menu on the left on the Open Project web page. 
2. Paste the URL into the `PROJECT_LINK` variable on line 11 in the script. 
3. In the Open Project web page, click on your profile circle on the top right. 
4. Click on "My Account" and go to "Access tokens" in the menu on the left. 
5. Generate an API token (it will look like a long string of random characters) and copy it from the pop-up. 
6. Paste the API token into the `API_KEY` variable on line 18 in the script. 
7. Run the script. 
8. The resultant XML file will match the name of the project and will be generated in the same folder as the script. 
9. To open in MS Project, open MS Project, click "Open", and select the XML file. 
10. You may have to change the file selection via the drop down on the bottom right to "All Files".

# Known Issues
+ Some exported Gantt Charts in Microsoft Project may be truncated by one (1) day. However Start and Finish date entries along with Duration are correct. 
