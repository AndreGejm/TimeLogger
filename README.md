Timelog Management Application
The Timelog Management Application is a Python script designed to manage and track user logon and logoff events, calculate the total logged time for each day, and maintain a backup of the log data in an Excel file.

Short Description
This application helps users keep track of their logon and logoff events, allowing them to monitor their activity and logged time. It records each event with a timestamp and calculates the total logged time for each day, storing the data in an Excel spreadsheet. Additionally, it creates a backup of the log data for each week.

Features
Record Events: Records logon and logoff events with timestamps.
Calculate Logged Time: Calculates the total logged time for each day based on logon and logoff events.
Create Backup: Creates a backup of the log data in an Excel file for each week.
Excel Integration: Utilizes the openpyxl library to interact with Excel spreadsheets for data storage and backup.
Usage
Run the Script: Execute the Python script timelog.py from the command line.

bash
Copy code
python timelog.py [event_type]
Specify Event Type: Provide the event type as an argument (logon or logoff).

Event Recording: The script records the specified event type with a timestamp.

Backup Creation: If a logoff event is recorded, the script calculates the total logged time for each day and creates a backup of the log data in an Excel file for the current week.

Configuration
Primary Excel Path: Specify the path to the primary Excel file (Timelog.xlsx) for storing the log data.
Backup Folder Path: Specify the path to the backup folder for storing weekly backup files.
Event Record Path: Specify the path to the event record file (event_record.txt) for recording log events.
Dependencies
Python 3.x
openpyxl library for Excel integration
License
This project is licensed under the MIT License. See the LICENSE file for details.

Author
[Your Name or Username] - [Your Website or GitHub Profile]
You can customize the README further by adding your name, website, GitHub profile, or any additional information you find relevant.
