# Excel Folder Automation

Folder Automation with Hyperlinked Entries in Excel
This VBA macro automates the creation of folders with a specific naming pattern and inserts hyperlinks to each folder directly in an Excel worksheet. Ideal for managing time-sensitive projects or organizing monthly files, this macro reduces manual work and ensures consistent folder naming conventions.

## Key Features
### Automated Folder Creation: 
Generates folders in a specified directory based on the current month and a sequential numbering pattern.
### Dynamic Hyperlinks: 
Inserts links in Excel to each created folder, making it easy to navigate directly from the worksheet.
### Customizable: 
The target directory and starting row can be adjusted as needed.

## How It Works

Determines the last used row in the sheet based on entries for the current month.
Identifies the highest sequence number for the month, then generates three new folders with the pattern ```MM-XXX-24``` (assumption year: 2024)
Adds entries and hyperlinks in Excel, updating the sheet with the latest folders.
This project is perfect for users who frequently create folders and manage them in bulk, saving time on repetitive tasks.
