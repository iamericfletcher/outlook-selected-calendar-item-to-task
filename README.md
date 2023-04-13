# Outlook-Calendar-To-Task-Converter
This VBA script for Microsoft Outlook automates the process of creating a new task from a selected calendar item. Specifically, I use this to automate a portion of my meeting rescheduling workflow. It greatly reduces the number of clicks thus saving me lots of time.

# Outlook Calendar to Task Script

This VBA script for Microsoft Outlook automates the process of creating a new task from a selected calendar item. When a calendar item is selected and the script is run, a new task is created in the specified task folder with the same title as the calendar item. The due date for the task is set to the date of the calendar item, and a specified category is applied. Additionally, the script updates the calendar item's title by adding "MOVE - " to the beginning and sets its show as status to free.

## Outlook Client and Version

This script has been developed and tested using the following Outlook client and version:

- Outlook Client: Microsoft Office LTSC Professional Plus 2021 (Desktop)
- Outlook Version: 2108

Please note that compatibility with other Outlook clients and versions is not guaranteed.

## Features

- Creates a new task from the selected calendar item
- Sets the due date of the task to the date of the calendar item
- Applies a specified category to the task
- Sets the reminder for the task to the current date
- Opens the calendar item for editing. This is a personal preference since I like to also add a note to the body of the calendar item detailing the need to reschedule, etc. This part cannot be automated at this time. 
- Updates the calendar item title by adding "MOVE - " to the beginning
- Sets the calendar item's show as status to free

## How to use

1. Open Microsoft Outlook.
2. Press `ALT + F11` to open the VBA editor.
3. Click on `Insert` > `Module` to create a new module.
4. Copy and paste the provided VBA script into the module.
5. Close the VBA editor.
6. Select a calendar item in Outlook.
7. Press `ALT + F8` to open the "Macro" dialog. Select the macro `CreateReschedulingTaskFromCalendarItem` and click "Run".

   Alternatively, you could add the macro to the Quick Access Toolbar:
   - Click on the small down arrow at the top of the Outlook window next to the Quick Access Toolbar.
   - Choose "More Commands" from the dropdown menu.
   - In the "Choose commands from" dropdown, select "Macros".
   - Find and select the macro `CreateReschedulingTaskFromCalendarItem` in the list, then click the "Add >>" button to add it to the Quick Access Toolbar. Click "OK" to save the changes.

## Customization

- Change the `TARGET_FOLDER_NAME` constant to specify the target task folder name.
- Change the `strCategory` variable to specify the category to apply to the new task.

## Note


## Note

- Please ensure that the target task folder is located in the default Tasks location in your Outlook folder structure. The script is designed to locate the folder in the Tasks location. If your task folder is located elsewhere, you might need to modify the script to accommodate the folder hierarchy.
- The calendar items this script was developed for typically have only one person added to them. Please keep this in mind when using the script with calendar items that have multiple attendees. You may need to make additional modifications to accommodate different scenarios involving multiple attendees.
