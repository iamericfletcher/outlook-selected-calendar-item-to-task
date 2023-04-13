Sub CreateReschedulingTaskFromCalendarItem()
    Dim objCalendarItem As AppointmentItem
    Dim objTask As TaskItem
    Dim objNamespace As NameSpace
    Dim objTaskFolder As Folder
    Dim strCategory As String
    
    ' Specify the target Task Folder name
    Const TARGET_FOLDER_NAME As String = ""
    
    ' Specify the category to apply
    strCategory = ""
    
    ' Handle errors while working with objects
    On Error Resume Next
    
    ' Get the currently selected item in Outlook
    Set objCalendarItem = Application.ActiveExplorer.Selection.Item(1)
    
    ' Check if a calendar item is selected
    If objCalendarItem Is Nothing Then
        MsgBox "Please select a calendar item first.", vbInformation
        Exit Sub
    End If
    
    ' Check if the selected item is an appointment
    If objCalendarItem.Class = olAppointment Then
        ' Get the MAPI namespace
        Set objNamespace = Application.GetNamespace("MAPI")
        
        ' Locate the target task folder
        Set objTaskFolder = objNamespace.GetDefaultFolder(olFolderTasks).Folders(TARGET_FOLDER_NAME)
        
        ' Check if the target task folder exists
        If objTaskFolder Is Nothing Then
            MsgBox "The specified task folder does not exist. Please check the folder name.", vbInformation
            Exit Sub
        End If
        
        ' Create a new task item in the target task folder
        Set objTask = objTaskFolder.Items.Add(olTaskItem)
        With objTask
            .Subject = UCase(objCalendarItem.Subject) ' Set the task subject to the calendar item subject in all caps
            .DueDate = objCalendarItem.Start ' Set the task due date to the calendar item date
            .Categories = strCategory ' Set the task category
            .ReminderSet = True ' Enable reminder for the task
            .ReminderTime = Date ' Set the reminder time to the current date
            .Save ' Save the new task item
        End With
        
        ' Update the calendar item: set show as status to Free and add "MOVE - " to the title
        With objCalendarItem
            .BusyStatus = olFree
            .Subject = "MOVE - " & .Subject
            .Save
        End With
        
        MsgBox "A new task has been created in the specified folder.", vbInformation
        
        ' Open the calendar item for editing
        objCalendarItem.Display
    Else
        MsgBox "Please select a calendar item.", vbInformation
    End If
    
    ' Release the object variables
    Set objCalendarItem = Nothing
    Set objTask = Nothing
    Set objNamespace = Nothing
    Set objTaskFolder = Nothing
End Sub

