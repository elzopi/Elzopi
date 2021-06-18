' Note that the code below is early bound and requires a reference to the Excel library.
' The key changes are: We reference the Outlook Application Object directly,
' instead of using GetObject or CreateObject.
' We have to qualify Excel references with the Excel.Application object,
' instead of "Application." Otherwise it's nearly identical, and avoids the OMG
' considerations.

Sub getTasksFromOL()
    
    Call GetTasksData("1/1/2014", "7/31/2014")
    
End Sub
Sub GetTasksData(StartDate As Date, Optional EndDate As Date)
' -------------------------------------------------
' Notes:
' End Date is optional, if you want to pull from only one day, use: Call GetTasksData("7/14/2008")
' Needs MS Excel Object Library reference
' -------------------------------------------------
 
Dim olApp As Outlook.Application
Dim olNS As Outlook.NameSpace
Dim myTaskItems As Outlook.Items
Dim ItemstoCheck As Outlook.Items
Dim ThisTask As Outlook.TaskItem
 
Dim xlApp As Excel.Application
Dim rng As Excel.Range
Dim rngStart As Excel.Range
Dim rngHeader As Excel.Range
Dim MyBook As Excel.Workbook
 
Dim i As Long
Dim NextRow As Long
Dim ColCount As Long
 
Dim MyItem As Object
Dim StringToCheck As String
Dim arrData() As Variant
 
' if no end date is specified, EndDate variable will be "12:00:00 AM"
' the requestor only wants one day, so set EndDate = StartDate
If EndDate = "12:00:00 AM" Then
    EndDate = StartDate
End If
 
If EndDate < StartDate Then
    MsgBox "Those dates seem switched, please check them and try again.", vbInformation
    GoTo ExitProc
End If
 
If EndDate - StartDate > 28 Then
    ' ask if the requestor wants so much info
    If MsgBox("This could take some time. Continue anyway?", vbInformation + vbYesNo) = vbNo Then
        GoTo ExitProc
    End If
End If
 
Set olApp = Outlook.Application
 
' hook into default Tasks folder
Set olNS = olApp.GetNamespace("MAPI")
Set myTaskItems = olNS.GetDefaultFolder(olFolderTasks).Items
 
' ------------------------------------------------------------------
' the following code adapted from:
' http://www.outlookcode.com/article.aspx?id=30
' http://weblogs.asp.net/whaggard/archive/2007/03/21/retrieving-your-
' outlook-appointments-for-a-given-date-range.aspx
'
With myTaskItems
    .Sort "[StartDate]", False
    .IncludeRecurrences = True
End With
'
StringToCheck = "[StartDate] >= " & Quote(StartDate) & " AND [DueDate] <= " & Quote(EndDate)
Debug.Print StringToCheck
'
Set ItemstoCheck = myTaskItems.Restrict(StringToCheck)
Debug.Print ItemstoCheck.Count
' ------------------------------------------------------------------
 
If ItemstoCheck.Count > 0 Then
    ' we found at least one task
    ' check to make sure we have actual tasks, not infinite recurrence issues
    If ItemstoCheck.Item(1) Is Nothing Then GoTo ExitProc
 
    Set xlApp = Excel.Application
 
    xlApp.ScreenUpdating = False
 
    Set MyBook = xlApp.Workbooks.Add
 
    xlApp.Visible = True
 
    MyBook.Sheets(1).Name = Format(StartDate, "MMDDYYYY") & " - " & Format(EndDate, "MMDDYYYY")
    Set rngStart = MyBook.Sheets(1).Range("A1")
 
    Set rngHeader = Range(rngStart, rngStart.Offset(0, 3))
 
    ' with assistance from Jon Peltier http://peltiertech.com/WordPress and
    ' http://support.microsoft.com/kb/306022
 
    rngHeader.Value = Array("Subject", "Body", "Start Date", "Due Date")
 
    ColCount = rngHeader.Columns.Count
 
    ' now that we know how many rows and columns we need,
    ' resize the array accordingly
    ReDim arrData(1 To ItemstoCheck.Count, 1 To ColCount)
 
    For i = 1 To ItemstoCheck.Count
 
          Set ThisTask = ItemstoCheck.Item(i)
 
            arrData(i, 1) = ThisTask.Subject
            arrData(i, 2) = ThisTask.Body
            arrData(i, 3) = Format(ThisTask.StartDate, "MM/DD/YYYY HH:MM AM/PM")
            arrData(i, 4) = Format(ThisTask.DueDate, "MM/DD/YYYY HH:MM AM/PM")
 
    Next i
 
    rngStart.Offset(1, 0).Resize(ItemstoCheck.Count, ColCount).Value = arrData
 
    xlApp.ScreenUpdating = True
 
Else
    MsgBox "There are no tasks during the time you specified. Exiting now.", vbCritical
End If
 
 
ExitProc:
Set myTaskItems = Nothing
Set olNS = Nothing
Set olApp = Nothing
Set xlApp = Nothing
StringToCheck = vbNullString
Set ItemstoCheck = Nothing
Set MyBook = Nothing
Set rngStart = Nothing
Set rngHeader = Nothing
Set ThisTask = Nothing
Erase arrData
 
End Sub

Function Quote(MyText)
' from Sue Mosher's excellent book "Microsoft Outlook Programming"

    Quote = Chr(34) & MyText & Chr(34)

End Function

