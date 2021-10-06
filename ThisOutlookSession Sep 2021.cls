VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Provides option to categorize and automatically create a task from an email just before sending it
' Code has to reside at ThisOutlookSession
' Requires class built to take advantage of vba autofill-enumerate when coding:
'
' List friendly names of Unicode characters
' Public Enum eUnicodeConst
'     RightArrow
'     LeftArrow
'     Clock2
' End Enum
' Felix Reta June 2021, please credit my work if you use my code or part of it, it's just professional courtesy

Option Explicit

Private WithEvents olSentItems As Items
Attribute olSentItems.VB_VarHelpID = -1
Private WithEvents olCalendarItems As Items
Attribute olCalendarItems.VB_VarHelpID = -1
 
Private Sub Application_Startup()
  Dim objNS As NameSpace
  Set objNS = Application.Session
  ' instantiate objects declared WithEvents
  Set olSentItems = objNS.GetDefaultFolder(olFolderSentMail).Items
  Set olCalendarItems = objNS.GetDefaultFolder(olFolderCalendar).Items
  
  Set objNS = Nothing
End Sub

' Before actually sending the email
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

Dim xPrompt As String
Dim xOkOrCancel As Integer
Dim NewEmail As MailItem
Dim objTask As Outlook.TaskItem
Dim xYesNoorCancel As Integer
Dim xYesorNo As Integer

On Error Resume Next

If TypeOf Item Is MailItem Then

' If Item.Class = olMail Then
    Set NewEmail = Item

    xPrompt = "Do you want to process before sending the email?"
    xYesNoorCancel = MsgBox(xPrompt, vbYesNoCancel)

' try case option WIP

    Select Case xYesNoorCancel

    Case vbYes
         xPrompt = "Is this Cofnidential?"
         xYesorNo = MsgBox(xPrompt, vbYesNo)
         If xYesorNo = vbYes Then
            NewEmail.Sensitivity = olConfidential
         Else
            NewEmail.Sensitivity = olNormal
         End If
         
         NewEmail.ShowCategoriesDialog
'         NewEmail.MarkAsTask olMarkNextWeek 'or olmarkthisweek
'         NewEmail.TaskDueDate = Now + 5
'         NewEmail.FlagRequest = "Follow Up " & NewEmail.SenderName
'         NewEmail.ReminderSet = True
'         NewEmail.ReminderTime = Now + 4
    
'Create a corresponding task for reminders & follow up
        Set objTask = Application.CreateItem(olTaskItem)

        With objTask
            .Subject = UniConst(PeaceLove) & " " & UniConst(SailBoat) & " " & Item.Subject
            .StartDate = Item.ReceivedTime
            .DueDate = Item.ReceivedTime + 5
            .Body = Item.Body
            .Categories = Item.Categories
            .ReminderSet = True
            .ReminderTime = Now + 4
            .Attachments.Add NewEmail
            
            .Save
        
        End With
        
        Set objTask = Nothing

    Case vbNo
        Cancel = False
        
    Case vbCancel
        Cancel = True
        
End Select

    Set NewEmail = Nothing

End If

End Sub