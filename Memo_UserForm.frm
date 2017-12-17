VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Memo_UserForm 
   Caption         =   "Insert Memo"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   OleObjectBlob   =   "Memo_UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Memo_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

    Call ConvertHTMLIntoMail(txtMeetingTitle.Value, FormatDateTime(calDate.Value, vbShortDate), txtLocation.text, chkExcludeExternal.Value, chkParticipants.Value, chkMainobjectives.Value, chkSummary.Value, chkNotes.Value, chkActions.Value)
    Unload Me

End Sub

Private Sub UserForm_Initialize()

    Dim oInspector As Inspector
    Set oInspector = Application.ActiveInspector

    If oInspector Is Nothing Then
        MsgBox "No active inspector"
        
    Else
        Dim CurrentAppointment As AppointmentItem, CurrentMail As MailItem
                
        ' Retrieve mail item
        On Error Resume Next
        Set CurrentMail = oInspector.CurrentItem
        
        If Err <> 0 Then
            ' Mail item not retrievable
            MsgBox "Mail item is invalid and cannot be handled. I will exit now."
        End If
        
        If CurrentMail.Sent Then
            MsgBox "This is not an editable email"
            
        Else
            txtMeetingTitle.text = FilterSubject(CurrentMail.Subject)
            txtLocation.text = "Skype"
            calDate.Value = Date
        End If
    End If

End Sub
