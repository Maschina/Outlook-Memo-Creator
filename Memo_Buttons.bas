Attribute VB_Name = "Memo_Buttons"
Dim OldHTMLBody As String

Sub InsertMemo()

    Dim oInspector As Inspector
    Set oInspector = Application.ActiveInspector

    If oInspector Is Nothing Then
        MsgBox "No active inspector"
        
    Else
        Dim CurrentAppointment As AppointmentItem, CurrentMail As MailItem
        
        ' Check if appointment item - create mail item first, if so
        On Error Resume Next
        Set CurrentAppointment = oInspector.CurrentItem
        If Not CurrentAppointment Is Nothing Then
            ' It's an appointment
            Set CurrentMail = Application.CreateItem(olMailItem)
            
            Dim CalRcpnt As Recipient, MailRcpnt As Recipient
            For Each CalRcpnt In CurrentAppointment.Recipients
                Set MailRcpnt = CalRcpnt
                
                If MailRcpnt.Type <> olBCC Then
                    If MailRcpnt.AddressEntry.GetExchangeUser Is Nothing Then
                        CurrentMail.Recipients.Add MailRcpnt.Name & " <" & MailRcpnt.Address & ">"
                    Else
                        CurrentMail.Recipients.Add MailRcpnt
                    End If
                End If
            Next CalRcpnt
            
            CurrentMail.Recipients.ResolveAll
            
            CurrentMail.Subject = CurrentAppointment.Subject
            CurrentMail.Body = "-----Original Appointment-----" & vbCrLf & CurrentAppointment.Body
            CurrentMail.Display
            
            oInspector.Close (olPromptForSave)
        End If
    End If

    Memo_UserForm.Show
    
End Sub

Public Sub CategoriesButton()
    Dim Item As Outlook.MailItem
    Set Item = Application.ActiveInspector.CurrentItem
    Item.ShowCategoriesDialog
End Sub
