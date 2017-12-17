Sub InsertMemo()
    ' Open HTML file
    Dim MemoHTMLFile As String, text As String, textline As String
    MemoHTMLFile = "C:\Users\HartmRob\Macros\Memo.html"
    Open MemoHTMLFile For Input As #1
    
    Dim MemoHTML As String
    
    Do Until EOF(1)
        Line Input #1, textline
        MemoHTML = MemoHTML & textline
    Loop
    
    Close #1

    
    ' Retrieve mail item
    Dim CurrentMail As MailItem, oInspector As Inspector
    Set oInspector = Application.ActiveInspector
    
    If oInspector Is Nothing Then
        MsgBox "No active inspector"
        
    Else
        Set CurrentMail = oInspector.CurrentItem
        
        If CurrentMail.Sent Then
            MsgBox "This is not an editable email"
            
        Else
            CurrentMail.Save
        
            ' Get current date
            Dim CurrentDate As String
            CurrentDate = FormatDateTime(Date, vbShortDate)
        
            ' Fill in placeholders
            Dim Subject As String
            Subject = CurrentMail.Subject
            
            MemoHTML = Replace(MemoHTML, "%DATE%", CurrentDate)
            MemoHTML = Replace(MemoHTML, "%SUBJECT%", Subject)
                        
            ' Add items to mail item
            CurrentMail.Subject = "[Memo " & CurrentDate & "] " & CurrentMail.Subject
            CurrentMail.Save
            CurrentMail.BodyFormat = olFormatHTML
            CurrentMail.HTMLBody = MemoHTML & CurrentMail.HTMLBody
        End If
    End If
End Sub

Public Sub CategoriesButton()
  Dim Item As Outlook.MailItem
  Set Item = Application.ActiveInspector.CurrentItem
  Item.ShowCategoriesDialog
End Sub
