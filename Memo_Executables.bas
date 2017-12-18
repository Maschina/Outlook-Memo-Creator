Attribute VB_Name = "Memo_Executables"
Sub ConvertHTMLIntoMail(Subject As String, CurrentDate As String, Location As String, SwitchExcludeExternals As Boolean, SwitchPARTICIPANTS As Boolean, SwitchMAINOBJECTIVES As Boolean, SwitchSUMMARY As Boolean, SwitchNOTES As Boolean, SwitchACTIONS As Boolean)

    ' Open HTML file
    Dim MemoHTMLFile As String, text As String, textline As String
    MemoHTMLFile = "C:\Users\HartmRob\Macros\Memo-New.htm"
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
            ' Fill in placeholders
            MemoHTML = Replace(MemoHTML, "%SUBJECT%", Subject)
            MemoHTML = Replace(MemoHTML, "%DATE%", CurrentDate)
            MemoHTML = Replace(MemoHTML, "%LOCATION%", Location)
            
            ' Switch on/off participant section
            If Not SwitchPARTICIPANTS Then
                MemoHTML = RemoveHTMLSection("PARTICIPANTS", MemoHTML)
            Else
                ' Fill in participants
                Dim CurrentUserIdx As Long
                On Error Resume Next
                CurrentUserIdx = CurrentMail.Recipients.Item(Application.Session.CurrentUser.Name).Index
                If Err <> 0 Then
                    CurrentMail.Recipients.Add Application.Session.CurrentUser.Name
                    CurrentUserIdx = CurrentMail.Recipients.Item(Application.Session.CurrentUser.Name).Index
                End If
                CurrentMail.Recipients.ResolveAll

                Dim RcpntsLst As Collection
                Set RcpntsLst = CopyRecipients(CurrentMail.Recipients)
                CurrentMail.Recipients.Remove CurrentUserIdx
                
                ' Group participants into company/domain
                Dim RcpntsInDomain As New Scripting.Dictionary
                Set RcpntsInDomain = GroupRcpntsInDomain(RcpntsLst)
                
                ' Get text templates for company/domain and participants
                Dim PcpntsCompanyLoop As String, PcpntsPersonLoop As String
                PcpntsCompanyLoop = GetHTMLSection("PARTICIPANTS-COMPANY-LOOP", MemoHTML)
                PcpntsPersonLoop = GetHTMLSection("PARTICIPANTS-PERSON-LOOP", MemoHTML)
                
                ' Remove text templates for company/domain and participants
                MemoHTML = RemoveHTMLSection("PARTICIPANTS-COMPANY-LOOP", MemoHTML)
                MemoHTML = RemoveHTMLSection("PARTICIPANTS-PERSON-LOOP", MemoHTML)
                
                ' Text to be inserted for company/domain
                Dim PcpntsHTML As String
                
                If RcpntsInDomain.Count = 0 Then
                    ' Handle empty list of recipients
                    PcpntsHTML = PcpntsHTML & Replace(PcpntsCompanyLoop, "%PARTICIPANT-COMPANY%", "")
                Else
                
                    ' Handle full list of recipients
                    Dim RcpntKey As Variant
                    For Each RcpntKey In RcpntsInDomain
                        Dim Domain As String
                        Dim DomainRcpnts As New Collection
                        Domain = RcpntKey
                        Set DomainRcpnts = RcpntsInDomain(RcpntKey)
                        
                        ' Insert company/domain text
                        PcpntsHTML = PcpntsHTML & Replace(PcpntsCompanyLoop, "%PARTICIPANT-COMPANY%", Domain)
                        
                        ' Insert participants text
                        Dim Rcpnt As Recipient
                        For Each Rcpnt In DomainRcpnts
                            If Rcpnt.AddressEntry.GetExchangeUser Is Nothing Then
                                PcpntsHTML = PcpntsHTML & Replace(PcpntsPersonLoop, "%PARTICIPANT-PERSON%", Rcpnt.Name & " (" & Rcpnt.Address & ")")
                            Else
                                PcpntsHTML = PcpntsHTML & Replace(PcpntsPersonLoop, "%PARTICIPANT-PERSON%", Rcpnt.Name)
                            End If
                        Next Rcpnt
                    Next RcpntKey
                End If
                
                ' Finally insert domains
                MemoHTML = InsertHTMLSection("PARTICIPANTS", PcpntsHTML, MemoHTML)
            End If
                        
            ' Switch on/off main objective section
            If Not SwitchMAINOBJECTIVES Then
                MemoHTML = RemoveHTMLSection("MAINOBJECTIVES", MemoHTML)
            End If
            
            ' Switch on/off summary section
            If Not SwitchSUMMARY Then
                MemoHTML = RemoveHTMLSection("SUMMARY", MemoHTML)
            End If
            
            ' Switch on/off notes section
            If Not SwitchNOTES Then
                MemoHTML = RemoveHTMLSection("NOTES", MemoHTML)
            End If
            
            ' Switch on/off actions section
            If Not SwitchACTIONS Then
                MemoHTML = RemoveHTMLSection("ACTIONS", MemoHTML)
            End If
            
            ' Optionally exclude externals (employees from other domain/company)
            If SwitchExcludeExternals Then
                Dim OtherDomain As String, OwnDomain As String, CurrentRcpntIdx As Long
                OwnDomain = GetDomain(Application.Session.CurrentUser)
                CurrentRcpntIdx = 1
                Do
                    Set Rcpnt = CurrentMail.Recipients.Item(CurrentRcpntIdx)
                    OtherDomain = GetDomain(Rcpnt)
                    If OtherDomain <> OwnDomain Then
                        CurrentMail.Recipients.Remove Rcpnt.Index
                        CurrentRcpntIdx = CurrentRcpntIdx - 1
                    End If
                    CurrentRcpntIdx = CurrentRcpntIdx + 1
                Loop Until CurrentRcpntIdx >= CurrentMail.Recipients.Count
            End If
            
            MemoHTML = FilterPlaceholders(MemoHTML)
                        
            ' Add items to mail item
            CurrentMail.Subject = "[Memo " & CurrentDate & "] " & Subject
            CurrentMail.Save
            CurrentMail.BodyFormat = olFormatHTML
            CurrentMail.HTMLBody = MemoHTML & CurrentMail.HTMLBody
                                    
            ' Set category to "Memo"
            CurrentMail.Categories = CurrentMail.Categories & "," & "Memo"
        End If
    End If

End Sub
