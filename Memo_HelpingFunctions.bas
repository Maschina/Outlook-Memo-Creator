Attribute VB_Name = "Memo_HelpingFunctions"
Function FilterPlaceholders(text As String) As String
    Dim regEx As New RegExp
        
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "(<!--%[\w-.]+%-->)" '<!--%...%-->
    End With
    
    FilterPlaceholders = regEx.Replace(text, "")
End Function

Function GroupRcpntsInDomain(ByRef RcpntsLst As Collection) As Scripting.Dictionary
    If RcpntsLst.Count = 0 Then
        Exit Function
    End If

    Dim RcpntsInDomain As New Scripting.Dictionary
    Dim Domain As String
    Dim DomainRcpnt As New Collection
    Dim RcpntIdx As Integer
    Dim Rcpnt As Recipient
    Do
        Set Rcpnt = RcpntsLst.Item(1)
        Domain = GetDomain(Rcpnt)
        Set DomainRcpnt = Nothing
        DomainRcpnt.Add Rcpnt
        RcpntsInDomain.Add Domain, DomainRcpnt
        RcpntsLst.Remove 1
        RcpntIdx = 1
        Do While RcpntsLst.Count > 0 And RcpntIdx <= RcpntsLst.Count
            Set Rcpnt = RcpntsLst.Item(RcpntIdx)
            If Domain = GetDomain(Rcpnt) Then
                Set DomainRcpnt = RcpntsInDomain(Domain)
                DomainRcpnt.Add Rcpnt
                RcpntsInDomain.Remove Domain
                RcpntsInDomain.Add Domain, DomainRcpnt
                RcpntsLst.Remove RcpntIdx
            Else
                RcpntIdx = RcpntIdx + 1
            End If
        Loop
    Loop While RcpntsLst.Count > 0
    
    Set GroupRcpntsInDomain = RcpntsInDomain
End Function

Function CopyRecipients(ByRef Coll As Recipients) As Collection
    Dim Copy As New Collection
    
    For Each El In Coll
        Copy.Add El
    Next El
    
    Set CopyRecipients = Copy
End Function

Function GetMailAddr(ByVal Rcpnt As Recipient) As String
    Dim RcpntAddr As String
    
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    RcpntAddr = Rcpnt.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
    
    GetMailAddr = RcpntAddr
End Function

Function GetDomain(ByVal Rcpnt As Recipient) As String
    Dim RcpntAddr, RcpntDomain As String
    RcpntAddr = GetMailAddr(Rcpnt)
    
    Dim regEx As New RegExp
        
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "^[_a-z0-9-.]+@([a-z0-9-]+)(.[a-z]{2,})$" 'Matching ro.hahn@posteo.de
    End With
    
    RcpntDomain = regEx.Replace(RcpntAddr, "$1")
    ' Capitalize first letter
    GetDomain = UCase(Left(RcpntDomain, 1)) & Mid(RcpntDomain, 2)
End Function

Function FilterSubject(Subject As String) As String
    Dim regEx As New RegExp
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "^([Rr][Ee]:\W)(.*)$" 'Matching patterns such as "Re: ..."
    End With
    
    FilterSubject = regEx.Replace(Subject, "$2")
End Function

Function GetHTMLSection(SeperatorString As String, text As String) As String
    Dim StartPosition, EndPosition As Long
    
    StartPosition = InStr(text, "<!--%" & SeperatorString & "-BEGIN%-->") - 1
    EndPosition = InStr(text, "<!--%" & SeperatorString & "-END%-->") + Len("<!--%" & SeperatorString & "-END%-->")
    
    GetHTMLSection = Mid(text, StartPosition, EndPosition - StartPosition)
End Function

Function RemoveHTMLSection(SeperatorString As String, text As String) As String
    Dim StartPosition, EndPosition As Long

    StartPosition = InStr(text, "<!--%" & SeperatorString & "-BEGIN%-->") - 1
    EndPosition = InStr(text, "<!--%" & SeperatorString & "-END%-->") + Len("<!--%" & SeperatorString & "-END%-->")
    
    RemoveHTMLSection = Left(text, StartPosition) & Mid(text, EndPosition)
End Function

Function InsertHTMLSection(PlaceholderString As String, InsertString As String, text As String) As String
    Dim InsertPosition As Long
    
    InsertPosition = InStr(text, "<!--%" & PlaceholderString & "-INSERTPOINT%-->") - 1
    
    InsertHTMLSection = Left(text, InsertPosition) & InsertString & Mid(text, InsertPosition)
End Function

