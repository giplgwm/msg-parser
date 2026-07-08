Attribute VB_Name = "CopyEmailDetails"
Option Explicit

'==============================================================================
' CopyEmailDetails.bas - Outlook (classic) macro
'
' Copies the selected (or open) email to the clipboard as plain text:
'
'   From: Name <email>
'   To: Name <email>; Name <email>
'   CC: Name <email>              <- line omitted when there are no CC recipients
'   Sent time: yyyy-mm-dd hh:nn:ss
'   Subject: <subject>
'
'   <plain-text body, trimmed to the newest message - see TRIM_TO_LATEST>
'
' SMTP addresses are resolved the same way msg-parser.js does: prefer
' PidTagSenderSmtpAddress / PidTagSmtpAddress, then Exchange directory lookup,
' then the raw address. Selecting multiple emails copies all of them,
' separated by a divider line.
'
' Install: Alt+F11 -> File > Import File... and pick this .bas file.
'          (If you paste into a module instead, skip the "Attribute" line at the top.)
' Run:     Select an email, Alt+F8 -> CopyEmailDetails.
'          Optional: add it to the Quick Access Toolbar or a ribbon button.
' Note:    Requires Outlook 2010 or later (VBA7) and macros enabled:
'          File > Options > Trust Center > Trust Center Settings > Macro Settings.
'==============================================================================

' --- Win32 clipboard (Unicode-safe, no extra references needed) --------------
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal Length As LongPtr)

Private Const CF_UNICODETEXT As Long = 13
Private Const GHND As Long = &H42                ' GMEM_MOVEABLE Or GMEM_ZEROINIT

' MAPI property tags (the same ones msg-parser.js prefers)
Private Const PR_SENDER_SMTP_ADDRESS As String = _
    "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"   ' PidTagSenderSmtpAddress
Private Const PR_SMTP_ADDRESS As String = _
    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"   ' PidTagSmtpAddress

' Cut the body at the first quoted-reply separator so only the newest message
' is copied. Set to False to copy the entire chain.
Private Const TRIM_TO_LATEST As Boolean = True

'------------------------------------------------------------------------------
' Entry point
'------------------------------------------------------------------------------
Public Sub CopyEmailDetails()
    Dim mails As Collection
    Set mails = SelectedMailItems()
    If mails.Count = 0 Then
        MsgBox "Select or open an email first.", vbExclamation, "Copy Email Details"
        Exit Sub
    End If

    Dim out As String, i As Long
    For i = 1 To mails.Count
        If i > 1 Then out = out & vbCrLf & vbCrLf & String(40, "-") & vbCrLf & vbCrLf
        out = out & FormatMail(mails(i))
    Next i

    If Not SetClipboardText(out) Then
        MsgBox "Could not access the clipboard.", vbExclamation, "Copy Email Details"
    End If
End Sub

'------------------------------------------------------------------------------
' Formatting
'------------------------------------------------------------------------------
Private Function FormatMail(mail As Outlook.MailItem) As String
    Dim toList As String, ccList As String
    BuildRecipientLists mail, toList, ccList

    Dim s As String
    s = "From: " & FormatAddress(mail.SenderName, SenderSmtp(mail)) & vbCrLf
    s = s & "To: " & toList & vbCrLf
    If Len(ccList) > 0 Then s = s & "CC: " & ccList & vbCrLf
    s = s & "Sent time: " & Format$(mail.SentOn, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    s = s & "Subject: " & mail.Subject & vbCrLf & vbCrLf
    If TRIM_TO_LATEST Then
        s = s & TrimToLatestReply(mail.Body)
    Else
        s = s & mail.Body
    End If
    FormatMail = s
End Function

Private Sub BuildRecipientLists(mail As Outlook.MailItem, _
                                ByRef toList As String, ByRef ccList As String)
    Dim r As Outlook.Recipient
    Dim entry As String
    For Each r In mail.Recipients
        entry = FormatAddress(r.Name, RecipientSmtp(r))
        Select Case r.Type
            Case olTo: toList = AppendEntry(toList, entry)
            Case olCC: ccList = AppendEntry(ccList, entry)
        End Select
    Next r
End Sub

Private Function AppendEntry(ByVal list As String, ByVal entry As String) As String
    If Len(list) = 0 Then
        AppendEntry = entry
    Else
        AppendEntry = list & "; " & entry
    End If
End Function

' "Name <email>", collapsing to just the email when the display name is empty
' or identical to the address (mirrors how msg-parser demo renders addresses).
Private Function FormatAddress(ByVal dispName As String, ByVal addr As String) As String
    dispName = Trim$(dispName)
    addr = Trim$(addr)
    If Len(addr) = 0 Then
        FormatAddress = dispName
    ElseIf Len(dispName) = 0 Or StrComp(dispName, addr, vbTextCompare) = 0 Then
        FormatAddress = addr
    Else
        FormatAddress = dispName & " <" & addr & ">"
    End If
End Function

'------------------------------------------------------------------------------
' SMTP address resolution (same preference order as msg-parser.js)
'------------------------------------------------------------------------------
Private Function SenderSmtp(mail As Outlook.MailItem) As String
    Dim s As String
    On Error Resume Next
    ' 1) PidTagSenderSmtpAddress (0x5D01)
    s = mail.PropertyAccessor.GetProperty(PR_SENDER_SMTP_ADDRESS)
    ' 2) Exchange directory lookup for EX-type senders
    If Len(s) = 0 And mail.SenderEmailType = "EX" Then
        Dim exUser As Outlook.ExchangeUser
        Set exUser = mail.Sender.GetExchangeUser
        If Not exUser Is Nothing Then s = exUser.PrimarySmtpAddress
    End If
    ' 3) Raw sender address
    If Len(s) = 0 Then s = mail.SenderEmailAddress
    On Error GoTo 0
    SenderSmtp = s
End Function

Private Function RecipientSmtp(r As Outlook.Recipient) As String
    Dim s As String
    On Error Resume Next
    ' 1) PidTagSmtpAddress (0x39FE)
    s = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
    ' 2) Exchange directory lookup
    If Len(s) = 0 Then
        Dim exUser As Outlook.ExchangeUser
        Set exUser = r.AddressEntry.GetExchangeUser
        If Not exUser Is Nothing Then s = exUser.PrimarySmtpAddress
    End If
    ' 3) Raw address (0x3003 equivalent)
    If Len(s) = 0 Then s = r.Address
    On Error GoTo 0
    RecipientSmtp = s
End Function

'------------------------------------------------------------------------------
' Selection handling: open inspector, or selected item(s) in the explorer
'------------------------------------------------------------------------------
Private Function SelectedMailItems() As Collection
    Dim col As New Collection
    On Error Resume Next
    Dim win As Object
    Set win = Application.ActiveWindow
    If Not win Is Nothing Then
        If TypeOf win Is Outlook.Inspector Then
            If TypeOf win.CurrentItem Is Outlook.MailItem Then col.Add win.CurrentItem
        ElseIf TypeOf win Is Outlook.Explorer Then
            Dim it As Object
            For Each it In win.Selection
                If TypeOf it Is Outlook.MailItem Then col.Add it
            Next it
        End If
    End If
    On Error GoTo 0
    Set SelectedMailItems = col
End Function

'------------------------------------------------------------------------------
' Reply-chain trimming (heuristic): cut at the first line that looks like the
' start of a quoted earlier message. If no separator is found, keep everything.
'------------------------------------------------------------------------------
Private Function TrimToLatestReply(ByVal body As String) As String
    Dim lines() As String
    lines = Split(Replace(body, vbCrLf, vbLf), vbLf)

    Dim i As Long, cut As Long
    cut = -1
    For i = LBound(lines) To UBound(lines)
        If IsReplySeparator(lines, i) Then
            cut = i
            Exit For
        End If
    Next i

    If cut < 0 Then
        TrimToLatestReply = body                 ' no separator: full body
    ElseIf cut = 0 Then
        TrimToLatestReply = ""
    Else
        ReDim Preserve lines(0 To cut - 1)
        TrimToLatestReply = RTrimBlank(Join(lines, vbCrLf))
    End If
End Function

Private Function IsReplySeparator(lines() As String, ByVal i As Long) As Boolean
    Dim t As String
    t = Trim$(lines(i))
    If Len(t) = 0 Then Exit Function

    ' "-----Original Message-----" (any number of dashes)
    If Left$(t, 2) = "--" And InStr(1, t, "Original Message", vbTextCompare) > 0 Then
        IsReplySeparator = True
        Exit Function
    End If

    ' Outlook's divider: a line of underscores above the quoted header block
    If Len(t) >= 10 And t = String$(Len(t), "_") Then
        IsReplySeparator = True
        Exit Function
    End If

    ' Gmail / Apple Mail style: "On <date>, <name> wrote:"
    If StrComp(Left$(t, 3), "On ", vbTextCompare) = 0 And _
       StrComp(Right$(t, 6), "wrote:", vbTextCompare) = 0 Then
        IsReplySeparator = True
        Exit Function
    End If

    ' Outlook quoted header block: "From:" line with Sent:/Date:/To: right after.
    ' The look-ahead avoids false positives on a stray "From:" in normal text.
    If StrComp(Left$(t, 5), "From:", vbTextCompare) = 0 Then
        Dim j As Long, u As String
        For j = i + 1 To i + 5
            If j > UBound(lines) Then Exit For
            u = Trim$(lines(j))
            If StrComp(Left$(u, 5), "Sent:", vbTextCompare) = 0 _
               Or StrComp(Left$(u, 5), "Date:", vbTextCompare) = 0 _
               Or StrComp(Left$(u, 3), "To:", vbTextCompare) = 0 Then
                IsReplySeparator = True
                Exit Function
            End If
        Next j
    End If
End Function

Private Function RTrimBlank(ByVal s As String) As String
    Dim ch As String
    Do While Len(s) > 0
        ch = Right$(s, 1)
        If ch = vbCr Or ch = vbLf Or ch = " " Or ch = vbTab Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    RTrimBlank = s
End Function

'------------------------------------------------------------------------------
' Clipboard (CF_UNICODETEXT via Win32, so Unicode survives intact)
'------------------------------------------------------------------------------
Private Function SetClipboardText(ByVal text As String) As Boolean
    If Len(text) = 0 Then text = " "             ' StrPtr("") is 0

    Dim hMem As LongPtr
    hMem = GlobalAlloc(GHND, CLngPtr(LenB(text)) + 2)  ' +2 for the null terminator
    If hMem = 0 Then Exit Function

    Dim pMem As LongPtr
    pMem = GlobalLock(hMem)
    If pMem = 0 Then
        GlobalFree hMem
        Exit Function
    End If
    CopyMemory pMem, StrPtr(text), LenB(text)
    GlobalUnlock hMem

    If OpenClipboard(0) = 0 Then
        GlobalFree hMem
        Exit Function
    End If
    EmptyClipboard
    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then
        CloseClipboard
        GlobalFree hMem
        Exit Function
    End If
    CloseClipboard                               ' clipboard now owns hMem
    SetClipboardText = True
End Function
