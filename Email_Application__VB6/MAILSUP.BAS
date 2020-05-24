Attribute VB_Name = "Module1"
Public Const conMailLongDate = 0
Public Const conMailListView = 1

Public Const conOptionGeneral = 1       ' Constant for Option Dialog Type - General Options
Public Const conOptionMessage = 2       ' Constant for Option Dialog Type - Message Options

Public Const vbRecipTypeTo = 1
Public Const vbRecipTypeCc = 2

Public Const vbMessageFetch = 1
Public Const vbMessageSendDlg = 2
Public Const vbMessageSend = 3
Public Const vbMessageSaveMsg = 4
Public Const vbMessageCopy = 5
Public Const vbMessageCompose = 6
Public Const vbMessageReply = 7
Public Const vbMessageReplyAll = 8
Public Const vbMessageForward = 9
Public Const vbMessageDelete = 10
Public Const vbMessageShowAdBook = 11
Public Const vbMessageShowDetails = 12
Public Const vbMessageResolveName = 13
Public Const vbRecipientDelete = 14
Public Const vbAttachmentDelete = 15

Public Const vbAttachTypeData = 0
Public Const vbAttachTypeEOLE = 1
Public Const vbAttachTypeSOLE = 2



Type ListDisplay
    Name As String * 20
    Subject As String * 40
    Date As String * 20
End Type

Public currentRCIndex As Integer
Public UnRead As Integer
Public SendWithMapi As Integer
Public ReturnRequest As Integer
Public OptionType As Integer

' Windows API functions
#If Win32 Then
    Declare Function GetProfileString Lib "kernel32" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
#Else
    Declare Function GetProfileString% Lib "Kernel" (ByVal lpSection$, ByVal lpEntry$, ByVal lpDefault$, ByVal Buffer$, ByVal cbBuffer%)
#End If

Sub Attachments(Msg As Form)
    ' Clear the current attachment list.
    Msg.aList.Clear

    ' If there are attachments, load them into the list box.
    If VBMail.MapiMess.AttachmentCount Then
        Msg.NumAtt = VBMail.MapiMess.AttachmentCount & " Files"
        For i% = 0 To VBMail.MapiMess.AttachmentCount - 1
            VBMail.MapiMess.AttachmentIndex = i%
            a$ = VBMail.MapiMess.AttachmentName
            Select Case VBMail.MapiMess.AttachmentType
                Case vbAttachTypeData
                    a$ = a$ + " (Data File)"
                Case vbAttachTypeEOLE
                    a$ = a$ + " (Embedded OLE Object)"
                Case vbAttachTypeSOLE
                    a$ = a$ + " (Static OLE Object)"
                Case Else
                    a$ = a$ + " (Unknown attachment type)"
            End Select
            Msg.aList.AddItem a$
        Next i%
        
        If Not Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = True
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height + Msg.AttachWin.Height
            ' End If
        End If
    
    Else
        If Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = False
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height - Msg.AttachWin.Height
            ' End If
        End If
    End If
    Msg.Refresh
End Sub

Sub CopyNamestoMsgBuffer(Msg As Form, fResolveNames As Integer)
    Call KillRecips(VBMail.MapiMess)
    Call SetRCList(Msg.TxtTo, VBMail.MapiMess, vbRecipTypeTo, fResolveNames)
    Call SetRCList(Msg.txtCC, VBMail.MapiMess, vbRecipTypeCc, fResolveNames)
End Sub

Function DateFromMapiDate$(ByVal S$, wFormat%)
' This procedure formats a MAPI date in one of
' two formats for viewing the message.
    Y$ = Left$(S$, 4)
    M$ = Mid$(S$, 6, 2)
    D$ = Mid$(S$, 9, 2)
    T$ = Mid$(S$, 12)
    Ds# = DateValue(M$ + "/" + D$ + "/" + Y$) + TimeValue(T$)
    Select Case wFormat
        Case conMailLongDate
            f$ = "dddd, mmmm d, yyyy, h:mmAM/PM"
        Case conMailListView
            f$ = "mm/dd/yy hh:mm"
    End Select
    DateFromMapiDate = Format$(Ds#, f$)
End Function

Sub DeleteMessage()
    ' If the currently active form is a message, set MListIndex to
    ' the correct value.
    If TypeOf Screen.ActiveForm Is MsgView Then
        MailLst.MList.ListIndex = Val(Screen.ActiveForm.Tag)
        ViewingMsg = True
    End If

   ' Delete the mail message.
    If MailLst.MList.ListIndex <> -1 Then
        VBMail.MapiMess.MsgIndex = MailLst.MList.ListIndex
        VBMail.MapiMess.Action = vbMessageDelete
        X% = MailLst.MList.ListIndex
        MailLst.MList.RemoveItem X%
        If X% < MailLst.MList.ListCount - 1 Then
            MailLst.MList.ListIndex = X%
        Else
            MailLst.MList.ListIndex = MailLst.MList.ListCount - 1
        End If
        VBMail.MsgCountLbl = Format$(VBMail.MapiMess.MsgCount) + " Messages"

        ' Adjust the index values for currently viewed messages.
        If ViewingMsg Then
            Screen.ActiveForm.Tag = Str$(-1)
        End If

        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) > X% Then
                    Forms(i).Tag = Val(Forms(i).Tag) - 1
                End If
            End If
        Next i
        
        ' If the user is viewing a message, load the next message into the MsgView form
        ' if the message isn't currently displayed.
        If ViewingMsg Then
            ' First check to see if the message is currently being viewed.
            WindowNum% = FindMsgWindow((MailLst.MList.ListIndex))
            If WindowNum% > 0 Then
                If Forms(WindowNum%).Caption <> Screen.ActiveForm.Caption Then
                    Unload Screen.ActiveForm
                     ' Find the correct window again and display it.  The index isn't valid after the unload.
                     Forms(FindMsgWindow((MailLst.MList.ListIndex))).Show
                Else
                     Forms(WindowNum%).Show
                End If
            Else
                Call LoadMessage(MailLst.MList.ListIndex, Screen.ActiveForm)
            End If
        Else
            ' Check to see if there was a window viewing the message, and unload the window.
            WindowNum% = FindMsgWindow(X%)
            If WindowNum% > 0 Then
                Unload Forms(X%)
            End If
        End If
     End If
End Sub

Sub DisplayAttachedFile(ByVal FileName As String)
On Error Resume Next
        ' Determine the filename extension.
        ext$ = FileName
        junk$ = Token$(ext$, ".")
        ' Get the application from the WIN.INI file.
        Buffer$ = String$(256, " ")
        errCode% = GetProfileString("Extensions", ext$, "NOTFOUND", Buffer$, Len(Left(Buffer$, Chr(0)) - 1))
        If errCode% Then
            Buffer$ = Mid$(Buffer$, 1, InStr(Buffer$, Chr(0)) - 1)
            If Buffer$ <> "NOTFOUND" Then
                ' Strip off the ^.EXT information from the string.
                EXEName$ = Token$(Buffer$, " ")
                errCode% = Shell(EXEName$ + " " + FileName, 1)
                If Err Then
                    MsgBox "Error occurred during the shell: " + Error$
                End If
            Else
                MsgBox "Application that uses: <" + ext$ + "> not found in WIN.INI"
            End If
        End If
End Sub

Function FindMsgWindow(Index As Integer) As Integer
' This function searches through the active windows
' and locates those with the MsgView type and then
' checks to see if the tag contains the index the user
' is searching for.
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) = Index Then
                    FindMsgWindow = i
                    Exit Function
                End If
            End If
        Next i
        FindMsgWindow = -1
End Function

Function GetHeader(Msg As Control) As String
Dim CR As String
CR = Chr$(13) + Chr$(10)
      Header$ = String$(25, "-") + CR
      Header$ = Header$ + "Form: " + Msg.MsgOrigDisplayName + CR
      Header$ = Header$ + "To: " + GetRCList(Msg, vbRecipTypeTo) + CR
      Header$ = Header$ + "Cc: " + GetRCList(Msg, vbRecipTypeCc) + CR
      Header$ = Header$ + "Subject: " + Msg.MsgSubject + CR
      Header$ = Header$ + "Date: " + DateFromMapiDate$(Msg.MsgDateReceived, conMailLongDate) + CR + CR
      GetHeader = Header$
End Function

Sub GetMessageCount()
    '  Reads all mail messages and displays the count.
    Screen.MousePointer = 11
    VBMail.MapiMess.FetchUnreadOnly = 0
    VBMail.MapiMess.Action = vbMessageFetch
    VBMail.MsgCountLbl = Format$(VBMail.MapiMess.MsgCount) + " Messages"
    Screen.MousePointer = 0
End Sub

Function GetRCList(Msg As Control, RCType As Integer) As String
' Given a list of recipients, this function returns
' a list of recipients of the specified type in the
' following format:
'
'       Person 1;Person 2;Person 3

    For i = 0 To Msg.RecipCount - 1
        Msg.RecipIndex = i
        If RCType = Msg.RecipType Then
                a$ = a$ + ";" + Msg.RecipDisplayName
        End If
    Next i
    If a$ <> "" Then
       a$ = Mid$(a$, 2)  ' Strip off the leading ";".
    End If
    GetRCList = a$
End Function

Sub KillRecips(MsgControl As Control)
    ' Delete each recipient.  Loop until no recipients exist.
    While MsgControl.RecipCount
        MsgControl.Action = vbRecipientDelete
    Wend
End Sub

Sub LoadList(mailctl As Control)
' This procedure loads the mail message headers
' into the MailLst.MList.  Unread messages have
' Chr$(187) placed at the beginning of the string.
    MailLst.MList.Clear
    UnRead = 0
    StartIndex = 0
    For i = 0 To mailctl.MsgCount - 1
        mailctl.MsgIndex = i
        If Not mailctl.MsgRead Then
            a$ = Chr$(187) + " "
            If UnRead = 0 Then
                StartIndex = i  ' Start position in the mail list.
            End If
            UnRead = UnRead + 1
        Else
            a$ = "  "
        End If
        a$ = a$ + Mid$(Format$(mailctl.MsgOrigDisplayName, "!" + String$(10, "@")), 1, 10)
        If mailctl.MsgSubject <> "" Then
            b$ = Mid$(Format$(mailctl.MsgSubject, "!" + String$(35, "@")), 1, 35)
        Else
            b$ = String$(30, " ")
        End If
        c$ = Mid$(Format$(DateFromMapiDate(mailctl.MsgDateReceived, conMailListView), "!" + String$(15, "@")), 1, 15)
        MailLst.MList.AddItem a$ + Chr$(9) + b$ + Chr$(9) + c$
        MailLst.MList.Refresh
    Next i

    MailLst.MList.ListIndex = StartIndex
    
    ' Enable the correct buttons.
    VBMail.Next.Enabled = True
    VBMail.Previous.Enabled = True
    VBMail![Delete].Enabled = True

    ' Adjust the value of the labels displaying message counts.
    If UnRead Then
        VBMail.UnreadLbl = " - " + Format$(UnRead) + " Unread"
        MailLst.Icon = MailLst.NewMail.Picture
    Else
        VBMail.UnreadLbl = ""
        MailLst.Icon = MailLst.nonew.Picture
    End If
End Sub
    

Sub LoadMessage(ByVal Index As Integer, Msg As Form)
' This procedure loads the specified mail message into
' a form to either view or edit a message.
    If TypeOf Msg Is MsgView Then
        a$ = MailLst.MList.List(Index)
        ' Message is unread; reset the text.
        If Mid$(a$, 1, 1) = Chr$(187) Then
            Mid$(a$, 1, 1) = " "
            MailLst.MList.List(Index) = a$
            UnRead = UnRead - 1
            If UnRead Then
                VBMail.UnreadLbl = Format$(UnRead) + " Unread"
            Else
                VBMail.UnreadLbl = ""
                ' Change the icon on the list window.
                MailLst.Icon = MailLst.nonew.Picture
            End If
        End If
    End If

    ' These fields only apply to viewing.
    If TypeOf Msg Is MsgView Then
        VBMail.MapiMess.MsgIndex = Index
        Msg.txtDate = DateFromMapiDate$(VBMail.MapiMess.MsgDateReceived, conMailLongDate)
        Msg.txtFrom = VBMail.MapiMess.MsgOrigDisplayName
        MailLst.MList.ItemData(Index) = True
    End If
    ' These fields apply to both form types.
    Call Attachments(Msg)
    Msg.txtNoteText = VBMail.MapiMess.MsgNoteText
    Msg.txtSubject = VBMail.MapiMess.MsgSubject
    Msg.Caption = VBMail.MapiMess.MsgSubject
    Msg.Tag = Index
    Call UpdateRecips(Msg)
    Msg.Refresh
    Msg.Show
End Sub

Sub LogOffUser()
    On Error Resume Next
    VBMail.MapiSess.Action = 2
    If Err <> 0 Then
        MsgBox "Logoff Failure: " + ErrorR
    Else
        VBMail.MapiMess.SessionID = 0
        ' Adjust the menu items.
        VBMail.Logoff.Enabled = 0
        VBMail.Logon.Enabled = -1
        ' Unload all forms except the MDI form.
        Do Until Forms.Count = 1
            i = Forms.Count - 1
            If TypeOf Forms(i) Is MDIForm Then
                ' Do nothing.
            Else
                Unload Forms(i)
            End If
        Loop
        ' Disable the toolbar buttons.
        VBMail.Next.Enabled = False
        VBMail.Previous.Enabled = False
        VBMail![Delete].Enabled = False
        VBMail.SendCtl(vbMessageCompose).Enabled = False
        VBMail.SendCtl(vbMessageReplyAll).Enabled = False
        VBMail.SendCtl(vbMessageReply).Enabled = False
        VBMail.SendCtl(vbMessageForward).Enabled = False
        VBMail.rMsgList.Enabled = False
        VBMail.PrintMessage.Enabled = False
        VBMail.DispTools.Enabled = False
        VBMail.EditDelete.Enabled = False
                          
        ' Reset the caption for the status bar labels.
        VBMail.MsgCountLbl = "Off Line"
        VBMail.UnreadLbl = ""
    End If

End Sub

Sub PrintLongText(ByVal LongText As String)
' This procedure prints a text stream to a printer and
' ensures that words are not split between lines and
' that they wrap as needed.
    Do Until LongText = ""
        Word$ = Token$(LongText, " ")
        If Printer.TextWidth(Word$) + Printer.CurrentX > Printer.Width - Printer.TextWidth("ZZZZZZZZ") Then
            Printer.Print
        End If
        Printer.Print " " + Word$;
    Loop
End Sub

Sub PrintMail()
    ' In List view, all selected messages are printed.
    ' In Message view, the selected message is printed.

    If TypeOf Screen.ActiveForm Is MsgView Then
        Call PrintMessage(VBMail.MapiMess, False)
        Printer.EndDoc
    ElseIf TypeOf Screen.ActiveForm Is MailLst Then
        For i = 0 To MailLst.MList.ListCount - 1
            If MailLst.MList.Selected(i) Then
                VBMail.MapiMess.MsgIndex = i
                Call PrintMessage(VBMail.MapiMess, False)
            End If
        Next i
        Printer.EndDoc
    End If
End Sub

Sub PrintMessage(Msg As Control, fNewPage As Integer)
'   This procedure prints a mail message.
    Screen.MousePointer = 11
    ' Start a new page if needed.
    If fNewPage Then
        Printer.NewPage
    End If
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.DrawWidth = 10
    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
    Printer.Print
    Printer.FontSize = 9.75
    Printer.Print "From:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print Msg.MsgOrigDisplayName
    Printer.Print "To:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print GetRCList(Msg, vbRecipTypeTo)
    Printer.Print "Cc:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print GetRCList(Msg, vbRecipTypeCc)
    Printer.Print "Subject:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print Msg.MsgSubject
    Printer.Print "Date:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print DateFromMapiDate$(Msg.MsgDateReceived, conMailLongDate)
    Printer.Print
    Printer.DrawWidth = 5
    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
    Printer.FontSize = 9.75
    Printer.FontBold = False
    Call PrintLongText(Msg.MsgNoteText)
    Printer.Print
    Screen.MousePointer = 0
End Sub

Sub SaveMessage(Msg As Form)
    ' Save the current subject and note text.
    ' Copy the message to the compose buffer.
    ' Reset the subject and message text.
    ' Save the message.
    svSub = Msg.txtSubject
    SVNote = Msg.txtNoteText
    VBMail.MapiMess.Action = vbMessageCopy
    VBMail.MapiMess.MsgSubject = svSub
    VBMail.MapiMess.MsgNoteText = SVNote
    VBMail.MapiMess.Action = vbMessageSaveMsg
End Sub

Sub SetRCList(ByVal NameList As String, Msg As Control, RCType As Integer, fResolveNames As Integer)
' Given a list of recipients:
'
'       Person 1;Person 2;Person 3
'
' this procedure places the names into the Msg.Recip
' structures.
    
    If NameList = "" Then
        Exit Sub
    End If

    i = Msg.RecipCount
    Do
        Msg.RecipIndex = i
        Msg.RecipDisplayName = Trim$(Token(NameList, ";"))
        If fResolveNames Then
            Msg.Action = vbMessageResolveName
        End If
        Msg.RecipType = RCType
        i = i + 1
    Loop Until (NameList = "")
End Sub

Sub SizeMessageWindow(MsgWindow As Form)
    If MsgWindow.WindowState <> 1 Then
        ' Determine the minimum window size based
        ' on the visiblity of AttachWin (Attachment window).
        If MsgWindow.AttachWin.Visible Then    ' Attachment window.
            MinSize = 3700
        Else
            MinSize = 3700 - MsgWindow.AttachWin.Height
        End If

        ' Maintain the minimum form size.
        If MsgWindow.Height < MinSize And (MsgWindow.WindowState = 0) Then
            MsgWindow.Height = MinSize
            Exit Sub

        End If
        ' Adjust the size of the text box.
        If MsgWindow.ScaleHeight > MsgWindow.txtNoteText.Top Then
            If MsgWindow.AttachWin.Visible Then
                X% = MsgWindow.AttachWin.Height
            Else
                X% = 0
            End If
            MsgWindow.txtNoteText.Height = MsgWindow.ScaleHeight - MsgWindow.txtNoteText.Top - X%
            MsgWindow.txtNoteText.Width = MsgWindow.ScaleWidth
        End If
    End If

End Sub

Function Token$(tmp$, search$)
    X = InStr(1, tmp$, search$)
    If X Then
       Token$ = Mid$(tmp$, 1, X - 1)
       tmp$ = Mid$(tmp$, X + 1)
    Else
       Token$ = tmp$
       tmp$ = ""
    End If
End Function

Sub UpdateRecips(Msg As Form)
' This procedure updates the correct edit fields and the
' recipient information.
    Msg.TxtTo.Text = GetRCList(VBMail.MapiMess, vbRecipTypeTo)
    Msg.txtCC.Text = GetRCList(VBMail.MapiMess, vbRecipTypeCc)
End Sub

Sub ViewNextMsg()
    ' Check to see if the message is currently loaded.
    ' If it is loaded, show that form.
    ' If it is not loaded, load the message.
    WindowNum% = FindMsgWindow((MailLst.MList.ListIndex))
    If WindowNum% > 0 Then
        Forms(WindowNum%).Show
    Else
        If TypeOf Screen.ActiveForm Is MsgView Then
            Call LoadMessage(MailLst.MList.ListIndex, Screen.ActiveForm)
        Else
            Dim Msg As New MsgView
            Call LoadMessage(MailLst.MList.ListIndex, Msg)
        End If
    End If
End Sub

