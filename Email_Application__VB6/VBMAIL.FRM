VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm VBMAIL 
   BackColor       =   &H8000000C&
   Caption         =   "VB Mail"
   ClientHeight    =   6090
   ClientLeft      =   375
   ClientTop       =   1995
   ClientWidth     =   9930
   Icon            =   "VBMAIL.frx":0000
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   9930
      TabIndex        =   0
      Top             =   5760
      Width           =   9930
      Begin VB.Line MsgBoxSide 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7260
         X2              =   7260
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line MsgBoxSide 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   60
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line MsgBoxLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   60
         X2              =   7260
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line MsgBoxLine 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   7260
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line TimeBoxSide 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   8580
         X2              =   8580
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line TimeBoxLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7320
         X2              =   8580
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line TimeBoxSide 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   7320
         X2              =   7320
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line TimeBoxLine 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   7320
         X2              =   8580
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   10800
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label UnreadLbl 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   60
         Width           =   1575
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   10800
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label TimeLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   205
         Left            =   7500
         TabIndex        =   10
         Top             =   75
         Width           =   345
      End
      Begin VB.Label MsgCountLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message Count Information"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   75
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9870
      TabIndex        =   8
      Top             =   525
      Visible         =   0   'False
      Width           =   9930
      Begin VB.Timer Timer1 
         Interval        =   15000
         Left            =   180
         Top             =   120
      End
      Begin VB.PictureBox MapiMess 
         Height          =   480
         Left            =   1320
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   15
         Top             =   120
         Width           =   1200
      End
      Begin VB.PictureBox MapiSess 
         Height          =   480
         Left            =   720
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   16
         Top             =   120
         Width           =   1200
      End
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   1920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "These controls are invisible at run time."
         Height          =   315
         Left            =   2700
         TabIndex        =   9
         Top             =   300
         Width           =   2835
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   9930
      TabIndex        =   12
      Top             =   0
      Width           =   9930
      Begin VB.CommandButton Command1 
         Caption         =   "Check Mail"
         Enabled         =   0   'False
         Height          =   330
         Left            =   8640
         TabIndex        =   14
         ToolTipText     =   "Check New Mail"
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton Delete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4980
         TabIndex        =   4
         ToolTipText     =   "Delete Message"
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton Next 
         Caption         =   "&Next"
         Enabled         =   0   'False
         Height          =   330
         Left            =   7440
         TabIndex        =   6
         ToolTipText     =   "Next Message"
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton Previous 
         Caption         =   "&Previous"
         Enabled         =   0   'False
         Height          =   330
         Left            =   6420
         TabIndex        =   5
         ToolTipText     =   "Previous Message"
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton SendCtl 
         Caption         =   "&Forward"
         Enabled         =   0   'False
         Height          =   330
         Index           =   9
         Left            =   3600
         TabIndex        =   3
         ToolTipText     =   "Foward Message"
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton SendCtl 
         Caption         =   "Reply &All"
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   2580
         TabIndex        =   2
         ToolTipText     =   "Reply To all"
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton SendCtl 
         Caption         =   "&Reply"
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   1560
         TabIndex        =   1
         ToolTipText     =   "Reply"
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton SendCtl 
         Caption         =   "&Compose"
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   180
         TabIndex        =   13
         ToolTipText     =   "New Message"
         Top             =   90
         Width           =   1035
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   15
         X2              =   10800
         Y1              =   505
         Y2              =   505
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   540
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   0
         X2              =   10800
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu PrintMessage 
         Caption         =   "&Print Message"
         Enabled         =   0   'False
      End
      Begin VB.Menu PrSetup 
         Caption         =   "Prin&ter Setup..."
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu EditDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Mail 
      Caption         =   "&Mail"
      Begin VB.Menu Logon 
         Caption         =   "Lo&gon"
      End
      Begin VB.Menu LogOff 
         Caption         =   "Log&off"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu rMsgList 
         Caption         =   "Update Message List"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu MailOpts 
         Caption         =   "&Mail..."
      End
      Begin VB.Menu FontS 
         Caption         =   "&Fonts"
         Begin VB.Menu FontScreen 
            Caption         =   "&Screen..."
         End
         Begin VB.Menu FontPrt 
            Caption         =   "&Printer..."
         End
      End
      Begin VB.Menu DispTools 
         Caption         =   "&Display Tools"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Address 
      Caption         =   "&Address"
      Begin VB.Menu ShowAB 
         Caption         =   "Show Address Book"
      End
   End
   Begin VB.Menu Window 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu wa 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu wa 
         Caption         =   "Tile Horizontally"
         Index           =   1
      End
      Begin VB.Menu wa 
         Caption         =   "Tile Vertically"
         Index           =   2
      End
      Begin VB.Menu wa 
         Caption         =   "Arrange Icons"
         Index           =   3
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "VBMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
    MsgBox "Mail Application made by Rizky Khapidsyah, and Microsoft's help on making this.", 0, "VB Mail"
End Sub

Private Sub Command1_Click()
Screen.MousePointer = 11
        GetMessageCount
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
End Sub

Private Sub Delete_Click()
' Delete a mail message.

    ' View all selected messages that are deleted.
    If TypeOf VBMail.ActiveForm Is MsgView Then
        Call DeleteMessage
    ElseIf TypeOf VBMail.ActiveForm Is MailLst Then
        ' Delete multiple selection.
        VBMail.MapiMess.MsgIndex = MailLst.MList.ListIndex
        Call DeleteMessage
    End If

End Sub

Private Sub DispTools_Click()
    DispTools.Checked = Not DispTools.Checked
    MailLst.Tools.Visible = DispTools.Checked

    
    If MailLst.Tools.Visible Then
        Factor = 1
        ToolsSize% = -MailLst.Tools.Height
    Else
        Factor = -1
        ToolsSize% = 0
    End If

    Select Case MailLst.WindowState
        Case 0    ' Change the size of the form to reflect the addition or deletion of a toolbar.
            MailLst.Height = MailLst.Height + (Factor * MailLst.Tools.Height)
        Case 2    ' If maximized, adjust the size of the list box.
            MailLst.MList.Height = ScaleHeight - 90 - MailLst.MList.Top + ToolsSize%
    End Select
End Sub

Private Sub EditDelete_Click()
' Delete the items in the list.
On Error GoTo Trap
    If TypeOf VBMail.ActiveForm Is MailLst Then
        Call Delete_Click
    End If
    Exit Sub

Trap:
    ' If an error occurs, there is probably no active form.
    ' Exit the Sub procedure.
    Exit Sub
End Sub

Private Sub Exit_Click()
    ' Close the application and log off.
    If MapiSess.SessionID <> 0 Then
        Call logoff_Click
    End If
    End
End Sub

Private Sub FontPrt_Click()
    ' Set the printer fonts.
    On Error Resume Next
    CMDialog1.Flags = 2
    CMDialog1.FontName = Printer.FontName
    CMDialog1.FontSize = Printer.FontSize
    CMDialog1.FontBold = Printer.FontBold
    CMDialog1.FontItalic = Printer.FontItalic
    CMDialog1.ShowFont
    If Err = 0 Then
        Printer.FontName = CMDialog1.FontName
        Printer.FontSize = CMDialog1.FontSize
        Printer.FontBold = CMDialog1.FontBold
        Printer.FontItalic = CMDialog1.FontItalic
    End If

End Sub

Private Sub FontScreen_Click()
    ' Set the screen fonts for the active control.
    On Error Resume Next
    CMDialog1.Flags = 1
    CMDialog1.FontName = VBMail.ActiveForm.ActiveControl.FontName
    CMDialog1.FontSize = VBMail.ActiveForm.ActiveControl.FontSize
    CMDialog1.FontBold = VBMail.ActiveForm.ActiveControl.FontBold
    CMDialog1.FontItalic = VBMail.ActiveForm.ActiveControl.FontItalic
    CMDialog1.ShowFont
    If Err = 0 Then
        VBMail.ActiveForm.ActiveControl.FontName = CMDialog1.FontName
        VBMail.ActiveForm.ActiveControl.FontSize = CMDialog1.FontSize
        VBMail.ActiveForm.ActiveControl.FontBold = CMDialog1.FontBold
        VBMail.ActiveForm.ActiveControl.FontItalic = CMDialog1.FontItalic
    End If
End Sub

Private Sub logoff_Click()
    ' Log off from the mail system.
    Call LogOffUser
End Sub

Private Sub Logon_Click()
    ' Log onto the mail system.
    On Error Resume Next
    MapiSess.Action = 1
    If Err <> 0 Then
        MsgBox "Logon Failure: " + Error$
    Else
        Screen.MousePointer = 11
        MapiMess.SessionID = MapiSess.SessionID
        ' Get the message count.
        GetMessageCount
        ' Load the mail list with envelope information.
        Screen.MousePointer = 11
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
        ' Adjust the buttons as needed.
        Logon.Enabled = False
        LogOff.Enabled = True
        VBMail.SendCtl(vbMessageCompose).Enabled = True
        VBMail.SendCtl(vbMessageReplyAll).Enabled = True
        VBMail.SendCtl(vbMessageReply).Enabled = True
        VBMail.SendCtl(vbMessageForward).Enabled = True
        VBMail.PrintMessage.Enabled = True
        VBMail.DispTools.Enabled = True
        VBMail.rMsgList.Enabled = True
        
        VBMail.Command1.Enabled = True
        
        VBMail.EditDelete.Enabled = True
      End If
End Sub

Private Sub MailOpts_Click()
    ' Display the Mail Options form.
    OptionType = conOptionGeneral
    MailOptFrm.Show 1
End Sub

Private Sub MDIForm_Load()
    ' Ensure all the controls are sized as needed.
    TimeLbl = Time$
     SendWithMapi = True
     Call Picture1_Resize
     Call Picture2_Resize
     VBMail.MsgCountLbl = "Off Line"
End Sub

Private Sub Next_Click()
    ' View the next message in the list.
    If MailLst.MList.ListIndex <> MailLst.MList.ListCount - 1 Then
        MailLst.MList.ItemData(MailLst.MList.ListIndex) = False
        MailLst.MList.ListIndex = MailLst.MList.ListIndex + 1
    End If
    Call ViewNextMsg
End Sub

Private Sub Picture1_Resize()
Const TimeBoxStartOffset = 1200
Const TimeBoxEndOffset = 60
Const MsgBoxStartOffset = 60
Const MsgBoxEndOffset = TimeBoxStartOffset + 90

    ' Adjust the sizes of the lines and position the time label.
    TimeLbl.Left = Picture1.Width - TimeLbl.Width - 265
    TopLine2(0).X2 = Picture1.Width
    TopLine2(1).X2 = Picture1.Width

    TimeBoxLine(0).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxLine(0).X2 = Picture1.Width - TimeBoxEndOffset

    TimeBoxLine(1).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxLine(1).X2 = Picture1.Width - TimeBoxEndOffset

    TimeBoxSide(0).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxSide(0).X2 = Picture1.Width - TimeBoxStartOffset

    TimeBoxSide(1).X1 = Picture1.Width - TimeBoxEndOffset
    TimeBoxSide(1).X2 = Picture1.Width - TimeBoxEndOffset

    MsgBoxLine(0).X2 = Picture1.Width - MsgBoxEndOffset
    MsgBoxLine(1).X2 = Picture1.Width - MsgBoxEndOffset

    MsgBoxSide(1).X1 = Picture1.Width - MsgBoxEndOffset
    MsgBoxSide(1).X2 = Picture1.Width - MsgBoxEndOffset

    Picture1.Refresh
End Sub

Private Sub Picture2_Resize()
    ' Adjust the positions of the lines.
    TopLine(0).X2 = Picture2.Width
    TopLine(1).X2 = Picture2.Width
    Picture2.Refresh
End Sub

Private Sub Previous_Click()
    ' View the previous message in the list.
    If MailLst.MList.ListIndex <> 0 Then
        MailLst.MList.ItemData(MailLst.MList.ListIndex) = False
        MailLst.MList.ListIndex = MailLst.MList.ListIndex - 1
    End If
    Call ViewNextMsg
End Sub

Private Sub PrintMessage_Click()
    ' Print mail.
    Call PrintMail
End Sub

Private Sub PrSetup_Click()
' Call the printer setup procedure in the common dialog control.
On Error Resume Next
    CMDialog1.Flags = &H40  ' Printer setup dialog box only.
    CMDialog1.ShowPrinter
End Sub

Private Sub rMsgList_Click()
        Screen.MousePointer = 11
        GetMessageCount
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
End Sub

Private Sub SendCtl_Click(Index As Integer)

NewMsg.Show

End Sub

Private Sub ShowAB_Click()
On Error Resume Next
    ' Show the address for the current message.
    VBMail.MapiMess.Action = vbMessageShowAdBook
    If Err Then
        If Err <> 32001 Then        ' User chose Cancel.
            MsgBox "Error: " + Error$ + " occurred trying to show the Address Book"
        End If
    Else
        If TypeOf VBMail.ActiveForm Is NewMsg Then
            Call UpdateRecips(VBMail.ActiveForm)
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    ' Update the time label.
    TimeLbl = Time$
End Sub

Private Sub wa_Click(Index As Integer)
    ' Arrange the windows as selected.
    VBMail.Arrange Index
End Sub

