VERSION 5.00
Begin VB.Form MsgView 
   Caption         =   "Mail Message"
   ClientHeight    =   4380
   ClientLeft      =   3150
   ClientTop       =   4800
   ClientWidth     =   6210
   Icon            =   "MSGVIEW.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   6210
   Begin VB.PictureBox AttachWin 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   6150
      TabIndex        =   12
      Top             =   3555
      Visible         =   0   'False
      Width           =   6210
      Begin VB.ListBox aList 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Width           =   5835
      End
      Begin VB.Label NumAtt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   60
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachments:"
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   930
      End
   End
   Begin VB.TextBox txtNoteText 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1740
      Width           =   6195
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   6150
      TabIndex        =   6
      Top             =   0
      Width           =   6210
      Begin VB.TextBox txtsubject 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   5
         Top             =   1320
         Width           =   4995
      End
      Begin VB.TextBox txtcc 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   4
         Top             =   1020
         Width           =   4995
      End
      Begin VB.TextBox txtTo 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   3
         Top             =   720
         Width           =   4995
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   2
         Top             =   420
         Width           =   4995
      End
      Begin VB.TextBox txtFrom 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   900
         TabIndex        =   1
         Top             =   120
         Width           =   4995
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cc:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   480
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H00FFFFFF&
         X1              =   6000
         X2              =   6000
         Y1              =   60
         Y2              =   1620
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H00FFFFFF&
         X1              =   60
         X2              =   6000
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00808080&
         X1              =   60
         X2              =   5940
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H00808080&
         X1              =   60
         X2              =   60
         Y1              =   60
         Y2              =   1620
      End
   End
End
Attribute VB_Name = "MsgView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aList_DblClick()
    ' ListIndex is the index into the attachment list.
    VBMail.MapiMess.AttachmentIndex = aList.ListIndex
    If VBMail.MapiMess.AttachmentType = vbAttachTypeData Then
        Call DisplayAttachedFile(VBMail.MapiMess.AttachmentPathName)
    Else
        MsgBox " This sample application doesn't view OLE-type attachments"
    End If
End Sub

Private Sub AttachWin_Resize()
    '  Update the widths of the fields and adjust the line
    '  controls as needed.
    aList.Width = AttachWin.Width - aList.Left - 315
End Sub

Private Sub Form_Activate()
    ' When the form is activated, update mailLst.MList
    ' to reflect the current item. The Tag property contains
    ' the index of the currently viewed message.
    MailLst.MList.ListIndex = Val(Me.Tag)
    MailLst.MList.ItemData(Val(Me.Tag)) = True
    VBMail.MapiMess.MsgIndex = Val(Me.Tag)
End Sub

Private Sub Form_Load()
    ' Ensure all resizing is done on startup.
    Call Picture1_Resize
    Call AttachWin_Resize
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    ' Adjust the window size if the form isn't minimized.
    Call SizeMessageWindow(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Tag is set to -1 after the currently viewed message
    ' is deleted.
    If Val(Me.Tag) > 0 Then
        MailLst.MList.ItemData(Val(Me.Tag)) = False
    End If
End Sub

Private Sub Picture1_Resize()
    '  Update the widths of the fields and adjust the line
    '  controls as needed.
    TopLine.X2 = Picture1.Width - 90
    BottomLine.X2 = Picture1.Width - 90
    RightLine.X1 = Picture1.Width - 90
    RightLine.X2 = Picture1.Width - 90
    lf% = txtTo.Left
    txtTo.Width = Picture1.Width - 120 - lf%
    txtDate.Width = Picture1.Width - 120 - lf%
    txtcc.Width = Picture1.Width - 120 - lf%
    txtsubject.Width = Picture1.Width - 120 - lf%
    txtFrom.Width = Picture1.Width - 120 - lf%
    Picture1.Refresh
End Sub

Private Sub txtcc_KeyPress(KeyAscii As Integer)
    ' Ignore all keystrokes.
    KeyAscii = 0
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    ' Ignore all keystrokes.
    KeyAscii = 0
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    ' Ignore all keystrokes.
    KeyAscii = 0
End Sub

Private Sub txtNoteText_KeyPress(KeyAscii As Integer)
    ' Ignore all keystrokes.
    KeyAscii = 0
End Sub

Private Sub txtsubject_KeyPress(KeyAscii As Integer)
    ' Ignore all keystrokes.
    KeyAscii = 0
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    ' Ignore all keystrokes.
    KeyAscii = 0
End Sub

