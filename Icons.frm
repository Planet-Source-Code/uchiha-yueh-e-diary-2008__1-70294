VERSION 5.00
Begin VB.Form Icons 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2520
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Icons.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Icons.frx":030A
   ScaleHeight     =   7665
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picApply 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   398
      MouseIcon       =   "Icons.frx":3F754
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":3F8A6
      ScaleHeight     =   435
      ScaleWidth      =   1725
      TabIndex        =   27
      Top             =   7080
      Visible         =   0   'False
      Width           =   1725
      Begin VB.Label lblApply 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   210
         Left            =   630
         MouseIcon       =   "Icons.frx":42054
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   120
         Width           =   435
      End
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   25
      Left            =   930
      MouseIcon       =   "Icons.frx":421A6
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":422F8
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   26
      ToolTipText     =   "Surprised"
      Top             =   6240
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   24
      Left            =   225
      MouseIcon       =   "Icons.frx":435FA
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":4374C
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   25
      ToolTipText     =   "Laughing"
      Top             =   6240
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   23
      Left            =   1665
      MouseIcon       =   "Icons.frx":44A4E
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":44BA0
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   24
      ToolTipText     =   "Ashamed"
      Top             =   5520
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   22
      Left            =   960
      MouseIcon       =   "Icons.frx":45EA2
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":45FF4
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   23
      ToolTipText     =   "Embarrassed"
      Top             =   5520
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   21
      Left            =   225
      MouseIcon       =   "Icons.frx":472F6
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":47448
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   22
      ToolTipText     =   "Love"
      Top             =   5520
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   20
      Left            =   1665
      MouseIcon       =   "Icons.frx":4874A
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":4889C
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   21
      ToolTipText     =   "Mad"
      Top             =   4800
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   19
      Left            =   960
      MouseIcon       =   "Icons.frx":49B9E
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":49CF0
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   20
      ToolTipText     =   "Friendly"
      Top             =   4800
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   18
      Left            =   225
      MouseIcon       =   "Icons.frx":4AFF2
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":4B144
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   19
      ToolTipText     =   "Sorry"
      Top             =   4800
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   17
      Left            =   1665
      MouseIcon       =   "Icons.frx":4C446
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":4C598
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   18
      ToolTipText     =   "Happy"
      Top             =   4080
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   16
      Left            =   960
      MouseIcon       =   "Icons.frx":4D89A
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":4D9EC
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   17
      ToolTipText     =   "Exhausted"
      Top             =   4080
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   15
      Left            =   225
      MouseIcon       =   "Icons.frx":4ECEE
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":4EE40
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   16
      ToolTipText     =   "Congratulate"
      Top             =   4080
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   14
      Left            =   1665
      MouseIcon       =   "Icons.frx":50142
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":50294
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   15
      ToolTipText     =   "Questioning"
      Top             =   3360
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   13
      Left            =   960
      MouseIcon       =   "Icons.frx":51596
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":516E8
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   14
      ToolTipText     =   "Confused"
      Top             =   3360
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   12
      Left            =   225
      MouseIcon       =   "Icons.frx":529EA
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":52B3C
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   13
      ToolTipText     =   "Liar"
      Top             =   3360
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   11
      Left            =   1665
      MouseIcon       =   "Icons.frx":53E3E
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":53F90
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   12
      ToolTipText     =   "Screaming"
      Top             =   2640
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   10
      Left            =   960
      MouseIcon       =   "Icons.frx":55292
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":553E4
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   11
      ToolTipText     =   "Mocking"
      Top             =   2640
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   9
      Left            =   225
      MouseIcon       =   "Icons.frx":566E6
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":56838
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   10
      ToolTipText     =   "Chagrined"
      Top             =   2640
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   8
      Left            =   1665
      MouseIcon       =   "Icons.frx":57B3A
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":57C8C
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   9
      ToolTipText     =   "Hurt"
      Top             =   1920
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   7
      Left            =   960
      MouseIcon       =   "Icons.frx":58F8E
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":590E0
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   8
      ToolTipText     =   "Smart"
      Top             =   1920
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   6
      Left            =   225
      MouseIcon       =   "Icons.frx":5A3E2
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":5A534
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   7
      ToolTipText     =   "Upset"
      Top             =   1920
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   5
      Left            =   1665
      MouseIcon       =   "Icons.frx":5B836
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":5B988
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   6
      ToolTipText     =   "Tearful"
      Top             =   1200
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   4
      Left            =   960
      MouseIcon       =   "Icons.frx":5CC8A
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":5CDDC
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   5
      ToolTipText     =   "Ecstatic"
      Top             =   1200
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   3
      Left            =   225
      MouseIcon       =   "Icons.frx":5E0DE
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":5E230
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   4
      ToolTipText     =   "Scared"
      Top             =   1200
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   2
      Left            =   1665
      MouseIcon       =   "Icons.frx":5F532
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":5F684
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   3
      ToolTipText     =   "Perplex"
      Top             =   480
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   1
      Left            =   960
      MouseIcon       =   "Icons.frx":60986
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":60AD8
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   2
      ToolTipText     =   "Yawning"
      Top             =   480
      Width           =   630
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   0
      Left            =   225
      MouseIcon       =   "Icons.frx":61DDA
      MousePointer    =   99  'Custom
      Picture         =   "Icons.frx":61F2C
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   1
      ToolTipText     =   "Angry"
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Icons"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   450
   End
End
Attribute VB_Name = "Icons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IconNo As Integer

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If (X >= 0 And Y >= 0) And (X <= 2160 And Y <= 360) Then
    If Button = vbLeftButton Then Call DragIt(Me.hwnd)
End If

End Sub

Private Sub lblApply_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

picApply.Move picApply.Left + 30, picApply.Top + 30

End Sub

Private Sub lblApply_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

picApply.Move 398, 7080
Apply_Click

End Sub

Private Sub picApply_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

picApply.Move picApply.Left + 30, picApply.Top + 30

End Sub

Private Sub picApply_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

picApply.Move 398, 7080
Apply_Click

End Sub

Private Sub picIcons_Click(Index As Integer)

Reset_Icons
picIcons(Index).Appearance = 1

lblApply.Enabled = True
picApply.Enabled = True

IconNo = Index

Apply_Click

End Sub

Sub Reset_Icons()

For i = 0 To 25
    picIcons(i).Appearance = 0
Next

End Sub

Sub Apply_Click()
Dim txt As String

txt = " [" & CStr(Chr(IconNo + 97)) & "] "

Clipboard.SetText txt

WriteDiary.txtDiary.SetFocus

WriteDiary.txtDiary.SelText = Clipboard.GetText

'WriteDiary.txtDiary.Text = _
'    WriteDiary.txtDiary.Text & " [" & CStr(Chr(IconNo + 97)) & "] "
'
'WriteDiary.txtDiary.SetFocus
'SendKeys "{end}"

End Sub
