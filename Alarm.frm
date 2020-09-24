VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "Fm20.dll"
Begin VB.Form Alarm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4560
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
   MouseIcon       =   "Alarm.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Alarm.frx":030A
   ScaleHeight     =   3495
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3720
      Tag             =   "OFF"
      Top             =   1560
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2363
      MouseIcon       =   "Alarm.frx":3415C
      MousePointer    =   99  'Custom
      Picture         =   "Alarm.frx":342AE
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   2640
      Width           =   1515
      Begin VB.Label lblCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Close"
         Height          =   210
         Left            =   480
         MouseIcon       =   "Alarm.frx":36560
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   105
         Width           =   405
      End
   End
   Begin VB.PictureBox picApply 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   683
      MouseIcon       =   "Alarm.frx":366B2
      MousePointer    =   99  'Custom
      Picture         =   "Alarm.frx":36804
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   2640
      Width           =   1515
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Apply"
         Height          =   210
         Left            =   520
         MouseIcon       =   "Alarm.frx":38AB6
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   105
         Width           =   420
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3720
      Tag             =   "OFF"
      Top             =   1080
   End
   Begin VB.ComboBox cboHr 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1575
      Width           =   615
   End
   Begin VB.ComboBox cboMin 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1935
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      Picture         =   "Alarm.frx":38C08
      ScaleHeight     =   855
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   600
      Width           =   750
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm is Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin MSForms.OptionButton optPM 
      Height          =   360
      Left            =   2640
      TabIndex        =   8
      Top             =   1920
      Width           =   615
      VariousPropertyBits=   1015023633
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "1085;635"
      Value           =   "0"
      Caption         =   "PM"
      FontName        =   "Arial"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton optAM 
      Height          =   360
      Left            =   2640
      TabIndex        =   7
      Top             =   1560
      Width           =   645
      VariousPropertyBits=   1015023633
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "1138;635"
      Value           =   "0"
      Caption         =   "AM"
      FontName        =   "Arial"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblMin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minute:"
      Height          =   210
      Left            =   1080
      TabIndex        =   6
      Top             =   1995
      Width           =   510
   End
   Begin VB.Label lblHr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hour:"
      Height          =   210
      Left            =   1200
      TabIndex        =   5
      Top             =   1635
      Width           =   390
   End
   Begin MSForms.CheckBox cboSet 
      Height          =   360
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   1305
      VariousPropertyBits=   1015023635
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2302;635"
      Value           =   "0"
      Caption         =   "Set an alarm"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current time is 00:00:00 AM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   2340
   End
End
Attribute VB_Name = "Alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboSet_Change()

cboHr.Enabled = Not cboHr.Enabled
cboMin.Enabled = Not cboMin.Enabled
optAM.Enabled = Not optAM.Enabled
optPM.Enabled = Not optPM.Enabled

End Sub

Private Sub Form_Load()

'for hour..
For i = 1 To 12
    cboHr.AddItem CStr(i)
Next

'for minutes..
For i = 0 To 59
    If i < 10 Then
        cboMin.AddItem "0" & CStr(i)
    Else
        cboMin.AddItem CStr(i)
    End If
Next

lblTime.Caption = "Current time is " & Time


End Sub

Private Sub lblApply_Click()

Apply_Click

End Sub

Private Sub lblCancel_Click()

Cancel_Click

End Sub

Private Sub picApply_Click()

Apply_Click

End Sub

Private Sub picCancel_Click()

Cancel_Click

End Sub

Private Sub Timer1_Timer()

lblTime.Caption = "Current time is " & Time

End Sub

Sub Apply_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

If cboSet.Value = False Then

    cboHr.ListIndex = -1
    cboMin.ListIndex = -1
    optAM.Value = False
    optPM.Value = False
    lblTitle.Caption = "Alarm is Off"
    
    Main.picPick.Tag = ""

    Unload Alarma
    Set Alarma = Nothing

    MsgBox "Alarm has been set off.", _
            vbInformation, _
            "Alarm is Off"
    
    Exit Sub

End If

'check alarm time..
If cboHr.Text = Empty Then
    MsgBox "Select a Hour.", , "Message"
    cboHr.SetFocus
    Exit Sub
ElseIf cboMin.Text = Empty Then
    MsgBox "Select a Minute.", , "Message"
    cboMin.SetFocus
    Exit Sub
ElseIf optAM.Value = False And optPM.Value = False Then
    MsgBox "Select if AM or PM.", , "Message"
    Exit Sub
End If

Main.picPick.Tag = cboHr.Text & ":" & cboMin.Text

If optAM.Value = True Then
    Main.picPick.Tag = Main.picPick.Tag & ":00 AM"
Else
    Main.picPick.Tag = Main.picPick.Tag & ":00 PM"
End If

MsgBox "Alarm has been set on " & Main.picPick.Tag, _
        vbInformation, _
        "Alarm is ON"

lblTitle.Caption = "Alarm is On - [" & Main.picPick.Tag & "]"

End Sub

Sub Cancel_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Me.Hide
Main.SetFocus

End Sub

Private Sub Timer2_Timer()

If Time = Main.picPick.Tag Then
    
    MakeTopMost Alarma.hwnd
    Alarma.Show

End If

End Sub

