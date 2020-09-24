VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Reminders 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6375
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
   Picture         =   "Reminders.frx":0000
   ScaleHeight     =   5190
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1920
      Picture         =   "Reminders.frx":6B7DE
      ScaleHeight     =   1545
      ScaleWidth      =   2385
      TabIndex        =   24
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
      Begin VB.PictureBox picSet 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   360
         MouseIcon       =   "Reminders.frx":ACFE0
         MousePointer    =   99  'Custom
         Picture         =   "Reminders.frx":AD132
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   31
         Top             =   960
         Width           =   1515
         Begin VB.Label lblSet 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set"
            Height          =   210
            Left            =   585
            MouseIcon       =   "Reminders.frx":AF3E4
            MousePointer    =   99  'Custom
            TabIndex        =   32
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.ComboBox cboHr 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   135
         Width           =   615
      End
      Begin VB.ComboBox cboMin 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   495
         Width           =   615
      End
      Begin MSForms.OptionButton optPM 
         Height          =   360
         Left            =   1680
         TabIndex        =   30
         Top             =   480
         Width           =   615
         VariousPropertyBits=   1015023635
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1085;635"
         Value           =   "0"
         Caption         =   "PM"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optAM 
         Height          =   360
         Left            =   1680
         TabIndex        =   29
         Top             =   120
         Width           =   645
         VariousPropertyBits=   1015023635
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1138;635"
         Value           =   "0"
         Caption         =   "AM"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblMin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minute:"
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   555
         Width           =   510
      End
      Begin VB.Label lblHr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour:"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   195
         Width           =   390
      End
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      DataField       =   "Time"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.PictureBox picNew 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   4680
      MouseIcon       =   "Reminders.frx":AF536
      MousePointer    =   99  'Custom
      Picture         =   "Reminders.frx":AF688
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   17
      Top             =   1200
      Width           =   1515
      Begin VB.Label lblNew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&New Reminder"
         Height          =   210
         Left            =   240
         MouseIcon       =   "Reminders.frx":B193A
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.PictureBox picEdit 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   4680
      MouseIcon       =   "Reminders.frx":B1A8C
      MousePointer    =   99  'Custom
      Picture         =   "Reminders.frx":B1BDE
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   1800
      Width           =   1515
      Begin VB.Label lblEdit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Edit Reminder"
         Height          =   210
         Left            =   270
         MouseIcon       =   "Reminders.frx":B3E90
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   120
         Width           =   1005
      End
   End
   Begin VB.PictureBox picDelete 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   4680
      MouseIcon       =   "Reminders.frx":B3FE2
      MousePointer    =   99  'Custom
      Picture         =   "Reminders.frx":B4134
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   13
      Top             =   2400
      Width           =   1515
      Begin VB.Label lblDelete 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Delete"
         Height          =   210
         Left            =   540
         MouseIcon       =   "Reminders.frx":B63E6
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.PictureBox picPrev 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   360
      MouseIcon       =   "Reminders.frx":B6538
      MousePointer    =   99  'Custom
      Picture         =   "Reminders.frx":B668A
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   4440
      Width           =   1515
      Begin VB.Label lblPrev 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Previous"
         Height          =   210
         Left            =   390
         MouseIcon       =   "Reminders.frx":B893C
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   120
         Width           =   645
      End
   End
   Begin VB.PictureBox picNext 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   1920
      MouseIcon       =   "Reminders.frx":B8A8E
      MousePointer    =   99  'Custom
      Picture         =   "Reminders.frx":B8BE0
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   4440
      Width           =   1515
      Begin VB.Label lblNext 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ne&xt"
         Height          =   210
         Left            =   540
         MouseIcon       =   "Reminders.frx":BAE92
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   4680
      MouseIcon       =   "Reminders.frx":BAFE4
      MousePointer    =   99  'Custom
      Picture         =   "Reminders.frx":BB136
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   4440
      Width           =   1515
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         Height          =   210
         Left            =   495
         MouseIcon       =   "Reminders.frx":BD3E8
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   120
         Width           =   435
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtRemind 
      Appearance      =   0  'Flat
      DataField       =   "Remind About"
      DataSource      =   "Data1"
      Height          =   1215
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "Date"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2805
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   62586881
      CurrentDate     =   37977
   End
   Begin VB.PictureBox picCancel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   4680
      MouseIcon       =   "Reminders.frx":BD53A
      MousePointer    =   99  'Custom
      Picture         =   "Reminders.frx":BD68C
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   1515
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cancel"
         Height          =   210
         Left            =   510
         MouseIcon       =   "Reminders.frx":BF93E
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.PictureBox picSave 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   4680
      MouseIcon       =   "Reminders.frx":BFA90
      MousePointer    =   99  'Custom
      Picture         =   "Reminders.frx":BFBE2
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   1515
      Begin VB.Label lblSave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Save"
         Height          =   210
         Left            =   570
         MouseIcon       =   "Reminders.frx":C1E94
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      DataField       =   "Date"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   210
      Left            =   240
      TabIndex        =   23
      Top             =   3315
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remind About:"
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reminders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "Reminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DTPicker1_Change()

txtDate.Text = DTPicker1.Value

End Sub

Private Sub DTPicker1_Click()

txtDate.Text = DTPicker1.Value

End Sub

Private Sub DTPicker1_GotFocus()

txtDate.Text = DTPicker1.Value

End Sub

Private Sub Form_Load()

DTPicker1.Value = Date

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

With Data1

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = "SELECT * FROM reminder WHERE UserName='" & Main.Tag & _
                "' ORDER BY date,time"
    .Refresh

End With

End Sub

Sub Prev_Click()

On Error Resume Next
With Data1.Recordset
    
    .MovePrevious
    If .BOF = True Then .MoveFirst
    
End With

End Sub

Sub Next_Click()

On Error Resume Next
With Data1.Recordset
    
    .MoveNext
    If .EOF = True Then .MoveLast
    
End With

End Sub

Private Sub lblCancel_Click()

Cancel_Click

End Sub

Private Sub lblClose_Click()

Unload Me

End Sub

Private Sub lblDelete_Click()

Delete_Click

End Sub

Private Sub lblEdit_Click()

Edit_Click

End Sub

Private Sub lblNew_Click()

Add_Click

End Sub

Private Sub lblNext_Click()

Next_Click

End Sub

Private Sub lblPrev_Click()

Prev_Click

End Sub

Private Sub lblSave_Click()

Save_Click

End Sub

Private Sub lblSet_Click()

Set_Click

End Sub

Private Sub picCancel_Click()

Cancel_Click

End Sub

Private Sub picClose_Click()

Unload Me

End Sub

Private Sub picDelete_Click()

Delete_Click

End Sub

Private Sub picEdit_Click()

Edit_Click

End Sub

Private Sub picNew_Click()

Add_Click

End Sub

Private Sub picNext_Click()

Next_Click

End Sub

Private Sub picPrev_Click()

Prev_Click

End Sub

Sub Form_Controls()

picNew.Visible = Not picNew.Visible
picEdit.Visible = Not picEdit.Visible
picDelete.Visible = Not picDelete.Visible
picSave.Visible = Not picSave.Visible
picCancel.Visible = Not picCancel.Visible
picClose.Visible = Not picClose.Visible

picTime.Visible = Not picTime.Visible

txtRemind.Locked = Not txtRemind.Locked
DTPicker1.Visible = Not DTPicker1.Visible
txtDate.Visible = Not txtDate.Visible

picPrev.Visible = Not picPrev.Visible
picNext.Visible = Not picNext.Visible

End Sub

Sub Add_Click()

Call Form_Controls
Data1.Recordset.AddNew

End Sub

Sub Edit_Click()

With Data1.Recordset

    If .RecordCount = 0 And .EOF = True Then
        MsgBox "No reminder to be edited.", vbInformation, "Reminder"
        Exit Sub
    End If

    Call Form_Controls
    .Edit

End With

End Sub

Sub Delete_Click()

With Data1.Recordset

    If .RecordCount = 0 Then
        MsgBox "Reminder is empty.", vbExclamation, "Reminder"
        RemindChecker.Tag = ""
        Exit Sub
    End If

    NoteTitle = .Fields("Remind About")
    ans = MsgBox("Are you sure?", vbQuestion + vbYesNo, _
            "Delete '" & NoteTitle & "' ?")
    
    If ans = vbNo Then Exit Sub
    
    .Delete
    MsgBox NoteTitle & " has been deleted.", , "Message"
    
    If .RecordCount <> 0 Then .MoveFirst
    
End With

End Sub

Sub Save_Click()

With Data1.Recordset

    If txtRemind.Text = "" Then
        MsgBox "Please type your remineder.", vbInformation, "Reminder"
        txtRemind.SetFocus
        Exit Sub
    End If

    If txtDate.Text = "" Then
        MsgBox "Please select a date.", vbInformation, "Reminder"
        DTPicker1.SetFocus
        Exit Sub
    End If

    If txtTime.Text = "" Then
        MsgBox "Please set the hour and minute.", vbInformation, "Reminder"
        cboHr.SetFocus
        Exit Sub
    End If

    .Fields("UserName") = Main.Tag
    .Update
    
    MsgBox "Reminder has been saved.", vbInformation, "Reminder"

    RemindChecker.Tag = "REMIND"
End With

Data1.Refresh
Call Form_Controls

End Sub

Sub Cancel_Click()

Data1.Recordset.CancelUpdate
Call Form_Controls

End Sub

Sub Set_Click()

If cboHr.Text = "" Then
    MsgBox "Select hour.", vbInformation, "Reminder"
    cboHr.SetFocus
    Exit Sub
End If

If cboMin.Text = "" Then
    MsgBox "Select minute.", vbInformation, "Reminder"
    cboMin.SetFocus
    Exit Sub
End If

If optAM.Value = False And optPM.Value = False Then
    MsgBox "Select AM or PM.", vbInformation, "Reminder"
    Exit Sub
End If

txtTime.Text = cboHr.Text & ":" & cboMin.Text & ":00 " & IIf(optAM.Value = True, "AM", "PM")

End Sub

Private Sub picSave_Click()

Save_Click

End Sub

Private Sub picSet_Click()

Set_Click

End Sub

Private Sub txtTime_GotFocus()

SendKeys "{home}+{end}"

End Sub
