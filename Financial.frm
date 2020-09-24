VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Financial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6360
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
   MouseIcon       =   "Financial.frx":0000
   Picture         =   "Financial.frx":030A
   ScaleHeight     =   5625
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   885
      Picture         =   "Financial.frx":75A68
      ScaleHeight     =   3495
      ScaleWidth      =   4560
      TabIndex        =   12
      Top             =   1050
      Visible         =   0   'False
      Width           =   4590
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   2
         Top             =   2295
         Width           =   2895
      End
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1320
         Width           =   3855
      End
      Begin VB.PictureBox picCancel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2400
         MouseIcon       =   "Financial.frx":A98BA
         MousePointer    =   99  'Custom
         Picture         =   "Financial.frx":A9A0C
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   17
         Top             =   2760
         Width           =   1515
         Begin VB.Label lblCancel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel"
            Height          =   210
            Left            =   495
            MouseIcon       =   "Financial.frx":ABCBE
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   105
            Width           =   495
         End
      End
      Begin VB.PictureBox picSave 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   720
         MouseIcon       =   "Financial.frx":ABE10
         MousePointer    =   99  'Custom
         Picture         =   "Financial.frx":ABF62
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   15
         Top             =   2760
         Width           =   1515
         Begin VB.Label lblSave 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
            Height          =   210
            Left            =   600
            MouseIcon       =   "Financial.frx":AE214
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   105
            Width           =   390
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         ItemData        =   "Financial.frx":AE366
         Left            =   840
         List            =   "Financial.frx":AE370
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   210
         Left            =   360
         TabIndex        =   19
         Top             =   2347
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         Height          =   210
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   210
         Left            =   360
         TabIndex        =   13
         Top             =   660
         Width           =   360
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   240
      Picture         =   "Financial.frx":AE387
      ScaleHeight     =   4215
      ScaleWidth      =   5895
      TabIndex        =   5
      Top             =   1080
      Width           =   5895
      Begin VB.PictureBox picView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2640
         MouseIcon       =   "Financial.frx":10457D
         MousePointer    =   99  'Custom
         Picture         =   "Financial.frx":1046CF
         ScaleHeight     =   435
         ScaleWidth      =   1275
         TabIndex        =   20
         Top             =   3720
         Width           =   1275
         Begin VB.Label lblView 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&View"
            Height          =   210
            Left            =   405
            MouseIcon       =   "Financial.frx":106411
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   105
            Width           =   390
         End
      End
      Begin VB.PictureBox picClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   4320
         MouseIcon       =   "Financial.frx":106563
         MousePointer    =   99  'Custom
         Picture         =   "Financial.frx":1066B5
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   10
         Top             =   3720
         Width           =   1515
         Begin VB.Label lblClose 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
            Height          =   210
            Left            =   525
            MouseIcon       =   "Financial.frx":108967
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   105
            Width           =   405
         End
      End
      Begin VB.PictureBox picDel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   1320
         MouseIcon       =   "Financial.frx":108AB9
         MousePointer    =   99  'Custom
         Picture         =   "Financial.frx":108C0B
         ScaleHeight     =   435
         ScaleWidth      =   1275
         TabIndex        =   8
         Top             =   3720
         Width           =   1275
         Begin VB.Label lblDel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Delete"
            Height          =   210
            Left            =   405
            MouseIcon       =   "Financial.frx":10A94D
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   105
            Width           =   450
         End
      End
      Begin VB.PictureBox picAdd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         MouseIcon       =   "Financial.frx":10AA9F
         MousePointer    =   99  'Custom
         Picture         =   "Financial.frx":10ABF1
         ScaleHeight     =   435
         ScaleWidth      =   1275
         TabIndex        =   6
         Top             =   3720
         Width           =   1275
         Begin VB.Label lblAdd 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Add"
            Height          =   210
            Left            =   405
            MouseIcon       =   "Financial.frx":10C933
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   105
            Width           =   420
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Financial.frx":10CA85
         Height          =   3495
         Left            =   0
         OleObjectBlob   =   "Financial.frx":10CA99
         TabIndex        =   3
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Planner"
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
      Width           =   1380
   End
End
Attribute VB_Name = "Financial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBGrid1_Click()

Data1.Refresh
Data2.Refresh

Data1.Recordset.AbsolutePosition = Data2.Recordset.AbsolutePosition

End Sub

Private Sub DBGrid1_SelChange(Cancel As Integer)

Cancel = 1

End Sub

Private Sub Form_Load()

With Data1

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM financial WHERE UserName='" _
                & Main.Tag & "'"
    .Refresh

End With

With Data2

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT financial.Type,financial.Notes,financial.Amount FROM financial WHERE UserName='" _
                & Main.Tag & "'"
    .Refresh

End With

End Sub

Private Sub lblAdd_Click()

Call Add_Click

End Sub

Private Sub lblCancel_Click()

Call Cancel_Click

End Sub

Private Sub lblClose_Click()

Unload Me

End Sub

Private Sub lblDel_Click()

Call Delete_Click

End Sub

Private Sub lblSave_Click()

Call Save_Click

End Sub

Private Sub lblView_Click()

Call View_Click

End Sub

Private Sub picAdd_Click()

Call Add_Click

End Sub

Private Sub picCancel_Click()

Call Cancel_Click

End Sub

Private Sub picClose_Click()

Unload Me

End Sub

Private Sub picDel_Click()

Call Delete_Click

End Sub

Private Sub picSave_Click()

Call Save_Click

End Sub

Private Sub picView_Click()

Call View_Click

End Sub

Private Sub txtAmount_GotFocus()

SendKeys "{home}+{end}"

End Sub

'txtAmount.Text = Format(Val(txtAmount), "P #,###,###,##0.00")

Private Sub txtAmount_KeyPress(KeyAscii As Integer)

If InStr(1, txtAmount.Text, ".") > 0 Then
    KeyAscii = 0
    Exit Sub
End If

Select Case KeyAscii
    Case 8
    Case 46, 48 To 57
    Case Else
        KeyAscii = 0
End Select

End Sub

Sub Add_Click()

picForm.Visible = True
picGrid.Enabled = False

Data1.Recordset.AddNew

txtNotes.Text = ""
txtAmount.Text = ""

Combo1.ListIndex = -1
Combo1.SetFocus

End Sub

Sub Delete_Click()

With Data1.Recordset

    If .RecordCount = 0 Then
        MsgBox "Financial Planner is empty.", vbExclamation, "Message"
        Exit Sub
    End If

    NoteTitle = .Fields("Type")
    ans = MsgBox("Are you sure?", vbQuestion + vbYesNo, _
            "Confirm")
    
    If ans = vbNo Then Exit Sub
    
    .Delete
    MsgBox "Deleted successfully.", , "Message"
    
    If .RecordCount <> 0 Then .MoveFirst
    
End With

Data1.Refresh
Data2.Refresh
DBGrid1.Refresh

End Sub

Sub Save_Click()

If Combo1.Text = "" Then
    MsgBox "Please select a type.", vbInformation, "Financial Planner"
    Combo1.SetFocus
    Exit Sub
End If

If txtNotes.Text = "" Then
    MsgBox "Please type a notes.", vbInformation, "Financial Planner"
    txtNotes.SetFocus
    Exit Sub
End If

If txtAmount.Text = "" Then
    MsgBox "Please type the amount.", vbInformation, "Financial Planner"
    txtAmount.SetFocus
    Exit Sub
End If
    
With Data1.Recordset

    .Fields("UserName") = Main.Tag
    .Fields("Type") = Combo1.Text
    .Fields("Notes") = txtNotes.Text
    .Fields("Amount") = Format(Val(txtAmount), "P #,###,###,##0.00")
    
    .Update

End With

MsgBox Combo1.Text & " has been saved.", vbInformation, "Financial Planenr"
Data1.Refresh
Data2.Refresh

DBGrid1.Refresh

picForm.Visible = False
picGrid.Enabled = True

End Sub

Sub Cancel_Click()

picForm.Visible = False
picGrid.Enabled = True

Combo1.Enabled = True
txtNotes.Locked = False
txtAmount.Locked = False

picSave.Visible = True
lblCancel.Caption = "Cancel"

End Sub

Sub View_Click()

On Error Resume Next

With Data1.Recordset
        
    If .RecordCount = 0 And .EOF = True Then Exit Sub
    
    .MoveLast
    .AbsolutePosition = Data2.Recordset.AbsolutePosition
    
    If .Fields("Type") = Combo1.List(0) Then
        Combo1.ListIndex = 0
    Else
        Combo1.ListIndex = 1
    End If
    
    txtNotes.Text = .Fields("Notes")
    txtAmount.Text = .Fields("Amount")

End With

Combo1.Enabled = False
txtNotes.Locked = True
txtAmount.Locked = True

picSave.Visible = False
lblCancel.Caption = "Close"

picForm.Visible = True
picGrid.Enabled = False

End Sub

Private Sub txtNotes_GotFocus()

SendKeys "{home}+{end}"

End Sub
