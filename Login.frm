VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4590
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
   MouseIcon       =   "Login.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Login.frx":030A
   ScaleHeight     =   3525
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picQuit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3000
      MouseIcon       =   "Login.frx":3593
      MousePointer    =   99  'Custom
      Picture         =   "Login.frx":36E5
      ScaleHeight     =   435
      ScaleWidth      =   1230
      TabIndex        =   9
      Top             =   2640
      Width           =   1230
      Begin VB.Label lblQuit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Quit"
         Height          =   210
         Left            =   480
         MouseIcon       =   "Login.frx":533F
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.PictureBox picOK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1620
      MouseIcon       =   "Login.frx":5491
      MousePointer    =   99  'Custom
      Picture         =   "Login.frx":55E3
      ScaleHeight     =   435
      ScaleWidth      =   1230
      TabIndex        =   7
      Top             =   2640
      Width           =   1230
      Begin VB.Label lblOK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&OK"
         Height          =   210
         Left            =   480
         MouseIcon       =   "Login.frx":723D
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   120
         Width           =   225
      End
   End
   Begin VB.PictureBox picSign 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   240
      MouseIcon       =   "Login.frx":738F
      MousePointer    =   99  'Custom
      Picture         =   "Login.frx":74E1
      ScaleHeight     =   435
      ScaleWidth      =   1230
      TabIndex        =   5
      Top             =   2640
      Width           =   1230
      Begin VB.Label lblSign 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sign-in"
         Height          =   210
         Left            =   360
         MouseIcon       =   "Login.frx":913B
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      MaxLength       =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Top             =   1245
      Width           =   2775
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot your password? Click here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1440
      MouseIcon       =   "Login.frx":928D
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2280
      Width           =   2835
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
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
      Width           =   900
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   210
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   210
      Left            =   480
      TabIndex        =   2
      Top             =   1755
      Width           =   750
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

With Data1

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = "Users"
    .Refresh

End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If (X >= 0 And Y >= 0) And (X <= 4200 And Y <= 460) Then
    If Button = vbLeftButton Then Call DragIt(Me.hwnd)
End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblHint_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
Hint_Click

End Sub

Private Sub lblOK_Click()

Ok_Click

End Sub

Private Sub lblQuit_Click()

Cancel_Click

End Sub

Private Sub lblSign_Click()

Create_Click

End Sub

Private Sub picOK_Click()

Ok_Click

End Sub

Private Sub picQuit_Click()

Cancel_Click

End Sub

Private Sub picSign_Click()

Create_Click

End Sub

Private Sub txtPass_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then Ok_Click

End Sub

Private Sub txtUser_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then txtPass.SetFocus
If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0

End Sub

Sub Cancel_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

ans = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Quit E-Diary 2008?")

If ans = vbYes Then End

End Sub

Sub Ok_Click()

If txtUser.Text = Empty Then
    
    sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
    
    MsgBox "Please type your User Name.", vbExclamation, "Required"
    txtUser.SetFocus
    Exit Sub

ElseIf txtPass.Text = Empty Then
    
    sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
    
    MsgBox "Please type your Password.", vbExclamation, "Required"
    txtPass.SetFocus
    Exit Sub

End If

With Data1.Recordset
    
    .MoveFirst
    .FindFirst "User='" & txtUser.Text & "'"
    
    If .NoMatch = True Then
    
        sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
        
        MsgBox "User Name not found.", vbExclamation, "Message"
        txtUser.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    
    End If

    If txtPass.Text <> .Fields("Password") Then
    
        sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
        
        MsgBox "Access Denied.", vbCritical, "Message"
        txtUser.Text = ""
        txtPass.Text = ""
        txtUser.SetFocus
        Exit Sub
    
    End If
        
    sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
    
    MsgBox "Access Granted.", vbInformation, "Message"
    Unload Me
    
    Main.picPick.Visible = True
    Main.lblUser = "User - [" & UCase(.Fields("Name")) & "]"
    Main.Tag = .Fields("User")
    
    RemindChecker.Show
    RemindChecker.Hide
    
End With

End Sub

Sub Create_Click()

Unload Me

Load Create
Create.Show vbModal

End Sub

Sub Hint_Click()

uname = InputBox("Enter your Username", "Password Hint", "Type here!!")

If uname = "Type here!!" Or uname = Empty Then Exit Sub

'check for the single quote (')
For i = 1 To Len(uname)
    If Mid(uname, i, 1) = "'" Then
        MsgBox "Do not use a single quote.", vbExclamation, "Message"
        Exit Sub
    ElseIf Mid(uname, i, 1) = """" Then
        MsgBox "Do not use a double quote.", vbExclamation, "Message"
        Exit Sub
    End If
Next

With Data1.Recordset

    If .RecordCount = 0 And .EOF = True Then
        MsgBox "Record is empty.", vbCritical, "Message"
        Exit Sub
    End If

    .MoveFirst
    .FindFirst "User='" & uname & "'"

    If .NoMatch = True Then
        MsgBox "Invalid Username.", , "Message"
        Exit Sub
    End If
    
    MsgBox "Hint: " & .Fields("Hint"), vbQuestion, "Password Hint"

End With

End Sub
