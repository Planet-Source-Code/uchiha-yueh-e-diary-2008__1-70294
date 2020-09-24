VERSION 5.00
Begin VB.Form Create 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5175
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
   Picture         =   "Create.frx":0000
   ScaleHeight     =   5175
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHint 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   50
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Tag             =   "Password Hint"
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Data dbUsers 
      Caption         =   "Users"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3345
      MouseIcon       =   "Create.frx":436C
      MousePointer    =   99  'Custom
      Picture         =   "Create.frx":44BE
      ScaleHeight     =   435
      ScaleWidth      =   1230
      TabIndex        =   17
      Top             =   4440
      Width           =   1230
      Begin VB.Label lblCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cancel"
         Height          =   210
         Left            =   360
         MouseIcon       =   "Create.frx":6118
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   120
         Width           =   495
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
      Left            =   1785
      MouseIcon       =   "Create.frx":626A
      MousePointer    =   99  'Custom
      Picture         =   "Create.frx":63BC
      ScaleHeight     =   435
      ScaleWidth      =   1230
      TabIndex        =   15
      Top             =   4440
      Width           =   1230
      Begin VB.Label lblSign 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sign-in"
         Height          =   210
         Left            =   360
         MouseIcon       =   "Create.frx":8016
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox txtConfirm 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2393
      MaxLength       =   15
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   6
      Tag             =   "Confirm Password"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2393
      MaxLength       =   15
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   5
      Tag             =   "Password"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtCel 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2393
      MaxLength       =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Tag             =   "Mobile"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtTel 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2393
      MaxLength       =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Tag             =   "Home Phone"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2393
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Tag             =   "Home Address"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2393
      MaxLength       =   50
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2393
      MaxLength       =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Tag             =   "User Name"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Hint:"
      Height          =   210
      Left            =   960
      TabIndex        =   20
      Top             =   4035
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sign-in (New User)"
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
      TabIndex        =   19
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label lblConfirm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      Height          =   210
      Left            =   720
      TabIndex        =   14
      Top             =   3555
      Width           =   1395
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   210
      Left            =   1320
      TabIndex        =   13
      Top             =   3075
      Width           =   795
   End
   Begin VB.Label lblCel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      Height          =   210
      Left            =   1635
      TabIndex        =   12
      Top             =   2595
      Width           =   495
   End
   Begin VB.Label lblTel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home Phone:"
      Height          =   210
      Left            =   1110
      TabIndex        =   11
      Top             =   2115
      Width           =   945
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home Address:"
      Height          =   210
      Left            =   930
      TabIndex        =   10
      Top             =   1635
      Width           =   1125
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   210
      Left            =   1695
      TabIndex        =   9
      Top             =   1155
      Width           =   450
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   210
      Left            =   1245
      TabIndex        =   8
      Top             =   675
      Width           =   840
   End
End
Attribute VB_Name = "Create"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

With dbUsers

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

Private Sub lblCancel_Click()

Cancel_Click

End Sub

Private Sub lblSign_Click()

SignIn_Click

End Sub

Private Sub picCancel_Click()

Cancel_Click

End Sub

Private Sub picSign_Click()

SignIn_Click

End Sub

Private Sub txtAddress_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 97 To 122                      'a-z
        KeyAscii = KeyAscii - 32        'Convert to capital letters
    Case vbKeyReturn
        txtTel.SetFocus

End Select

End Sub

Private Sub txtAddress_LostFocus()

If txtAddress.Text = Empty Then txtAddress.Text = "N/A"

End Sub

Private Sub txtCel_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtCel_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        txtPass.SetFocus                'Set focus to Birth Date
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtCel_LostFocus()

If txtCel.Text = Empty Then txtCel.Text = "N/A"

End Sub

Private Sub txtConfirm_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then txtHint.SetFocus

End Sub

Private Sub txtHint_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtHint_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 97 To 122                      'a-z
        KeyAscii = KeyAscii - 32        'Convert to capital letters
    Case 34, 39
        KeyAscii = 0
    Case vbKeyReturn
        SignIn_Click

End Select

End Sub

Private Sub txtName_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 46                             'Period (.)
    Case 65 To 90                       'A-Z
    Case 97 To 122                      'a-z
        KeyAscii = KeyAscii - 32        'Convert to capital letters
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        txtAddress.SetFocus             'Set focus to txtAddress
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtPass_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then txtConfirm.SetFocus

End Sub

Private Sub txtTel_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        txtCel.SetFocus                 'Set focus to txtCel
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtTel_LostFocus()

If txtTel.Text = Empty Then txtTel.Text = "N/A"

End Sub

Private Sub txtUser_Click()

SendKeys "{home}+{end}"

End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                     'minus sign (-)
    Case 95                     'underscore sign (_)
    Case 48 To 57               '0-9
    Case 65 To 90               'A-Z
    Case 97 To 122              'a-z
    Case vbKeyBack              'Backspace key
    Case vbKeySpace             'Spacebar key
    Case vbKeyReturn            'Enter key
        txtName.SetFocus
    Case Else
        KeyAscii = 0
        
End Select

End Sub

Sub Display_Message(ctrl As Control)

msg = "Please fill-up the following:"
msg = msg & vbCrLf
msg = msg & ctrl.Tag

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

MsgBox msg, , "Message"

ctrl.SetFocus

End Sub

Sub Cancel_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Unload Me

Load Login
Login.Show vbModal

End Sub

Sub SignIn_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

If txtUser.Text = Empty Then
    Call Display_Message(txtUser)
    Exit Sub
ElseIf txtName.Text = Empty Then
    Call Display_Message(txtName)
    Exit Sub
ElseIf txtAddress.Text = Empty Then
    Call Display_Message(txtAddress)
    Exit Sub
End If

If txtPass.Text = Empty Then
    Call Display_Message(txtPass)
    Exit Sub
ElseIf txtConfirm.Text = Empty Then
    Call Display_Message(txtConfirm)
    Exit Sub
ElseIf txtPass.Text <> txtConfirm.Text Then
    MsgBox "Password and confirm password are not the same.", , "Message"
    txtPass.Text = ""
    txtConfirm.Text = ""
    txtPass.SetFocus
    Exit Sub
End If

If txtHint.Text = Empty Then
    Call Display_Message(txtHint)
    Exit Sub
End If

If txtTel.Text = Empty Then txtTel.Text = "N/A"
If txtCel.Text = Empty Then txtCel.Text = "N/A"

With dbUsers.Recordset
    
    If .RecordCount >= 1 Then
    
        .MoveFirst
        .FindFirst "User='" & txtUser.Text & "'"
        
        If .NoMatch = False Then
            
            sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
            
            MsgBox "User Name found. Enter a different User Name.", , "Message"
            txtUser.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
            
        End If
    
    End If
    
    If Not txtPass.Text = txtConfirm.Text Then
    
        sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
        
        MsgBox "Password and Confirm Password must be the same.", , "Message"
        txtPass.Text = ""
        txtConfirm.Text = ""
        txtPass.SetFocus
        Exit Sub
    
    End If

    .AddNew
    
    .Fields("User") = txtUser.Text
    .Fields("Name") = txtName.Text
    .Fields("Address") = txtAddress.Text
    .Fields("Home Phone") = txtTel.Text
    .Fields("Mobile") = txtCel.Text
    
    .Fields("Password") = txtPass.Text
    .Fields("Hint") = txtHint.Text
    
    .Update
    
   
    sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC
    MsgBox "User " & txtUser.Text & " has been saved.", , "Message"
    

    Main.picPick.Visible = True
    Main.lblUser = "User - [" & UCase(txtUser) & "]"
    Main.Tag = txtUser.Text
    
    Unload Me
    
    RemindChecker.Show
    RemindChecker.Hide

End With

End Sub

