VERSION 5.00
Begin VB.Form Profile 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5175
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
   MouseIcon       =   "Profile.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Profile.frx":030A
   ScaleHeight     =   5175
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picChange 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   885
      Picture         =   "Profile.frx":6BAE8
      ScaleHeight     =   3495
      ScaleWidth      =   4560
      TabIndex        =   25
      Top             =   825
      Visible         =   0   'False
      Width           =   4590
      Begin VB.TextBox txtConfirm 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   7
         Tag             =   "Confirm Password"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtNew 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         Tag             =   "New Password"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtOld 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   5
         Tag             =   "Old Password"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.PictureBox picCancel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Left            =   2400
         MouseIcon       =   "Profile.frx":9F93A
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":9FA8C
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   28
         Top             =   2760
         Width           =   1725
         Begin VB.Label lblCancel 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel"
            Height          =   210
            Left            =   585
            MouseIcon       =   "Profile.frx":A223A
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   120
            Width           =   525
         End
      End
      Begin VB.PictureBox picSavePass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Left            =   600
         MouseIcon       =   "Profile.frx":A238C
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":A24DE
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   26
         Top             =   2760
         Width           =   1725
         Begin VB.Label lblSavePass 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Save Password"
            Height          =   210
            Left            =   255
            MouseIcon       =   "Profile.frx":A4C8C
            MousePointer    =   99  'Custom
            TabIndex        =   27
            Top             =   120
            Width           =   1185
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         Height          =   210
         Left            =   480
         TabIndex        =   33
         Top             =   120
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         Height          =   210
         Left            =   480
         TabIndex        =   32
         Top             =   2235
         Width           =   1395
      End
      Begin VB.Label lblConfirm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
         Height          =   210
         Left            =   780
         TabIndex        =   31
         Top             =   1755
         Width           =   1185
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
         Height          =   210
         Left            =   870
         TabIndex        =   30
         Top             =   1275
         Width           =   1080
      End
   End
   Begin VB.PictureBox picProfile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   240
      Picture         =   "Profile.frx":A4DDE
      ScaleHeight     =   3855
      ScaleWidth      =   5895
      TabIndex        =   9
      Top             =   1080
      Width           =   5895
      Begin VB.PictureBox picSaveProfile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Left            =   1200
         MouseIcon       =   "Profile.frx":FB9AC
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":FBAFE
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   23
         Top             =   3120
         Visible         =   0   'False
         Width           =   1725
         Begin VB.Label lblSaveProfile 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Save Profile"
            Height          =   210
            Left            =   405
            MouseIcon       =   "Profile.frx":FE2AC
            MousePointer    =   99  'Custom
            TabIndex        =   24
            Top             =   120
            Width           =   885
         End
      End
      Begin VB.PictureBox picCancelChange 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Left            =   3000
         MouseIcon       =   "Profile.frx":FE3FE
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":FE550
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   1725
         Begin VB.Label lblCancelChange 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel  Changes"
            Height          =   210
            Left            =   225
            MouseIcon       =   "Profile.frx":100CFE
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   120
            Width           =   1245
         End
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
         Left            =   1380
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1080
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.PictureBox picClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Left            =   3840
         MouseIcon       =   "Profile.frx":100E50
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":100FA2
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   19
         Top             =   3120
         Width           =   1725
         Begin VB.Label lblClose 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Close"
            Height          =   210
            Left            =   630
            MouseIcon       =   "Profile.frx":103750
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   120
            Width           =   435
         End
      End
      Begin VB.PictureBox picEditProfile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Left            =   240
         MouseIcon       =   "Profile.frx":1038A2
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":1039F4
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   17
         Top             =   3120
         Width           =   1725
         Begin VB.Label lblEditProfile 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Edit Profile"
            Height          =   210
            Left            =   465
            MouseIcon       =   "Profile.frx":1061A2
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.PictureBox picChangePass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Left            =   2040
         MouseIcon       =   "Profile.frx":1062F4
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":106446
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   15
         Top             =   3120
         Width           =   1725
         Begin VB.Label lblChangePass 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change &Password"
            Height          =   210
            Left            =   165
            MouseIcon       =   "Profile.frx":108BF4
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   120
            Width           =   1365
         End
      End
      Begin VB.TextBox txtMobile 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   20
         MousePointer    =   3  'I-Beam
         TabIndex        =   4
         Tag             =   "Mobile"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1935
         Locked          =   -1  'True
         MaxLength       =   20
         MousePointer    =   3  'I-Beam
         TabIndex        =   3
         Tag             =   "Home Phone"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   2
         Tag             =   "Home Address"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   3  'I-Beam
         TabIndex        =   1
         Tag             =   "Name"
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   20
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         Tag             =   "User Name"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblMobile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
         Height          =   210
         Left            =   1230
         TabIndex        =   14
         Top             =   2235
         Width           =   495
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Phone:"
         Height          =   210
         Left            =   780
         TabIndex        =   13
         Top             =   1755
         Width           =   945
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address:"
         Height          =   210
         Left            =   600
         TabIndex        =   12
         Top             =   1275
         Width           =   1125
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   210
         Left            =   1275
         TabIndex        =   11
         Top             =   795
         Width           =   450
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   210
         Left            =   885
         TabIndex        =   10
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.Label lblProfile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profile"
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
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "Profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

With dbUsers

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM Users WHERE User='" & Main.Tag & "'"
    .Refresh
    
    Call Display_Record

End With

End Sub

Private Sub lblCancel_Click()

picProfile.Enabled = True
picChange.Visible = False
txtUser.SetFocus

End Sub

Private Sub lblCancelChange_Click()

Form_Controls
txtUser.SetFocus

End Sub

Private Sub lblChangePass_Click()

picProfile.Enabled = False
picChange.Visible = True

txtOld.Text = ""
txtNew.Text = ""
txtConfirm.Text = ""

txtOld.SetFocus

End Sub

Private Sub lblClose_Click()

Unload Me

End Sub

Private Sub lblEditProfile_Click()

Form_Controls
txtAddress.SetFocus

End Sub

Private Sub lblSavePass_Click()

Call Save_Pass

End Sub

Private Sub lblSaveProfile_Click()

Call SavePro

End Sub

Private Sub picCancel_Click()

picProfile.Enabled = True
picChange.Visible = False
txtUser.SetFocus

End Sub

Private Sub picCancelChange_Click()

Form_Controls
txtUser.SetFocus

End Sub

Private Sub picChangePass_Click()

picProfile.Enabled = False
picChange.Visible = True

txtOld.Text = ""
txtNew.Text = ""
txtConfirm.Text = ""

txtOld.SetFocus

End Sub

Private Sub picClose_Click()

Unload Me

End Sub

Private Sub picEditProfile_Click()

Form_Controls
txtAddress.SetFocus

End Sub

Private Sub picSavePass_Click()

Call SavePro

End Sub

Private Sub picSaveProfile_Click()

Call SavePro
picCancel_Click

End Sub

Private Sub txtAddress_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 97 To 122                      'a-z
        KeyAscii = KeyAscii - 32        'Convert to capital letters
    Case vbKeyReturn
        txtPhone.SetFocus

End Select

End Sub

Private Sub txtAddress_LostFocus()

If txtAddress.Text = Empty Then
    txtAddress.Text = "N/A"
End If

End Sub

Private Sub txtConfirm_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then Call Save_Pass

End Sub

Private Sub txtMobile_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        Call SavePro
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtMobile_LostFocus()

If txtMobile.Text = Empty Then
    txtMobile.Text = "N/A"
End If

End Sub

Private Sub txtName_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtNew_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    txtConfirm.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    txtNew.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub txtPhone_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        txtMobile.SetFocus             'Set focus to txtPager
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtPhone_LostFocus()

If txtPhone.Text = Empty Then
    txtPhone.Text = "N/A"
End If

End Sub

Sub Form_Controls()

txtAddress.Locked = Not txtAddress.Locked
txtPhone.Locked = Not txtPhone.Locked
txtMobile.Locked = Not txtMobile.Locked

picEditProfile.Visible = Not picEditProfile.Visible
picChangePass.Visible = Not picChangePass.Visible
picClose.Visible = Not picClose.Visible

picSaveProfile.Visible = Not picSaveProfile.Visible
picCancelChange.Visible = Not picCancelChange.Visible

End Sub

Sub SavePro()

If txtAddress.Text = Empty Then txtAddress.Text = "N/A"
If txtPhone.Text = Empty Then txtPhone.Text = "N/A"
If txtMobile.Text = Empty Then txtMobile.Text = "N/A"

With dbUsers.Recordset

    .Edit

    .Fields("Address") = txtAddress.Text
    .Fields("Home Phone") = txtPhone.Text
    .Fields("Mobile") = txtMobile.Text
    
    .Update
    
    MsgBox "Profile has been changed.", vbInformation, "Message"

End With

Call Form_Controls
txtUser.SetFocus

End Sub

Sub Display_Record()

With dbUsers.Recordset

    txtUser.Text = .Fields("User")
    txtName.Text = .Fields("Name")
    txtAddress.Text = .Fields("Address")
    txtPhone.Text = .Fields("Home Phone")
    txtMobile.Text = .Fields("Mobile")

    picChange.Tag = .Fields("Password")

End With

End Sub

Sub Save_Pass()

If txtOld.Text = Empty Then
    MsgBox "Please enter your old password.", vbExclamation, "Message"
    txtOld.SetFocus
    Exit Sub
ElseIf txtOld.Text <> picChange.Tag Then
    MsgBox "Old password is incorrect.", vbExclamation, "Message"
    txtOld.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
ElseIf txtNew.Text = Empty Then
    MsgBox "Please enter your new password.", vbExclamation, "Message"
    txtOld.SetFocus
    Exit Sub
ElseIf txtConfirm.Text = Empty Then
    MsgBox "Please enter your confirm password.", vbExclamation, "Message"
    txtConfirm.SetFocus
    Exit Sub
ElseIf txtNew.Text <> txtConfirm Then
    MsgBox "New and confirm password must be the same.", _
        vbExclamation, "Message"
    txtNew.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

With dbUsers.Recordset

    .Edit
    
    .Fields("Password") = txtNew.Text
    .Update
    
    MsgBox "Password has been changed."

End With

picCancel_Click

End Sub

Private Sub txtUser_GotFocus()

SendKeys "{home}+{end}"

End Sub
