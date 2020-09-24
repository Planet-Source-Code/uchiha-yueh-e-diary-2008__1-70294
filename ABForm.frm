VERSION 5.00
Begin VB.Form ABForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
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
   MouseIcon       =   "ABForm.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "ABForm.frx":030A
   ScaleHeight     =   5655
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3390
      MouseIcon       =   "ABForm.frx":75A68
      MousePointer    =   99  'Custom
      Picture         =   "ABForm.frx":75BBA
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   17
      Top             =   4920
      Width           =   1515
      Begin VB.Label lblCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cancel"
         Height          =   210
         Left            =   480
         MouseIcon       =   "ABForm.frx":77E6C
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   80
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
      Left            =   1470
      MouseIcon       =   "ABForm.frx":77FBE
      MousePointer    =   99  'Custom
      Picture         =   "ABForm.frx":78110
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   16
      Top             =   4920
      Width           =   1515
      Begin VB.Label lblSave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Save"
         Height          =   210
         Left            =   600
         MouseIcon       =   "ABForm.frx":7A3C2
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   75
         Width           =   375
      End
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   15
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Top             =   2970
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   50
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   2280
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Top             =   2025
      Width           =   2895
   End
   Begin VB.TextBox txtPager 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   15
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Top             =   2505
      Width           =   2895
   End
   Begin VB.TextBox txtBPhone 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Top             =   3450
      Width           =   2895
   End
   Begin VB.TextBox txtFax 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      MousePointer    =   3  'I-Beam
      TabIndex        =   13
      Top             =   3915
      Width           =   2895
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   100
      MousePointer    =   3  'I-Beam
      TabIndex        =   15
      Top             =   4395
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   210
      Left            =   1590
      TabIndex        =   0
      Top             =   675
      Width           =   405
   End
   Begin VB.Label lblHAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home Address"
      Height          =   210
      Left            =   915
      TabIndex        =   2
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label lblHPhone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home Phone"
      Height          =   210
      Left            =   1095
      TabIndex        =   4
      Top             =   2070
      Width           =   900
   End
   Begin VB.Label lblPager 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pager"
      Height          =   210
      Left            =   1575
      TabIndex        =   6
      Top             =   2580
      Width           =   420
   End
   Begin VB.Label lblMobile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      Height          =   210
      Left            =   1545
      TabIndex        =   8
      Top             =   3045
      Width           =   450
   End
   Begin VB.Label lblBPhone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Business Phone"
      Height          =   210
      Left            =   825
      TabIndex        =   10
      Top             =   3525
      Width           =   1170
   End
   Begin VB.Label lblBFax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Business Fax"
      Height          =   210
      Left            =   1005
      TabIndex        =   12
      Top             =   3990
      Width           =   990
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      Height          =   210
      Left            =   960
      TabIndex        =   14
      Top             =   4470
      Width           =   1035
   End
End
Attribute VB_Name = "ABForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

On Error Resume Next

With Data2

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM Address WHERE Address.UserName='" & _
        Main.Tag & "' "
    .Refresh

End With

If AddressBook.Tag = "Add" Then
    
    lblTitle.Caption = "Add Person"

ElseIf AddressBook.Tag = "Edit" Then
    
    lblTitle.Caption = "Edit Person"
    
    With Data2.Recordset
        
        .MoveFirst
        .FindFirst "Name='" & AddressBook.Data1.Recordset.Fields("name") & "'"

        txtName.Text = .Fields("Name")
        txtAddress.Text = .Fields("Home Address")
        txtPhone.Text = .Fields("Home Phone")
        txtPager.Text = .Fields("Pager")
        txtMobile.Text = .Fields("Mobile")
        txtBPhone.Text = .Fields("Business Phone")
        txtFax.Text = .Fields("Business Fax")
        txtEmail.Text = .Fields("Email Address")
    
    End With
    
ElseIf AddressBook.Tag = "View" Then
    
    lblTitle.Caption = "View Person"
    Locked_Textbox

End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If (X >= 0 And Y >= 0) And (X <= 4200 And Y <= 460) Then
    If Button = vbLeftButton Then Call DragIt(Me.hwnd)
End If

End Sub

Private Sub lblCancel_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Cancel_Click

End Sub

Private Sub lblSave_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Save_Click

End Sub

Private Sub picCancel_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Cancel_Click

End Sub

Private Sub picSave_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Save_Click

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

If txtAddress.Text = Empty Or txtAddress.Text = vbCrLf Then
    txtAddress.Text = "N/A"
End If

End Sub

Private Sub txtBPhone_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtBPhone_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        txtFax.SetFocus               'Set focus to txtEmail
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtBPhone_LostFocus()

If txtBPhone.Text = Empty Then
    txtBPhone.Text = "N/A"
End If

End Sub

Private Sub txtEmail_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)

If AddressBook.Tag = "View" Then Exit Sub

Select Case KeyAscii
    
    Case 97 To 122                      'a-z
        KeyAscii = KeyAscii - 32        'Convert to capital letters
    Case vbKeyReturn
        Save_Click

End Select

End Sub

Private Sub txtEmail_LostFocus()

If txtEmail.Text = Empty Then
    txtEmail.Text = "N/A"
End If

End Sub

Private Sub txtFax_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        txtEmail.SetFocus               'Set focus to txtEmail
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtFax_LostFocus()

If txtFax.Text = Empty Then
    txtFax.Text = "N/A"
End If

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
        txtBPhone.SetFocus              'Set focus to txtBPhone
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

Private Sub txtName_LostFocus()

txtName.Text = UCase(txtName.Text)

End Sub

Private Sub txtPager_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtPager_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        txtMobile.SetFocus             'Set focus to txtMobile
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtPager_LostFocus()

If txtPager.Text = Empty Then
    txtPager.Text = "N/A"
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
        txtPager.SetFocus             'Set focus to txtPager
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtPhone_LostFocus()

If txtPhone.Text = Empty Then
    txtPhone.Text = "N/A"
End If

End Sub

Sub Save_Click()

On Error Resume Next

If txtName.Text = Empty Then
    Call Display_Error(txtName)
    Exit Sub
End If

If txtAddress.Text = Empty Then txtAddress.Text = "N/A"
If txtPhone.Text = Empty Then txtPhone.Text = "N/A"
If txtPager.Text = Empty Then txtPager.Text = "N/A"
If txtMobile.Text = Empty Then txtMobile.Text = "N/A"
If txtBPhone.Text = Empty Then txtBPhone.Text = "N/A"
If txtFax.Text = Empty Then txtFax.Text = "N/A"
If txtEmail.Text = Empty Then txtEmail.Text = "N/A"

With Data2.Recordset

    If AddressBook.Tag = "Add" Then
        .AddNew
        .Fields("UserName") = Main.Tag
    ElseIf AddressBook.Tag = "Edit" Then
        .Edit
    End If
    
    .Fields("Name") = txtName.Text
    .Fields("Home Address") = txtAddress.Text
    .Fields("Home Phone") = txtPhone.Text
    .Fields("Pager") = txtPager.Text
    .Fields("Mobile") = txtMobile.Text
    .Fields("Business Phone") = txtBPhone.Text
    .Fields("Business Fax") = txtFax.Text
    .Fields("Email Address") = txtEmail.Text
    
    .Update

End With
    
With AddressBook

    .Data1.Refresh
    
    AddressBook.lblMenu(1).Enabled = True
    AddressBook.lblMenu(2).Enabled = True
    AddressBook.lblMenu(3).Enabled = True
        

    With .DBGrid1

        .Refresh
        .Columns(0).Width = 2000
        .Columns(1).Width = 4000
        .Columns(2).Width = 1500
        .Columns(3).Width = 1500

    End With

End With
    
MsgBox "Person has been saved.", , "Message"
Cancel_Click

End Sub

Sub Display_Error(ctrl As Control)

msg = "Please fill-up the following:"
msg = msg & vbCrLf
msg = msg & ctrl.Tag

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

MsgBox msg, , "Message"

ctrl.SetFocus

End Sub

Sub Cancel_Click()

If AddressBook.Tag = "View" Then Locked_Textbox
Unload Me

End Sub

Sub Locked_Textbox()

On Error Resume Next

With ABForm
    .txtName.Locked = Not .txtName.Locked
    .txtAddress.Locked = Not .txtAddress.Locked
    .txtPhone.Locked = Not .txtPhone.Locked
    .txtPager.Locked = Not .txtPager.Locked
    .txtMobile.Locked = Not .txtMobile.Locked
    .txtBPhone.Locked = Not .txtBPhone.Locked
    .txtFax.Locked = Not .txtFax.Locked
    .txtEmail.Locked = Not .txtEmail.Locked
End With

End Sub

