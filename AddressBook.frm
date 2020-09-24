VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form AddressBook 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5145
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
   MouseIcon       =   "AddressBook.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "AddressBook.frx":030A
   ScaleHeight     =   5145
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
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Address"
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
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   120
      Picture         =   "AddressBook.frx":6BAE8
      ScaleHeight     =   2085
      ScaleWidth      =   1860
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1860
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&Close"
         Height          =   210
         Index           =   4
         Left            =   120
         MouseIcon       =   "AddressBook.frx":78526
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&View Person"
         Height          =   210
         Index           =   3
         Left            =   120
         MouseIcon       =   "AddressBook.frx":78678
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&Delete Person"
         Height          =   210
         Index           =   2
         Left            =   120
         MouseIcon       =   "AddressBook.frx":787CA
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&Edit Person"
         Height          =   210
         Index           =   1
         Left            =   120
         MouseIcon       =   "AddressBook.frx":7891C
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&Add Person"
         Height          =   210
         Index           =   0
         Left            =   120
         MouseIcon       =   "AddressBook.frx":78A6E
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   120
         Width           =   1605
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "AddressBook.frx":78BC0
      Height          =   3135
      Left            =   240
      OleObjectBlob   =   "AddressBook.frx":78BD4
      TabIndex        =   3
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Type name or select from list:"
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   1275
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address Book"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      Height          =   210
      Left            =   195
      MouseIcon       =   "AddressBook.frx":79597
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "AddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBGrid1_Click()

picMenu.Visible = False

End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Reset_Menu

End Sub

Private Sub Form_Click()

picMenu.Visible = False

End Sub

Private Sub Form_Load()

With Data1

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT Address.Name, Address.[Home Address], " & _
        "Address.[Home Phone], Address.Mobile FROM Address " & _
        "WHERE Address.UserName='" & Main.Tag & "' " & _
        "ORDER BY Address.Name"
    .Refresh

    With .Recordset
    
        If .RecordCount = 0 And .EOF = True Then
            
            lblMenu(1).Enabled = False
            lblMenu(2).Enabled = False
            lblMenu(3).Enabled = False
        
        End If
    
    End With

End With

With DBGrid1

    .Columns(0).Width = 2000
    .Columns(1).Width = 4000
    .Columns(2).Width = 1500
    .Columns(3).Width = 1500

End With

With Data2

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM Address WHERE Address.UserName='" & _
        Main.Tag & "' "
    .Refresh

End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If (X >= 0 And Y >= 0) And (X <= 4200 And Y <= 460) Then
    If Button = vbLeftButton Then Call DragIt(Me.hwnd)
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Reset_Menu

End Sub

Private Sub lblFile_Click()

picMenu.Visible = Not picMenu.Visible

End Sub

Private Sub lblMenu_Click(Index As Integer)

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

picMenu.Visible = False

Select Case Index

    Case 0
        Add_Person
    Case 1
        Edit_Person
    Case 2
        Delete_Person
    Case 3
        View_Person
    Case 4
        Unload Me
        Main.SetFocus
End Select

End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If lblMenu(Index).FontUnderline = True Then Exit Sub
Reset_Menu
lblMenu(Index).FontUnderline = True

End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Reset_Menu

End Sub

Private Sub txtSearch_Change()

On Error Resume Next

With Data1.Recordset

    .MoveFirst
    .FindFirst "Name like '" & txtSearch.Text & "*'"

End With

End Sub

Private Sub txtSearch_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

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

Private Sub txtSearch_Click()

picMenu.Visible = False

End Sub

Private Sub txtSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Reset_Menu

End Sub

Sub Reset_Menu()

For i = 0 To 4
    lblMenu(i).FontUnderline = False
Next

End Sub

Sub Add_Person()

Me.Tag = "Add"

Load ABForm
ABForm.Show , Me

End Sub

Sub Delete_Person()

With Data1.Recordset

    If .RecordCount = 0 Then
        MsgBox "Address Book is empty.", vbExclamation, "Message"
        lblMenu(1).Enabled = False
        lblMenu(2).Enabled = False
        lblMenu(3).Enabled = False
        Exit Sub
    End If
    
    person = .Fields("Name")
    ans = MsgBox("Are you sure?", vbQuestion + vbYesNo, _
            "Delete " & person & " ?")
    
    If ans = vbNo Then Exit Sub
    
    .Delete
    MsgBox person & " has been deleted.", , "Message"
    
    If .RecordCount <> 0 Then .MoveFirst

End With

Refresh_Records

End Sub

Sub Edit_Person()

Me.Tag = "Edit"

Data2.Refresh
With Data2.Recordset

    If .RecordCount = 0 And .EOF = True Then
        MsgBox "Address Book is empty.", vbExclamation, "Message"
        Exit Sub
    End If

End With

ABForm.Show , Me

End Sub

Sub View_Person()

Me.Tag = "View"

With Data2.Recordset

    If .RecordCount = 0 Then
        MsgBox "No person to be view.", vbExclamation, "Message"
        Exit Sub
    End If

    .MoveFirst
    .FindFirst "Name='" & Data1.Recordset.Fields("Name") & "'"
    
    ABForm.txtName.Text = .Fields("Name")
    ABForm.txtAddress.Text = .Fields("Home Address")
    ABForm.txtPhone.Text = .Fields("Home Phone")
    ABForm.txtPager.Text = .Fields("Pager")
    ABForm.txtMobile.Text = .Fields("Mobile")
    ABForm.txtBPhone.Text = .Fields("Business Phone")
    ABForm.txtFax.Text = .Fields("Business Fax")
    ABForm.txtEmail.Text = .Fields("Email Address")

End With

ABForm.Show , Me

End Sub

Sub Refresh_Records()

With Data1
    
    .Refresh

    If .Recordset.RecordCount = 0 Then
        lblMenu(1).Enabled = False
        lblMenu(2).Enabled = False
        lblMenu(3).Enabled = False
        Exit Sub
    End If

End With

With DBGrid1

    .Refresh
    .Columns(0).Width = 2000
    .Columns(1).Width = 4000
    .Columns(2).Width = 1500
    .Columns(3).Width = 1500

End With

End Sub
