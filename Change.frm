VERSION 5.00
Begin VB.Form ChangePass 
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Change.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Change.frx":030A
   ScaleHeight     =   3495
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2340
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
      Left            =   2318
      MouseIcon       =   "Change.frx":3415C
      MousePointer    =   99  'Custom
      Picture         =   "Change.frx":342AE
      ScaleHeight     =   435
      ScaleWidth      =   1725
      TabIndex        =   9
      Top             =   2640
      Width           =   1725
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cancel"
         Height          =   210
         Left            =   585
         MouseIcon       =   "Change.frx":36A5C
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.PictureBox picSave 
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
      Left            =   518
      MouseIcon       =   "Change.frx":36BAE
      MousePointer    =   99  'Custom
      Picture         =   "Change.frx":36D00
      ScaleHeight     =   435
      ScaleWidth      =   1725
      TabIndex        =   7
      Top             =   2640
      Width           =   1725
      Begin VB.Label lblSave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Save Password"
         Height          =   210
         Left            =   255
         MouseIcon       =   "Change.frx":394AE
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.TextBox txtConfirm 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2228
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Tag             =   "Confirm Password"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtNew 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2228
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Tag             =   "New Password"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtOld 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2228
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "Old Password"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      Height          =   210
      Left            =   525
      TabIndex        =   6
      Top             =   2115
      Width           =   1395
   End
   Begin VB.Label lblConfirm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      Height          =   210
      Left            =   735
      TabIndex        =   5
      Top             =   1635
      Width           =   1185
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      Height          =   210
      Left            =   840
      TabIndex        =   4
      Top             =   1155
      Width           =   1080
   End
   Begin VB.Label lblProfile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "ChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub lblSave_Click()

End Sub

Private Sub picSave_Click()

End Sub
