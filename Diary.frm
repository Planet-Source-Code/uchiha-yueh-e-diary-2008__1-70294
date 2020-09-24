VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Diary 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Diary.frx":0000
   ScaleHeight     =   5670
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEdit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2520
      MouseIcon       =   "Diary.frx":7575E
      MousePointer    =   99  'Custom
      Picture         =   "Diary.frx":758B0
      ScaleHeight     =   435
      ScaleWidth      =   1725
      TabIndex        =   8
      Top             =   4920
      Width           =   1725
      Begin VB.Label lblEdit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Edit Diary"
         Height          =   195
         Left            =   495
         MouseIcon       =   "Diary.frx":7805E
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.Data Data2 
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
      Top             =   3120
      Visible         =   0   'False
      Width           =   2100
   End
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
      Top             =   2760
      Visible         =   0   'False
      Width           =   2100
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
      MouseIcon       =   "Diary.frx":781B0
      MousePointer    =   99  'Custom
      Picture         =   "Diary.frx":78302
      ScaleHeight     =   435
      ScaleWidth      =   1725
      TabIndex        =   6
      Top             =   4920
      Width           =   1725
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Close Diary"
         Height          =   195
         Left            =   435
         MouseIcon       =   "Diary.frx":7AAB0
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.PictureBox picWrite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   720
      MouseIcon       =   "Diary.frx":7AC02
      MousePointer    =   99  'Custom
      Picture         =   "Diary.frx":7AD54
      ScaleHeight     =   435
      ScaleWidth      =   1725
      TabIndex        =   4
      Top             =   4920
      Width           =   1725
      Begin VB.Label lblWrite 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Write to Diary"
         Height          =   195
         Left            =   360
         MouseIcon       =   "Diary.frx":7D502
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      ExtentX         =   10398
      ExtentY         =   6588
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   990
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   62586881
      CurrentDate     =   37977
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diary"
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   675
      Width           =   330
   End
End
Attribute VB_Name = "Diary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim neo As New FileSystemObject
Dim morpheus As TextStream

Private Sub DTPicker1_Change()

Data1.Refresh

With Data1.Recordset
    
    If .RecordCount = 0 Or .EOF = True Then
        WebBrowser1.Navigate App.Path & "\blank.htm"
        
        lblWrite.Enabled = True
        picWrite.Enabled = True
        
        lblEdit.Enabled = False
        picEdit.Enabled = False
        
        Exit Sub
    End If
    
    .MoveFirst
    .FindFirst "Date='" & DTPicker1.Value & "'"
    
    If .NoMatch = True Then
        WebBrowser1.Navigate App.Path & "\blank.htm"
        lblWrite.Enabled = True
        picWrite.Enabled = True
        
        lblEdit.Enabled = False
        picEdit.Enabled = False
        
        Exit Sub
    End If
        
    WriteHTML .Fields("HTML")
    WebBrowser1.Navigate App.Path & "\sample.htm"
    
    lblWrite.Enabled = False
    picWrite.Enabled = False
        
    lblEdit.Enabled = True
    picEdit.Enabled = True
    
End With

End Sub

Private Sub Form_Activate()

DTPicker1_Change

End Sub

Private Sub Form_Load()

With Data2
    
    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = "Script"
    .Refresh
    
    With .Recordset
    
        JScript = .Fields("Script")
    
    End With
    
End With

With Data1
    
    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM Diary WHERE Username ='" & Main.Tag & _
        "' ORDER BY Date"
    .Refresh

    With .Recordset
    
        DTPicker1.Value = Date
        Call DTPicker1_Change
    
        If .RecordCount = 0 And .EOF = True Then
            WebBrowser1.Navigate App.Path & "\blank.htm"
            lblWrite.Enabled = True
            picWrite.Enabled = True
            Exit Sub
        End If
        
    End With

End With

End Sub

Sub WriteHTML(html As String)

On Error Resume Next

Set morpheus = neo.OpenTextFile(App.Path & "\noname.htm", ForWriting)

With morpheus

    .WriteLine "<html>"
    
    .WriteLine "<head>"
    .WriteLine "</head>"
    
    .WriteLine JScript
    
    .WriteLine "<body>"
    .WriteLine "<font face=""Arial"">"
        
    .WriteLine html
    
    .WriteLine "</font>"
    .WriteLine "</body>"
    .WriteLine "</html>"

    .Close

End With

FileCopy App.Path & "\noname.htm", App.Path & "\sample.htm"

End Sub

Private Sub lblClose_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Unload Me
Main.Enabled = True
Main.SetFocus

End Sub

Private Sub lblEdit_Click()

Edit_Click

End Sub

Private Sub lblWrite_Click()

Write_Click

End Sub

Sub Write_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Load WriteDiary
WriteDiary.Show , Me

WriteDiary.Tag = DTPicker1.Value

Me.Enabled = False
Main.Enabled = False

End Sub

Private Sub picClose_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Unload Me
Main.Enabled = True
Main.SetFocus

End Sub

Private Sub picEdit_Click()

Edit_Click

End Sub

Private Sub picWrite_Click()

Write_Click

End Sub

Sub Edit_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Load WriteDiary
WriteDiary.Show , Me

WriteDiary.Tag = DTPicker1.Value
WriteDiary.txtDiary.Text = Data1.Recordset.Fields("diary")

Me.Enabled = False
Main.Enabled = False

End Sub
