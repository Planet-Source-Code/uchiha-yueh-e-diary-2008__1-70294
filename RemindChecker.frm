VERSION 5.00
Begin VB.Form RemindChecker 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RemindChecker.frx":0000
   ScaleHeight     =   2535
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   3480
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Width           =   1140
   End
   Begin VB.TextBox txtRemind 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblClick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to close Reminder"
      Height          =   210
      Left            =   1800
      MouseIcon       =   "RemindChecker.frx":26172
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
      Width           =   2040
   End
End
Attribute VB_Name = "RemindChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim num As String

With Data1

    num = Time

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM reminder WHERE Username='" & Main.Tag & _
        "' AND date=#" & Date & _
        "# AND time>=#" & CStr(Hour(Time)) & ":00:00 " & _
        Right(num, 2) & "# ORDER BY TIME"
    .Refresh

    With .Recordset
    
        If .RecordCount = 0 And .EOF = True Then Exit Sub
        List1.Clear
        
        .MoveFirst
        
        Do While .EOF = False
        
            List1.AddItem .Fields("time")
            .MoveNext
        
        Loop
    
    End With

End With

End Sub

Private Sub lblClick_Click()

Timer1.Enabled = True
Timer2.Enabled = False

Me.Hide

End Sub

Private Sub Timer1_Timer()

For i = 0 To List1.ListCount - 1

    If List1.List(i) = Time Then
            
        With Data1.Recordset
            .MoveFirst
            .FindFirst "Time=#" & List1.List(i) & "#"
            txtRemind.Text = .Fields("Remind About")
        End With
        
        MakeTopMost Me.hwnd
        Me.Show
        sndPlaySound App.Path & "\sounds\remind.wav", SND_ASYNC
        Timer1.Enabled = False
        Timer2.Enabled = True
        Exit For
        
    End If

Next

Data1.Refresh
With Data1.Recordset
    
    If .RecordCount = 0 And .EOF = True Then Exit Sub
    
    List1.Clear
    .MoveFirst
        
    Do While .EOF = False
        
        List1.AddItem .Fields("time")
        .MoveNext
        
    Loop
    
End With


End Sub

Private Sub Timer2_Timer()

sndPlaySound App.Path & "\sounds\bell.wav", SND_ASYNC

End Sub
