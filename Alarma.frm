VERSION 5.00
Begin VB.Form Alarma 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   165
   ClientWidth     =   4530
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
   Icon            =   "Alarma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Alarma.frx":000C
   MousePointer    =   99  'Custom
   Picture         =   "Alarma.frx":0316
   ScaleHeight     =   2520
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   480
      Top             =   1320
   End
   Begin VB.Label lblClick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to close Alarm"
      Height          =   210
      Left            =   2040
      MouseIcon       =   "Alarma.frx":26488
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1440
      Width           =   1785
   End
End
Attribute VB_Name = "Alarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

sndPlaySound App.Path & "\sounds\bell.wav", SND_SYNC

End Sub

Private Sub lblClick_Click()

sndPlaySound App.Path & "\sounds\longhorn_question.wav", SND_ASYNC

Unload Me
Set Alarm = Nothing

End Sub

Private Sub Timer1_Timer()

sndPlaySound App.Path & "\sounds\bell.wav", SND_ASYNC

End Sub
