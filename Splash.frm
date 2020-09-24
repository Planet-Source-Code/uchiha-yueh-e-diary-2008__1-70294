VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   3720
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2760
      Top             =   600
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

sndPlaySound App.Path & "\sounds\longhorn_ringout.wav", SND_ASYNC

End Sub

Private Sub Timer1_Timer()

Load Main
Main.Show

Unload Me

Load Login
Login.Show vbModal

End Sub
