VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8565
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
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Main.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "Main.frx":0BD4
   ScaleHeight     =   7065
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Glenn"
   Begin VB.PictureBox picPick 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   240
      Picture         =   "Main.frx":D652
      ScaleHeight     =   5025
      ScaleWidth      =   8265
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   8265
      Begin VB.PictureBox picChoices 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   6
         Left            =   4335
         MouseIcon       =   "Main.frx":94D9C
         MousePointer    =   99  'Custom
         Picture         =   "Main.frx":94EEE
         ScaleHeight     =   855
         ScaleWidth      =   705
         TabIndex        =   8
         Top             =   2760
         Width           =   705
      End
      Begin VB.PictureBox picChoices 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   5
         Left            =   4335
         MouseIcon       =   "Main.frx":95959
         MousePointer    =   99  'Custom
         Picture         =   "Main.frx":95AAB
         ScaleHeight     =   855
         ScaleWidth      =   705
         TabIndex        =   7
         Top             =   1800
         Width           =   705
      End
      Begin VB.PictureBox picChoices 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   4
         Left            =   4335
         MouseIcon       =   "Main.frx":9664D
         MousePointer    =   99  'Custom
         Picture         =   "Main.frx":9679F
         ScaleHeight     =   855
         ScaleWidth      =   705
         TabIndex        =   6
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox picChoices 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   3
         Left            =   975
         MouseIcon       =   "Main.frx":9764E
         MousePointer    =   99  'Custom
         Picture         =   "Main.frx":977A0
         ScaleHeight     =   855
         ScaleWidth      =   705
         TabIndex        =   5
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox picChoices 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   2
         Left            =   975
         MouseIcon       =   "Main.frx":9855C
         MousePointer    =   99  'Custom
         Picture         =   "Main.frx":986AE
         ScaleHeight     =   855
         ScaleWidth      =   705
         TabIndex        =   4
         Top             =   2760
         Width           =   705
      End
      Begin VB.PictureBox picChoices 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   975
         MouseIcon       =   "Main.frx":993CD
         MousePointer    =   99  'Custom
         Picture         =   "Main.frx":9951F
         ScaleHeight     =   855
         ScaleWidth      =   705
         TabIndex        =   3
         Top             =   1800
         Width           =   705
      End
      Begin VB.PictureBox picChoices 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   0
         Left            =   975
         MouseIcon       =   "Main.frx":9A3DE
         MousePointer    =   99  'Custom
         Picture         =   "Main.frx":9A530
         ScaleHeight     =   855
         ScaleWidth      =   705
         TabIndex        =   2
         Top             =   840
         Width           =   705
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profile"
         Height          =   210
         Index           =   6
         Left            =   5295
         MouseIcon       =   "Main.frx":9B32F
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   3082
         Width           =   450
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Planner"
         Height          =   210
         Index           =   5
         Left            =   5295
         MouseIcon       =   "Main.frx":9B481
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   2122
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alarm"
         Height          =   210
         Index           =   4
         Left            =   5295
         MouseIcon       =   "Main.frx":9B5D3
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1162
         Width           =   420
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reminders"
         Height          =   210
         Index           =   3
         Left            =   1935
         MouseIcon       =   "Main.frx":9B725
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   4042
         Width           =   765
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   210
         Index           =   2
         Left            =   1935
         MouseIcon       =   "Main.frx":9B877
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   3082
         Width           =   420
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address Book"
         Height          =   210
         Index           =   1
         Left            =   1935
         MouseIcon       =   "Main.frx":9B9C9
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2122
         Width           =   1035
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diary"
         Height          =   210
         Index           =   0
         Left            =   1935
         MouseIcon       =   "Main.frx":9BB1B
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1162
         Width           =   375
      End
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   6555
      Width           =   3975
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quit E-Diary"
      Height          =   210
      Index           =   7
      Left            =   6960
      MouseIcon       =   "Main.frx":9BC6D
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6555
      Width           =   855
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If (X >= 1155 And Y >= 0) And (X <= 6720 And Y <= 300) Then
    If Button = vbLeftButton Then Call DragIt(Me.hwnd)
End If

End Sub

Sub Reset_Captions()

For i = 0 To 7
    lblCaption(i).FontUnderline = False
Next

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Reset_Captions

End Sub

Private Sub lblCaption_Click(Index As Integer)
Dim ans

Select Case Index

    Case 0
        Load Diary
        Diary.Show , Me
        
    Case 1
        Load AddressBook
        AddressBook.Show , Me
    Case 2
        Load Notes
        Notes.Show , Me
    Case 3
        Load Reminders
        Reminders.Show , Me
    Case 4
        Load Alarm
        Alarm.Show , Me
    Case 5
        Load Financial
        Financial.Show , Me
    Case 6
        Load Profile
        Profile.Show , Me
        
    Case 7
    
        ans = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Quit E-Diary 2008?")
        
        If ans = vbNo Then Exit Sub
    
        End

End Select

End Sub

Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If lblCaption(Index).FontUnderline = True Then Exit Sub

Reset_Captions
lblCaption(Index).FontUnderline = True

End Sub

Private Sub picChoices_Click(Index As Integer)

lblCaption_Click Index

End Sub

Private Sub picChoices_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If lblCaption(Index).FontUnderline = True Then Exit Sub

Reset_Captions
lblCaption(Index).FontUnderline = True

End Sub

Private Sub picPick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Reset_Captions

End Sub
