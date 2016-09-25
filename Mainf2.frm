VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Table Management System"
   ClientHeight    =   7305
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   11610
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Mainf2.frx":0000
   ScaleHeight     =   7305
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSubD 
      Height          =   1455
      Left            =   240
      Picture         =   "Mainf2.frx":734B
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox picStaffD 
      Height          =   1455
      Left            =   2160
      Picture         =   "Mainf2.frx":8C63
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox picLabV 
      Height          =   1455
      Left            =   9840
      Picture         =   "Mainf2.frx":A2E3
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox picClassV 
      Height          =   1455
      Left            =   7920
      Picture         =   "Mainf2.frx":B532
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox picStaffV 
      Height          =   1455
      Left            =   6000
      Picture         =   "Mainf2.frx":C84B
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox picTime 
      Height          =   1455
      Left            =   4080
      Picture         =   "Mainf2.frx":DBB7
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStaffV 
         Caption         =   "&View Staff Time Table"
      End
      Begin VB.Menu mnuClassV 
         Caption         =   "View &Class Time Table"
      End
      Begin VB.Menu mnuLabV 
         Caption         =   "View &Lab Time Table"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuSubD 
         Caption         =   "&Subject Details"
      End
      Begin VB.Menu mnuStaffD 
         Caption         =   "S&taff Details"
      End
      Begin VB.Menu mnuCreateTT 
         Caption         =   "&Create Timetable"
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuClassV_Click()
Unload Me
Form7.Show
End Sub

Private Sub mnuClose_Click()
Dim a As Integer
a = MsgBox("Do you want to exit ?", vbOKCancel, "Message")
If a = 1 Then
 End
Else
 Exit Sub
End If
End Sub

Private Sub mnuCreateTT_Click()
Unload Me
Form5.Show
End Sub

Private Sub mnuStaffD_Click()
Unload Me
Form4.Show
End Sub

Private Sub mnuSubD_Click()
Unload Me
Form3.Show
End Sub

Private Sub picClassV_Click()
Unload Me
Form7.Show

End Sub

Private Sub picTime_Click()
Unload Me
Form8.Show
End Sub

Private Sub picStaffD_Click()
Unload Me
Form4.Show
End Sub

Private Sub picSubD_Click()
Unload Me
Form3.Show
End Sub

Private Sub cmdQuit_Click()
Dim a As Integer
a = MsgBox("Do you want to exit ?", vbOKCancel, "Message")
If a = 1 Then
 End
Else
 Exit Sub
End If

End Sub

