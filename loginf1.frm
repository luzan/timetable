VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Timetable Management System"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "loginf1.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
Dim a As Integer
a = MsgBox("Do you want to exit ?", vbOKCancel, "Message")
If a = 1 Then
 End
Else
 Exit Sub
End If

End Sub

Private Sub cmdLogin_Click()
If txtPassword.Text = "admin" Then
    
    Unload Me
    Form2.Show
Else
    MsgBox "Invalid Password,Try Again!", , "Login"
    txtPassword.Text = ""
    txtPassword.SetFocus
End If

End Sub

Private Sub Form_Load()

txtUsername.Enabled = False
txtUsername.Text = "Administration"
End Sub

