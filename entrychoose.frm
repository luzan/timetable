VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00008000&
   Caption         =   "Form8"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form8"
   ScaleHeight     =   4800
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLab 
      Caption         =   "Lab"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      MaskColor       =   &H8000000D&
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdClass 
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H8000000D&
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.PictureBox picClass 
      Height          =   1455
      Left            =   840
      Picture         =   "entrychoose.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox picLab 
      Height          =   1455
      Left            =   3360
      Picture         =   "entrychoose.frx":17E5
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClass_Click()
Unload Me
Form7.Show
End Sub

Private Sub cmdLab_Click()
Unload Me
Form6.Show
End Sub

Private Sub picClass_Click()
Unload Me
Form7.Show

End Sub

Private Sub picLab_Click()
Unload Me
Form6.Show
End Sub
