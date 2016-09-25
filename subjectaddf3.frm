VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timetable Management System - Subject Details Form"
   ClientHeight    =   5790
   ClientLeft      =   585
   ClientTop       =   1965
   ClientWidth     =   8760
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8760
   Begin VB.CommandButton remove 
      Caption         =   "DEL"
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6240
      OLEDropMode     =   1  'Manual
      Picture         =   "subjectaddf3.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   735
      TabIndex        =   11
      ToolTipText     =   "To Main Menu"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdMPrevious 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      ToolTipText     =   "Move Previous"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdMFirst 
      Caption         =   "|<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      ToolTipText     =   "Move First"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdMLast 
      Caption         =   ">>|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      ToolTipText     =   "Move last"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdMNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Move Next"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtFaculty 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtSub 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtSubID 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddnew 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2040
      Top             =   4560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\projectsave\timetable.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\projectsave\timetable.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from subjectentry"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   480
      X2              =   480
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   6960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "Faculty      :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1080
      TabIndex        =   14
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Subject ID :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   13
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "Subject      :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Subject Details"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   525
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3105
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStaffTT 
         Caption         =   "View &Staff Time Table"
      End
      Begin VB.Menu mnuClassTT 
         Caption         =   "View &Class Time Table"
      End
      Begin VB.Menu mnuLabTT 
         Caption         =   "View &Lab Time Table"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuSubD 
         Caption         =   "S&ubject Details"
      End
      Begin VB.Menu mnuStaffD 
         Caption         =   "S&taff Details"
      End
      Begin VB.Menu mnuCTT 
         Caption         =   "C&reate Timetable"
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub clear()
txtSubID.Text = ""
txtSub.Text = ""
txtFaculty.Text = ""

End Sub

Private Sub cmdAddnew_Click()
Dim aa As Integer
If txtSubID.Text <> "" And txtSub.Text <> "" And txtFaculty.Text <> "" Then
aa = MsgBox("Are You sure to add data", vbOKCancel, "Add Message")
If aa = 1 Then

Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = txtSubID.Text

Adodc1.Recordset.Fields(1) = txtSub.Text

Adodc1.Recordset.Fields(2) = txtFaculty.Text

Adodc1.Recordset.Update
aa = MsgBox("Data successfully inserted", vbOKOnly, "Success Message")
clear
Adodc1.Refresh

End If
End If

End Sub

Private Sub cmdMFirst_Click()
Adodc1.Recordset.MoveFirst
txtSubID.Text = Adodc1.Recordset.Fields(0)
txtSub.Text = Adodc1.Recordset.Fields(1)
txtFaculty.Text = Adodc1.Recordset.Fields(2)
End Sub

Private Sub cmdMLast_Click()
Adodc1.Recordset.MoveLast
txtSubID.Text = Adodc1.Recordset.Fields(0)
txtSub.Text = Adodc1.Recordset.Fields(1)
txtFaculty.Text = Adodc1.Recordset.Fields(2)
End Sub

Private Sub cmdMNext_Click()
Adodc1.Recordset.MoveNext

If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveLast
End If
txtSubID.Text = Adodc1.Recordset.Fields(0)
txtSub.Text = Adodc1.Recordset.Fields(1)
txtFaculty.Text = Adodc1.Recordset.Fields(2)

End Sub

Private Sub cmdMPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst
End If

txtSubID.Text = Adodc1.Recordset.Fields(0)
txtSub.Text = Adodc1.Recordset.Fields(1)
txtFaculty.Text = Adodc1.Recordset.Fields(2)
End Sub

Private Sub cmdNext_Click()
Form4.Show
Unload Form3

End Sub

Private Sub cmdUpdate_Click()
Adodc1.Recordset.Fields(0) = txtSubID.Text
Adodc1.Recordset.Fields(1) = txtSub.Text
Adodc1.Recordset.Fields(2) = txtFaculty.Text

Adodc1.Recordset.Update


End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "select * from subjectentry;"
Adodc1.Refresh
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

Private Sub Picture1_Click()
Unload Me
Form2.Show
End Sub


Private Sub remove_Click()
Adodc1.Recordset.Delete
End Sub
