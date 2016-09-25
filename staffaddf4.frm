VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timetable Management System - Staff Details Form"
   ClientHeight    =   6360
   ClientLeft      =   375
   ClientTop       =   1230
   ClientWidth     =   8385
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8385
   Begin VB.CommandButton remove 
      Caption         =   "DEL"
      Height          =   495
      Left            =   6600
      TabIndex        =   17
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6360
      OLEDropMode     =   1  'Manual
      Picture         =   "staffaddf4.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   735
      TabIndex        =   13
      ToolTipText     =   "To Main Menu"
      Top             =   240
      Width           =   735
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
      Left            =   6600
      TabIndex        =   12
      ToolTipText     =   "MoveNext"
      Top             =   1680
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
      Left            =   6600
      TabIndex        =   11
      ToolTipText     =   "Movelast"
      Top             =   3120
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
      Left            =   6600
      TabIndex        =   10
      ToolTipText     =   "MoveFirst"
      Top             =   2640
      Width           =   495
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
      Left            =   6600
      TabIndex        =   9
      ToolTipText     =   "MovePrevious"
      Top             =   2160
      Width           =   495
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
      Left            =   5040
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
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
      Left            =   3660
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
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
      Left            =   2280
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox comSub3 
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox comSub2 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox comSub1 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtStaffID 
      Height          =   300
      Left            =   2880
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtStaffname 
      Height          =   300
      Left            =   2880
      TabIndex        =   1
      Top             =   2280
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   2400
      Top             =   5520
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select *from subjectentry"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   0
      Top             =   5520
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "select *from staffentry"
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
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   240
      X2              =   7080
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   600
      X2              =   600
      Y1              =   480
      Y2              =   5160
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "Staff ID                :"
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
      Left            =   840
      TabIndex        =   16
      Top             =   1680
      Width           =   1860
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      Caption         =   "Subject Handled :"
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
      Left            =   840
      TabIndex        =   15
      Top             =   3120
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "Staff Name          :"
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
      Left            =   840
      TabIndex        =   14
      Top             =   2280
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Staff Details"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2700
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub clear()
txtStaffID.Text = ""
txtStaffname.Text = ""
'txtDepartment.Text = ""
comSub1.Text = ""
comSub2.Text = ""
comSub3.Text = ""
End Sub
Private Sub cmdAddnew_Click()
Dim aa As Integer
If txtStaffID.Text <> "" And txtStaffname.Text <> "" Then
aa = MsgBox("Are You sure to add data", vbOKCancel, "Add Message")
If aa = 1 Then

Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = txtStaffID.Text

Adodc1.Recordset.Fields(1) = txtStaffname.Text

Adodc1.Recordset.Fields(2) = comSub1.Text
Adodc1.Recordset.Fields(3) = comSub2.Text
Adodc1.Recordset.Fields(4) = comSub3.Text

Adodc1.Recordset.Update
Adodc1.Refresh

aa = MsgBox("Data successfully inserted", vbOKOnly, "Success Message")
clear
End If
End If
End Sub
Private Sub cmdMFirst_Click()
Adodc1.Recordset.MoveFirst
txtStaffID.Text = Adodc1.Recordset.Fields(0)
txtStaffname.Text = Adodc1.Recordset.Fields(1)
comSub1.Text = Adodc1.Recordset.Fields(2)
End Sub

Private Sub cmdMLast_Click()
Adodc1.Recordset.MoveLast
txtStaffID.Text = Adodc1.Recordset.Fields(0)
txtStaffname.Text = Adodc1.Recordset.Fields(1)
comSub1.Text = Adodc1.Recordset.Fields(2)
End Sub

Private Sub cmdMNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveLast
End If
txtStaffID.Text = Adodc1.Recordset.Fields(0)
txtStaffname.Text = Adodc1.Recordset.Fields(1)
comSub1.Text = Adodc1.Recordset.Fields(2)
End Sub

Private Sub cmdMPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst
End If
txtStaffID.Text = Adodc1.Recordset.Fields(0)
txtStaffname.Text = Adodc1.Recordset.Fields(1)
comSub1.Text = Adodc1.Recordset.Fields(2)
End Sub

Private Sub cmdNext_Click()
Unload Me
Form5.Show
End Sub

Private Sub cmdUpdate_Click()
Adodc1.Recordset.Fields(0) = txtStaffID.Text
Adodc1.Recordset.Fields(1) = txtStaffname.Text
Adodc1.Recordset.Fields(2) = comSub1.Text
Adodc1.Recordset.Fields(3) = comSub2.Text
Adodc1.Recordset.Fields(4) = comSub3.Text

Adodc1.Recordset.Update
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "select * from staffentry;"
Adodc1.Refresh
Adodc2.RecordSource = "select * from subjectentry;"
Adodc2.Refresh

Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
comSub1.AddItem Adodc2.Recordset(1)
comSub2.AddItem Adodc2.Recordset(1)
comSub3.AddItem Adodc2.Recordset(1)
Adodc2.Recordset.MoveNext
Loop

End Sub


Private Sub Picture1_Click()
Unload Me
Form2.Show

End Sub

Private Sub remove_Click()
Adodc1.Recordset.Delete

End Sub
