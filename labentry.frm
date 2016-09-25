VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00008000&
   Caption         =   "Form6"
   ClientHeight    =   6540
   ClientLeft      =   225
   ClientTop       =   990
   ClientWidth     =   10410
   LinkTopic       =   "Form6"
   ScaleHeight     =   6540
   ScaleWidth      =   10410
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   360
      Top             =   4080
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "select * from staffentry"
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
      Left            =   360
      Top             =   3000
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "select * from labentry"
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
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   7800
      TabIndex        =   24
      Text            =   " "
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   7800
      TabIndex        =   23
      Text            =   " "
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   7800
      TabIndex        =   22
      Text            =   " "
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4800
      TabIndex        =   21
      Text            =   " "
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox Combo0 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Text            =   " "
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4800
      TabIndex        =   10
      Text            =   " "
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "labentry.frx":0000
      Left            =   4800
      List            =   "labentry.frx":0002
      TabIndex        =   9
      Text            =   " "
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   7800
      TabIndex        =   8
      Text            =   " "
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   7800
      TabIndex        =   7
      Text            =   " "
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   7800
      TabIndex        =   6
      Text            =   " "
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox Combo12 
      Height          =   315
      Left            =   7800
      TabIndex        =   5
      Text            =   " "
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   7800
      TabIndex        =   4
      Text            =   " "
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      Left            =   7800
      TabIndex        =   3
      Text            =   " "
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ADD"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   " UPDATE"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   " NEXT"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Day3"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   28
      Top             =   4440
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Session"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   27
      Top             =   3960
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Lab name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   26
      Top             =   4920
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Class"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   25
      Top             =   4440
      Width           =   810
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Staff Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Day2"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   19
      Top             =   2760
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Day1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   18
      Top             =   1080
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Session"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   17
      Top             =   600
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Session"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   16
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lab Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   15
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Class"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   14
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Lab name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   13
      Top             =   3240
      Width           =   810
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Class"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   12
      Top             =   2760
      Width           =   810
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Adodc1.RecordSource = "select * from labentry;"
Adodc1.Refresh
Adodc2.RecordSource = "select * from staffentry;"
Adodc2.Refresh
End Sub

Public Sub init()

Combo1.AddItem "Sonday"
Combo1.AddItem "Monday"
Combo1.AddItem "Tuesday"
Combo1.AddItem "Wednesday"
Combo1.AddItem "Thursday"
Combo1.AddItem "Friday"

Combo2.AddItem "Sonday"
Combo2.AddItem "Monday"
Combo2.AddItem "Tuesday"
Combo2.AddItem "Wednesday"
Combo2.AddItem "Thursday"
Combo2.AddItem "Friday"

Combo3.AddItem "Sonday"
Combo3.AddItem "Monday"
Combo3.AddItem "Tuesday"
Combo3.AddItem "Wednesday"
Combo3.AddItem "Thursday"
Combo3.AddItem "Friday"

Combo6.AddItem "Fundamental Lab"
Combo6.AddItem "Network Lab"
Combo6.AddItem "OS Lab"
Combo6.AddItem "VB Lab"
Combo6.AddItem "C Lab"
Combo6.AddItem "C++ Lab"
Combo6.AddItem "Multimedia Lab"
Combo6.AddItem "SE Lab"
Combo6.AddItem "Graphics Lab"
Combo6.AddItem "DSA Lab"
Combo6.AddItem "AutoCAD Lab"
Combo6.AddItem "Internet Lab"
Combo6.AddItem "CRM Lab"

Combo9.AddItem "Fundamental Lab"
Combo9.AddItem "Network Lab"
Combo9.AddItem "OS Lab"
Combo9.AddItem "VB Lab"
Combo9.AddItem "C Lab"
Combo9.AddItem "C++ Lab"
Combo9.AddItem "Multimedia Lab"
Combo9.AddItem "SE Lab"
Combo9.AddItem "Graphics Lab"
Combo9.AddItem "DSA Lab"
Combo9.AddItem "AutoCAD Lab"
Combo9.AddItem "Internet Lab"
Combo9.AddItem "CRM Lab"

Combo12.AddItem "Fundamental Lab"
Combo12.AddItem "Network Lab"
Combo12.AddItem "OS Lab"
Combo12.AddItem "VB Lab"
Combo12.AddItem "C Lab"
Combo12.AddItem "C++ Lab"
Combo12.AddItem "Multimedia Lab"
Combo12.AddItem "SE Lab"
Combo12.AddItem "Graphics Lab"
Combo12.AddItem "DSA Lab"
Combo12.AddItem "AutoCAD Lab"
Combo12.AddItem "Internet Lab"
Combo12.AddItem "CRM Lab"

Combo4.AddItem "First Shift"
Combo4.AddItem "Second Shift"

Combo7.AddItem "First Shift"
Combo7.AddItem "Second Shift"

Combo10.AddItem "First Shift"
Combo10.AddItem "Second Shift"

Combo5.AddItem "Ia"
Combo5.AddItem "Ib"
Combo5.AddItem "Ic"
Combo5.AddItem "Id"
Combo5.AddItem "Ie"

Combo8.AddItem "Ia"
Combo8.AddItem "Ib"
Combo8.AddItem "Ic"
Combo8.AddItem "Id"
Combo8.AddItem "Ie"

Combo11.AddItem "Ia"
Combo11.AddItem "Ib"
Combo11.AddItem "Ic"
Combo11.AddItem "Id"
Combo11.AddItem "Ie"

End Sub

Public Sub assign()
Adodc1.Recordset(0) = Combo0.Text
Adodc1.Recordset(1) = Combo1.Text
Adodc1.Recordset(2) = Combo4.Text
Adodc1.Recordset(3) = Combo5.Text
Adodc1.Recordset(4) = Combo6.Text
Adodc1.Recordset(5) = Combo2.Text
Adodc1.Recordset(6) = Combo7.Text
Adodc1.Recordset(7) = Combo8.Text
Adodc1.Recordset(8) = Combo9.Text
Adodc1.Recordset(9) = Combo3.Text
Adodc1.Recordset(10) = Combo10.Text
Adodc1.Recordset(11) = Combo11.Text
Adodc1.Recordset(12) = Combo12.Text


End Sub

Private Sub Combo6_GotFocus()
Do While Not Adodc1.Recordset.EOF
If (((Combo1.Text = Adodc1.Recordset(1)) And (Combo4.Text = Adodc1.Recordset(2)) And (Combo6.Text = Adodc1.Recordset(4))) Or (Combo1.Text = Adodc1.Recordset(1)) And (Combo4.Text = Adodc1.Recordset(2)) And (Combo5.Text = Adodc1.Recordset(3))) Then
MsgBox "Please change the session, already alloted!"
End If
Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub Combo9_GotFocus()
Do While Not Adodc1.Recordset.EOF
If (((Combo2.Text = Adodc1.Recordset(5)) And (Combo7.Text = Adodc1.Recordset(6)) And (Combo9.Text = Adodc1.Recordset(8))) Or (Combo2.Text = Adodc1.Recordset(5)) And (Combo7.Text = Adodc1.Recordset(6)) And (Combo8.Text = Adodc1.Recordset(7))) Then
MsgBox "Please change the session, already alloted!"
End If
Adodc1.Recordset.MoveNext
Loop
End Sub
Private Sub Combo12_GotFocus()
Do While Not Adodc1.Recordset.EOF
If (((Combo3.Text = Adodc1.Recordset(9)) And (Combo10.Text = Adodc1.Recordset(10)) And (Combo12.Text = Adodc1.Recordset(12))) Or (Combo3.Text = Adodc1.Recordset(9)) And (Combo10.Text = Adodc1.Recordset(10)) And (Combo11.Text = Adodc1.Recordset(11))) Then
MsgBox "Please change the session, already alloted!"
End If
Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub Command3_Click()
Unload Me
Form5.Show
End Sub
Private Sub Command2_Click()
assign
Adodc1.Recordset.Update
MsgBox "Record updated"
End Sub
Private Sub Command1_Click()
Do While Not Adodc2.Recordset.EOF
Combo0.AddItem Adodc2.Recordset(1)
Adodc2.Recordset.MoveNext
Loop
init
Adodc1.Recordset.AddNew
End Sub

