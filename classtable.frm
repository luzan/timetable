VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00008000&
   Caption         =   "Form7"
   ClientHeight    =   7530
   ClientLeft      =   555
   ClientTop       =   1635
   ClientWidth     =   12045
   LinkTopic       =   "Form7"
   ScaleHeight     =   7530
   ScaleWidth      =   12045
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   735
      Left            =   7320
      Top             =   7560
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   9720
      Top             =   7320
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "select * from classtt"
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
      Height          =   615
      Left            =   5040
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
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
      RecordSource    =   "select * from entry"
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
   Begin VB.TextBox Text48 
      Height          =   375
      Left            =   10440
      TabIndex        =   64
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text47 
      Height          =   375
      Left            =   9240
      TabIndex        =   63
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text46 
      Height          =   375
      Left            =   8040
      TabIndex        =   62
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text45 
      Height          =   375
      Left            =   6840
      TabIndex        =   61
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text44 
      Height          =   375
      Left            =   5640
      TabIndex        =   60
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text43 
      Height          =   375
      Left            =   4440
      TabIndex        =   59
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text42 
      Height          =   375
      Left            =   3240
      TabIndex        =   58
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text41 
      Height          =   375
      Left            =   2040
      TabIndex        =   57
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text40 
      Height          =   375
      Left            =   10440
      TabIndex        =   56
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text39 
      Height          =   375
      Left            =   9240
      TabIndex        =   55
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text38 
      Height          =   375
      Left            =   8040
      TabIndex        =   54
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text37 
      Height          =   375
      Left            =   6840
      TabIndex        =   53
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text36 
      Height          =   375
      Left            =   5640
      TabIndex        =   52
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   37
      Text            =   " "
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3240
      TabIndex        =   36
      Text            =   " "
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4440
      TabIndex        =   35
      Text            =   " "
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5640
      TabIndex        =   34
      Text            =   " "
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6840
      TabIndex        =   33
      Text            =   " "
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8040
      TabIndex        =   32
      Text            =   " "
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   9240
      TabIndex        =   31
      Text            =   " "
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   10440
      TabIndex        =   30
      Text            =   " "
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2040
      TabIndex        =   29
      Text            =   "   "
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   3240
      TabIndex        =   28
      Text            =   " "
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Text            =   " "
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Text            =   " "
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   6840
      TabIndex        =   25
      Text            =   " "
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   8040
      TabIndex        =   24
      Text            =   " "
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   9240
      TabIndex        =   23
      Text            =   " "
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   10440
      TabIndex        =   22
      Text            =   " "
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Text            =   " "
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Text            =   " "
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Text            =   " "
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Text            =   " "
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   6840
      TabIndex        =   17
      Text            =   " "
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   8040
      TabIndex        =   16
      Text            =   " "
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   9240
      TabIndex        =   15
      Text            =   " "
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   10440
      TabIndex        =   14
      Text            =   " "
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Text            =   " "
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Text            =   " "
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text27 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Text            =   " "
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text28 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Text            =   " "
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text29 
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Text            =   " "
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text30 
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Text            =   " "
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text31 
      Height          =   375
      Left            =   9240
      TabIndex        =   7
      Text            =   " "
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text32 
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Text            =   " "
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text33 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Text            =   " "
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Text            =   " "
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text35 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Text            =   " "
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Print"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   " End"
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Sunday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   65
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Hour8"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   51
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   " CLASS TIMETABLE"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   4080
      TabIndex        =   50
      Top             =   120
      Width           =   3600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Hour1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   49
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Hour2"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   48
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Hour3"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   47
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Hour4"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   46
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Hour5"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   45
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Hour6"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   44
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Hour7"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   43
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Monday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   42
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Tuesday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   3360
      Width           =   1140
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Wednesday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   40
      Top             =   4200
      Width           =   1140
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Thursday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   5040
      Width           =   1140
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Friday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   5880
      Width           =   1140
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As String
Dim d, e As Integer

Private Sub Command1_Click()
Dim a As Integer

On Error GoTo soln
c = InputBox("enter the Class:")
Adodc2.Recordset.AddNew
Adodc2.Recordset(0) = Trim(c)
Do While Not Adodc1.Recordset.EOF
If (Adodc1.Recordset(44) = c) Then
d = Val(Adodc1.Recordset(8))
Adodc2.Recordset(d) = Adodc1.Recordset(26)

Exit Sub
soln:
a = MsgBox("Empty class,Try Again!", vbOKCancel, "Error Message")

If d = "1" Then
Text1.Text = Adodc2.Recordset(d)
ElseIf d = "2" Then
Text2.Text = Adodc2.Recordset(d)
ElseIf d = "3" Then
Text3.Text = Adodc2.Recordset(d)
ElseIf d = "4" Then
Text4.Text = Adodc2.Recordset(d)
ElseIf d = "5" Then
Text5.Text = Adodc2.Recordset(d)
ElseIf d = "6" Then
Text6.Text = Adodc2.Recordset(d)
ElseIf d = "7" Then
Text7.Text = Adodc2.Recordset(d)
ElseIf d = "8" Then
Text8.Text = Adodc2.Recordset(d)
End If
End If

If (Adodc1.Recordset(45) = c) Then
d = Val(Adodc1.Recordset(9))
Adodc2.Recordset(d) = Adodc1.Recordset(27)
If d = "1" Then
Text1.Text = Adodc2.Recordset(d)
ElseIf d = "2" Then
Text2.Text = Adodc2.Recordset(d)
ElseIf d = "3" Then
Text3.Text = Adodc2.Recordset(d)
ElseIf d = "4" Then
Text4.Text = Adodc2.Recordset(d)
ElseIf d = "5" Then
Text5.Text = Adodc2.Recordset(d)
ElseIf d = "6" Then
Text6.Text = Adodc2.Recordset(d)
ElseIf d = "7" Then
Text7.Text = Adodc2.Recordset(d)
ElseIf d = "8" Then
Text8.Text = Adodc2.Recordset(d)
End If
End If
d = Val(Adodc1.Recordset(10))
If d = "0" Then
    Adodc3.Recordset.MoveFirst
    Do While Not Adodc3.Recordset.EOF
        If ((Adodc3.Recordset(2) = "First Shift") And (Adodc3.Recordset(1) = "Sunday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text1 = Adodc3.Recordset(4)
        Text2 = Adodc3.Recordset(4)
        Text3 = Adodc3.Recordset(4)
        Text4 = Adodc3.Recordset(4)
        Adodc2.Recordset(1) = Adodc3.Recordset(4)
        Adodc2.Recordset(2) = Adodc3.Recordset(4)
        Adodc2.Recordset(3) = Adodc3.Recordset(4)
        Adodc2.Recordset(4) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "First Shift") And (Adodc3.Recordset(5) = "Sunday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text1 = Adodc3.Recordset(8)
        Text2 = Adodc3.Recordset(8)
        Text3 = Adodc3.Recordset(8)
        Text4 = Adodc3.Recordset(8)
        Adodc2.Recordset(1) = Adodc3.Recordset(8)
        Adodc2.Recordset(2) = Adodc3.Recordset(8)
        Adodc2.Recordset(3) = Adodc3.Recordset(8)
        Adodc2.Recordset(4) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "First Shift") And (Adodc3.Recordset(9) = "Sunday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text1 = Adodc3.Recordset(12)
        Text2 = Adodc3.Recordset(12)
        Text3 = Adodc3.Recordset(12)
        Text4 = Adodc3.Recordset(12)
        Adodc2.Recordset(1) = Adodc3.Recordset(12)
        Adodc2.Recordset(2) = Adodc3.Recordset(12)
        Adodc2.Recordset(3) = Adodc3.Recordset(12)
        Adodc2.Recordset(4) = Adodc3.Recordset(12)
        ElseIf ((Adodc3.Recordset(2) = "Second Shift") And (Adodc3.Recordset(1) = "Sunday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text5 = Adodc3.Recordset(4)
        Text6 = Adodc3.Recordset(4)
        Text7 = Adodc3.Recordset(4)
        Text8 = Adodc3.Recordset(4)
        Adodc2.Recordset(5) = Adodc3.Recordset(4)
        Adodc2.Recordset(6) = Adodc3.Recordset(4)
        Adodc2.Recordset(7) = Adodc3.Recordset(4)
        Adodc2.Recordset(8) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Second Shift") And (Adodc3.Recordset(5) = "Sunday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text5 = Adodc3.Recordset(8)
        Text6 = Adodc3.Recordset(8)
        Text7 = Adodc3.Recordset(8)
        Text8 = Adodc3.Recordset(8)
        Adodc2.Recordset(5) = Adodc3.Recordset(8)
        Adodc2.Recordset(6) = Adodc3.Recordset(8)
        Adodc2.Recordset(7) = Adodc3.Recordset(8)
        Adodc2.Recordset(7) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Second Shift") And (Adodc3.Recordset(9) = "Sunday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text5 = Adodc3.Recordset(12)
        Text6 = Adodc3.Recordset(12)
        Text7 = Adodc3.Recordset(12)
        Text8 = Adodc3.Recordset(12)
        Adodc2.Recordset(5) = Adodc3.Recordset(12)
        Adodc2.Recordset(6) = Adodc3.Recordset(12)
        Adodc2.Recordset(7) = Adodc3.Recordset(12)
        Adodc2.Recordset(7) = Adodc3.Recordset(12)
        End If
    Adodc3.Recordset.MoveNext
    Loop
ElseIf (Adodc1.Recordset(46) = c) Then
d = Val(Adodc1.Recordset(10))
Adodc2.Recordset(d) = Adodc1.Recordset(28)
If d = "1" Then
Text1.Text = Adodc2.Recordset(d)
ElseIf d = "2" Then
Text2.Text = Adodc2.Recordset(d)
ElseIf d = "3" Then
Text3.Text = Adodc2.Recordset(d)
ElseIf d = "4" Then
Text4.Text = Adodc2.Recordset(d)
ElseIf d = "5" Then
Text5.Text = Adodc2.Recordset(d)
ElseIf d = "6" Then
Text6.Text = Adodc2.Recordset(d)
ElseIf d = "7" Then
Text7.Text = Adodc2.Recordset(d)
ElseIf d = "8" Then
Text8.Text = Adodc2.Recordset(d)
End If
End If

''''''''''''''''''''''''''
'call monday
Call monday

''''''''''''''''''''''''''
'call tuesday
Call tuesday

''''''''''''''''''''''''''
'call wednesday
Call wednesday

''''''''''''''''''''''''''
'call thrusday
Call thursday

''''''''''''''''''''''''''
'call friday
Call friday

''''''''''''''''''''''''''
Adodc1.Recordset.MoveNext
Loop
Adodc2.Recordset.Update
Adodc2.Recordset.Close

End Sub


Public Sub tuesday()
 If (Adodc1.Recordset(50) = c) Then
d = Val(Adodc1.Recordset(14))
e = d + 14
Adodc2.Recordset(e) = Adodc1.Recordset(32)
If d = "1" Then
Text17.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text18.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text19.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text20.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text21.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text22.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text23.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text24.Text = Adodc2.Recordset(e)
End If
End If

If (Adodc1.Recordset(51) = c) Then
d = Val(Adodc1.Recordset(15))
e = d + 14
Adodc2.Recordset(e) = Adodc1.Recordset(33)
If d = "1" Then
Text17.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text18.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text19.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text20.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text21.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text22.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text23.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text24.Text = Adodc2.Recordset(e)
End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'text boxes are managed up to here as well as Adodc3.Recordset db
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

d = Val(Adodc1.Recordset(16))
If d = "0" Then
    Adodc3.Recordset.MoveFirst
    Do While Not Adodc3.Recordset.EOF
        If ((Adodc3.Recordset(2) = "Fisrt Shift") And (Adodc3.Recordset(1) = "Tuesday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text17 = Adodc3.Recordset(4)
        Text18 = Adodc3.Recordset(4)
        Text19 = Adodc3.Recordset(4)
        Text20 = Adodc3.Recordset(4)
        Adodc2.Recordset(17) = Adodc3.Recordset(4)
        Adodc2.Recordset(18) = Adodc3.Recordset(4)
        Adodc2.Recordset(19) = Adodc3.Recordset(4)
        Adodc2.Recordset(20) = Adodc3.Recordset(4)
        End If
        If ((Adodc3.Recordset(6) = "Fisrt Shift") And (Adodc3.Recordset(5) = "Tuesday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text17 = Adodc3.Recordset(8)
        Text18 = Adodc3.Recordset(8)
        Text19 = Adodc3.Recordset(8)
        Text20 = Adodc3.Recordset(8)
        Adodc2.Recordset(17) = Adodc3.Recordset(8)
        Adodc2.Recordset(18) = Adodc3.Recordset(8)
        Adodc2.Recordset(19) = Adodc3.Recordset(8)
        Adodc2.Recordset(20) = Adodc3.Recordset(8)
        End If
        If ((Adodc3.Recordset(10) = "Fisrt Shift") And (Adodc3.Recordset(9) = "Tuesday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text17 = Adodc3.Recordset(12)
        Text18 = Adodc3.Recordset(12)
        Text19 = Adodc3.Recordset(12)
        Text20 = Adodc3.Recordset(12)
        Adodc2.Recordset(17) = Adodc3.Recordset(12)
        Adodc2.Recordset(18) = Adodc3.Recordset(12)
        Adodc2.Recordset(19) = Adodc3.Recordset(12)
        Adodc2.Recordset(20) = Adodc3.Recordset(12)
        End If
        If (((Adodc3.Recordset(2) = "Second Shift") And (Adodc3.Recordset(1) = "Tuesday")) And (Adodc3.Recordset(3) = Trim(c))) Then
        Text21 = Adodc3.Recordset(4)
        Text22 = Adodc3.Recordset(4)
        Text23 = Adodc3.Recordset(4)
        Text24 = Adodc3.Recordset(4)
        Adodc2.Recordset(21) = Adodc3.Recordset(4)
        Adodc2.Recordset(22) = Adodc3.Recordset(4)
        Adodc2.Recordset(23) = Adodc3.Recordset(4)
        Adodc2.Recordset(24) = Adodc3.Recordset(4)
        End If
        If (((Adodc3.Recordset(6) = "Second Shift") And (Adodc3.Recordset(5) = "Tuesday")) And (Adodc3.Recordset(7) = Trim(c))) Then
        Text21 = Adodc3.Recordset(8)
        Text22 = Adodc3.Recordset(8)
        Text23 = Adodc3.Recordset(8)
        Text24 = Adodc3.Recordset(8)
        Adodc2.Recordset(21) = Adodc3.Recordset(8)
        Adodc2.Recordset(22) = Adodc3.Recordset(8)
        Adodc2.Recordset(23) = Adodc3.Recordset(8)
        Adodc2.Recordset(24) = Adodc3.Recordset(8)
        End If
        If (((Adodc3.Recordset(10) = "Second Shift") And (Adodc3.Recordset(9) = "Tuesday")) And (Adodc3.Recordset(11) = Trim(c))) Then
        Text21 = Adodc3.Recordset(12)
        Text22 = Adodc3.Recordset(12)
        Text23 = Adodc3.Recordset(12)
        Text24 = Adodc3.Recordset(12)
        Adodc2.Recordset(21) = Adodc3.Recordset(12)
        Adodc2.Recordset(22) = Adodc3.Recordset(12)
        Adodc2.Recordset(23) = Adodc3.Recordset(12)
        Adodc2.Recordset(24) = Adodc3.Recordset(12)
        End If
        
     Adodc3.Recordset.MoveNext
    Loop
End If
If (Adodc1.Recordset(52) = c) Then
d = Val(Adodc1.Recordset(16))
e = d + 14
Adodc2.Recordset(e) = Adodc1.Recordset(34)
If d = "1" Then
Text17.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text18.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text19.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text20.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text21.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text22.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text23.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text24.Text = Adodc2.Recordset(e)
End If
End If
End Sub
Public Sub wednesday()

If (Adodc1.Recordset(53) = c) Then
d = Val(Adodc1.Recordset(17))
e = d + 21
Adodc2.Recordset(e) = Adodc1.Recordset(35)
If d = "1" Then
Text25.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text26.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text27.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text28.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text29.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text30.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text31.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text32.Text = Adodc2.Recordset(e)
End If
End If

If (Adodc1.Recordset(54) = c) Then
d = Val(Adodc1.Recordset(18))
e = d + 21
Adodc2.Recordset(e) = Adodc1.Recordset(36)
If d = "1" Then
Text25.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text26.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text27.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text28.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text29.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text30.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text31.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text32.Text = Adodc2.Recordset(e)
End If
End If

d = Val(Adodc1.Recordset(19))
If d = "0" Then
Adodc3.Recordset.MoveFirst
    Do While Not Adodc3.Recordset.EOF
        If ((Adodc3.Recordset(2) = "Fisrt Shift") And (Adodc3.Recordset(1) = "Wednesday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text25 = Adodc3.Recordset(4)
        Text26 = Adodc3.Recordset(4)
        Text27 = Adodc3.Recordset(4)
        Text28 = Adodc3.Recordset(4)
        Adodc2.Recordset(25) = Adodc3.Recordset(4)
        Adodc2.Recordset(26) = Adodc3.Recordset(4)
        Adodc2.Recordset(27) = Adodc3.Recordset(4)
        Adodc2.Recordset(28) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Fisrt Shift") And (Adodc3.Recordset(5) = "Wednesday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text25 = Adodc3.Recordset(8)
        Text26 = Adodc3.Recordset(8)
        Text27 = Adodc3.Recordset(8)
        Text28 = Adodc3.Recordset(8)
        Adodc2.Recordset(25) = Adodc3.Recordset(8)
        Adodc2.Recordset(26) = Adodc3.Recordset(8)
        Adodc2.Recordset(27) = Adodc3.Recordset(8)
        Adodc2.Recordset(28) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Fisrt Shift") And (Adodc3.Recordset(9) = "Wednesday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text25 = Adodc3.Recordset(12)
        Text26 = Adodc3.Recordset(12)
        Text27 = Adodc3.Recordset(12)
        Text28 = Adodc3.Recordset(12)
        Adodc2.Recordset(25) = Adodc3.Recordset(12)
        Adodc2.Recordset(26) = Adodc3.Recordset(12)
        Adodc2.Recordset(27) = Adodc3.Recordset(12)
        Adodc2.Recordset(28) = Adodc3.Recordset(12)
        ElseIf ((Adodc3.Recordset(2) = "Second Shift") And (Adodc3.Recordset(1) = "Wednesday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text29 = Adodc3.Recordset(4)
        Text30 = Adodc3.Recordset(4)
        Text31 = Adodc3.Recordset(4)
        Text32 = Adodc3.Recordset(4)
        Adodc2.Recordset(29) = Adodc3.Recordset(4)
        Adodc2.Recordset(30) = Adodc3.Recordset(4)
        Adodc2.Recordset(31) = Adodc3.Recordset(4)
        Adodc2.Recordset(32) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Second Shift") And (Adodc3.Recordset(5) = "Wednesday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text29 = Adodc3.Recordset(8)
        Text30 = Adodc3.Recordset(8)
        Text31 = Adodc3.Recordset(8)
        Text32 = Adodc3.Recordset(8)
        Adodc2.Recordset(29) = Adodc3.Recordset(8)
        Adodc2.Recordset(30) = Adodc3.Recordset(8)
        Adodc2.Recordset(31) = Adodc3.Recordset(8)
        Adodc2.Recordset(32) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Second Shift") And (Adodc3.Recordset(9) = "Wednesday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text29 = Adodc3.Recordset(12)
        Text30 = Adodc3.Recordset(12)
        Text31 = Adodc3.Recordset(12)
        Text32 = Adodc3.Recordset(12)
        Adodc2.Recordset(29) = Adodc3.Recordset(12)
        Adodc2.Recordset(30) = Adodc3.Recordset(12)
        Adodc2.Recordset(31) = Adodc3.Recordset(12)
        Adodc2.Recordset(32) = Adodc3.Recordset(12)
        End If
    Adodc3.Recordset.MoveNext
    Loop
ElseIf (Adodc1.Recordset(55) = c) Then
d = Val(Adodc1.Recordset(19))
e = d + 21
Adodc2.Recordset(e) = Adodc1.Recordset(37)
If d = "1" Then
Text25.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text26.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text27.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text28.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text29.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text30.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text31.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text32.Text = Adodc2.Recordset(e)
End If
End If
End Sub
Public Sub thursday()

If (Adodc1.Recordset(56) = c) Then
d = Val(Adodc1.Recordset(20))
e = d + 28
Adodc2.Recordset(e) = Adodc1.Recordset(38)
If d = "1" Then
Text33.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text34.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text35.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text36.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text37.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text38.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text39.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text40.Text = Adodc2.Recordset(e)
End If
End If

If (Adodc1.Recordset(57) = c) Then
d = Val(Adodc1.Recordset(21))
e = d + 28
Adodc2.Recordset(e) = Adodc1.Recordset(39)
If d = "1" Then
Text33.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text34.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text35.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text36.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text37.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text38.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text39.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text40.Text = Adodc2.Recordset(e)
End If
End If

d = Val(Adodc1.Recordset(22))
If d = "0" Then
    
    Do While Not Adodc3.Recordset.EOF
        If ((Adodc3.Recordset(2) = "Fisrt Shift") And (Adodc3.Recordset(1) = "Thrusday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text33 = Adodc3.Recordset(4)
        Text34 = Adodc3.Recordset(4)
        Text35 = Adodc3.Recordset(4)
        Text36 = Adodc3.Recordset(4)
        Adodc2.Recordset(33) = Adodc3.Recordset(4)
        Adodc2.Recordset(34) = Adodc3.Recordset(4)
        Adodc2.Recordset(35) = Adodc3.Recordset(4)
        Adodc2.Recordset(36) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Fisrt Shift") And (Adodc3.Recordset(5) = "Thrusday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text33 = Adodc3.Recordset(8)
        Text34 = Adodc3.Recordset(8)
        Text35 = Adodc3.Recordset(8)
        Text36 = Adodc3.Recordset(8)
        Adodc2.Recordset(33) = Adodc3.Recordset(8)
        Adodc2.Recordset(34) = Adodc3.Recordset(8)
        Adodc2.Recordset(35) = Adodc3.Recordset(8)
        Adodc2.Recordset(36) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Fisrt Shift") And (Adodc3.Recordset(9) = "Thrusday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text33 = Adodc3.Recordset(12)
        Text34 = Adodc3.Recordset(12)
        Text35 = Adodc3.Recordset(12)
        Text36 = Adodc3.Recordset(12)
        Adodc2.Recordset(33) = Adodc3.Recordset(12)
        Adodc2.Recordset(34) = Adodc3.Recordset(12)
        Adodc2.Recordset(35) = Adodc3.Recordset(12)
        Adodc2.Recordset(36) = Adodc3.Recordset(12)
        ElseIf ((Adodc3.Recordset(2) = "Second Shift") And (Adodc3.Recordset(1) = "Thrusday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text37 = Adodc3.Recordset(4)
        Text38 = Adodc3.Recordset(4)
        Text39 = Adodc3.Recordset(4)
        Text40 = Adodc3.Recordset(4)
        Adodc2.Recordset(37) = Adodc3.Recordset(4)
        Adodc2.Recordset(38) = Adodc3.Recordset(4)
        Adodc2.Recordset(39) = Adodc3.Recordset(4)
        Adodc2.Recordset(40) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Second Shift") And (Adodc3.Recordset(5) = "Thrusday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text37 = Adodc3.Recordset(8)
        Text38 = Adodc3.Recordset(8)
        Text39 = Adodc3.Recordset(8)
        Text40 = Adodc3.Recordset(8)
        Adodc2.Recordset(37) = Adodc3.Recordset(8)
        Adodc2.Recordset(38) = Adodc3.Recordset(8)
        Adodc2.Recordset(39) = Adodc3.Recordset(8)
        Adodc2.Recordset(40) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Second Shift") And (Adodc3.Recordset(9) = "Thrusday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text37 = Adodc3.Recordset(12)
        Text38 = Adodc3.Recordset(12)
        Text39 = Adodc3.Recordset(12)
        Text40 = Adodc3.Recordset(12)
        Adodc2.Recordset(37) = Adodc3.Recordset(12)
        Adodc2.Recordset(38) = Adodc3.Recordset(12)
        Adodc2.Recordset(39) = Adodc3.Recordset(12)
        Adodc2.Recordset(40) = Adodc3.Recordset(12)
        End If
    Adodc3.Recordset.MoveNext
    Loop
ElseIf (Adodc1.Recordset(58) = c) Then
d = Val(Adodc1.Recordset(22))
e = d + 28
Adodc2.Recordset(e) = Adodc1.Recordset(40)
If d = "1" Then
Text33.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text34.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text35.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text36.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text37.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text38.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text39.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text40.Text = Adodc2.Recordset(e)
End If
End If
End Sub
Public Sub friday()

If (Adodc1.Recordset(59) = c) Then
d = Val(Adodc1.Recordset(23))
e = d + 35
Adodc2.Recordset(e) = Adodc1.Recordset(41)
If d = "1" Then
Text41.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text42.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text43.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text44.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text45.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text46.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text47.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text48.Text = Adodc2.Recordset(e)
End If
End If

If (Adodc1.Recordset(60) = c) Then
d = Val(Adodc1.Recordset(24))
e = d + 35
Adodc2.Recordset(e) = Adodc1.Recordset(42)
If d = "1" Then
Text41.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text42.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text43.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text44.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text45.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text46.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text47.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text48.Text = Adodc2.Recordset(e)
End If
End If

d = Val(Adodc1.Recordset(25))
If d = "0" Then
    
    Do While Not Adodc3.Recordset.EOF
        If ((Adodc3.Recordset(2) = "Fisrt Shift") And (Adodc3.Recordset(1) = "Friday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text41 = Adodc3.Recordset(4)
        Text42 = Adodc3.Recordset(4)
        Text43 = Adodc3.Recordset(4)
        Text44 = Adodc3.Recordset(4)
        Adodc2.Recordset(41) = Adodc3.Recordset(4)
        Adodc2.Recordset(42) = Adodc3.Recordset(4)
        Adodc2.Recordset(43) = Adodc3.Recordset(4)
        Adodc2.Recordset(44) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Fisrt Shift") And (Adodc3.Recordset(5) = "Friday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text41 = Adodc3.Recordset(8)
        Text42 = Adodc3.Recordset(8)
        Text43 = Adodc3.Recordset(8)
        Text44 = Adodc3.Recordset(8)
        Adodc2.Recordset(41) = Adodc3.Recordset(8)
        Adodc2.Recordset(42) = Adodc3.Recordset(8)
        Adodc2.Recordset(43) = Adodc3.Recordset(8)
        Adodc2.Recordset(44) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Fisrt Shift") And (Adodc3.Recordset(9) = "Friday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text41 = Adodc3.Recordset(12)
        Text42 = Adodc3.Recordset(12)
        Text43 = Adodc3.Recordset(12)
        Text44 = Adodc3.Recordset(12)
        Adodc2.Recordset(41) = Adodc3.Recordset(12)
        Adodc2.Recordset(42) = Adodc3.Recordset(12)
        Adodc2.Recordset(43) = Adodc3.Recordset(12)
        Adodc2.Recordset(44) = Adodc3.Recordset(12)
        ElseIf ((Adodc3.Recordset(2) = "Second Shift") And (Adodc3.Recordset(1) = "Friday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text45 = Adodc3.Recordset(4)
        Text46 = Adodc3.Recordset(4)
        Text47 = Adodc3.Recordset(4)
        Text48 = Adodc3.Recordset(4)
        Adodc2.Recordset(45) = Adodc3.Recordset(4)
        Adodc2.Recordset(46) = Adodc3.Recordset(4)
        Adodc2.Recordset(47) = Adodc3.Recordset(4)
        Adodc2.Recordset(48) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Second Shift") And (Adodc3.Recordset(5) = "Friday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text45 = Adodc3.Recordset(8)
        Text46 = Adodc3.Recordset(8)
        Text47 = Adodc3.Recordset(8)
        Text48 = Adodc3.Recordset(8)
        Adodc2.Recordset(45) = Adodc3.Recordset(8)
        Adodc2.Recordset(46) = Adodc3.Recordset(8)
        Adodc2.Recordset(47) = Adodc3.Recordset(8)
        Adodc2.Recordset(48) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Second Shift") And (Adodc3.Recordset(9) = "Friday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text45 = Adodc3.Recordset(12)
        Text46 = Adodc3.Recordset(12)
        Text47 = Adodc3.Recordset(12)
        Text48 = Adodc3.Recordset(12)
        Adodc2.Recordset(45) = Adodc3.Recordset(12)
        Adodc2.Recordset(46) = Adodc3.Recordset(12)
        Adodc2.Recordset(47) = Adodc3.Recordset(12)
        Adodc2.Recordset(48) = Adodc3.Recordset(12)
        End If
    Adodc3.Recordset.MoveNext
    Loop
ElseIf (Adodc1.Recordset(61) = c) Then
d = Val(Adodc1.Recordset(25))
e = d + 35
Adodc2.Recordset(e) = Adodc1.Recordset(43)
If d = "1" Then
Text41.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text42.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text43.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text44.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text45.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text46.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text47.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text48.Text = Adodc2.Recordset(e)
End If
End If
End Sub

Private Sub Command2_Click()
DataReport2.Show
End Sub

Private Sub Command3_Click()
Dim a As Integer
a = MsgBox("Do you want to exit ?", vbOKCancel, "Message")
If a = 1 Then
 End
Else
 Exit Sub
End If

End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "select * from entry;"
Adodc1.Refresh
Adodc2.RecordSource = "select * from classtt;"
Adodc2.Refresh
Adodc3.RecordSource = "select * from labentry;"
Adodc3.Refresh
End Sub
Public Sub monday()
If (Adodc1.Recordset(47) = c) Then
d = Val(Adodc1.Recordset(11))
e = d + 7
Adodc2.Recordset(e) = Adodc1.Recordset(29)
If d = "1" Then
Text9.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text10.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text11.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text12.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text13.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text14.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text15.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text16.Text = Adodc2.Recordset(e)
End If
End If

If (Adodc1.Recordset(48) = c) Then
d = Val(Adodc1.Recordset(12))
e = d + 7
Adodc2.Recordset(e) = Adodc1.Recordset(30)
If d = "1" Then
Text9.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text10.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text11.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text12.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text13.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text14.Text = Adodc2.Recordset(e)
ElseIf d = "7" Then
Text15.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text16.Text = Adodc2.Recordset(e)
End If
End If


d = Val(Adodc1.Recordset(13))
If d = "0" Then
    Adodc3.Recordset.MoveFirst
    Do While Not Adodc3.Recordset.EOF
        If ((Adodc3.Recordset(2) = "Fisrt Shift") And (Adodc3.Recordset(1) = "Monday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text9 = Adodc3.Recordset(4)
        Text10 = Adodc3.Recordset(4)
        Text11 = Adodc3.Recordset(4)
        Text12 = Adodc3.Recordset(4)
        Adodc2.Recordset(9) = Adodc3.Recordset(4)
        Adodc2.Recordset(10) = Adodc3.Recordset(4)
        Adodc2.Recordset(11) = Adodc3.Recordset(4)
        Adodc2.Recordset(12) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Fisrt Shift") And (Adodc3.Recordset(5) = "Monday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text9 = Adodc3.Recordset(8)
        Text10 = Adodc3.Recordset(8)
        Text11 = Adodc3.Recordset(8)
        Text12 = Adodc3.Recordset(8)
        Adodc2.Recordset(9) = Adodc3.Recordset(8)
        Adodc2.Recordset(10) = Adodc3.Recordset(8)
        Adodc2.Recordset(11) = Adodc3.Recordset(8)
        Adodc2.Recordset(12) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Fisrt Shift") And (Adodc3.Recordset(9) = "Monday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text9 = Adodc3.Recordset(12)
        Text10 = Adodc3.Recordset(12)
        Text11 = Adodc3.Recordset(12)
        Text12 = Adodc3.Recordset(12)
        Adodc2.Recordset(9) = Adodc3.Recordset(12)
        Adodc2.Recordset(10) = Adodc3.Recordset(12)
        Adodc2.Recordset(11) = Adodc3.Recordset(12)
        Adodc2.Recordset(12) = Adodc3.Recordset(12)
        ElseIf ((Adodc3.Recordset(2) = "Second Shift") And (Adodc3.Recordset(1) = "Monday") And (Adodc3.Recordset(3) = Trim(c))) Then
        Text13 = Adodc3.Recordset(4)
        Text14 = Adodc3.Recordset(4)
        Text15 = Adodc3.Recordset(4)
        Text16 = Adodc3.Recordset(4)
        Adodc2.Recordset(13) = Adodc3.Recordset(4)
        Adodc2.Recordset(14) = Adodc3.Recordset(4)
        Adodc2.Recordset(15) = Adodc3.Recordset(4)
        Adodc2.Recordset(16) = Adodc3.Recordset(4)
        ElseIf ((Adodc3.Recordset(6) = "Second Shift") And (Adodc3.Recordset(5) = "Monday") And (Adodc3.Recordset(7) = Trim(c))) Then
        Text13 = Adodc3.Recordset(8)
        Text14 = Adodc3.Recordset(8)
        Text15 = Adodc3.Recordset(8)
        Text16 = Adodc3.Recordset(8)
        Adodc2.Recordset(13) = Adodc3.Recordset(8)
        Adodc2.Recordset(14) = Adodc3.Recordset(8)
        Adodc2.Recordset(15) = Adodc3.Recordset(8)
        Adodc2.Recordset(15) = Adodc3.Recordset(8)
        ElseIf ((Adodc3.Recordset(10) = "Second Shift") And (Adodc3.Recordset(9) = "Monday") And (Adodc3.Recordset(11) = Trim(c))) Then
        Text13 = Adodc3.Recordset(12)
        Text14 = Adodc3.Recordset(12)
        Text15 = Adodc3.Recordset(12)
        Text16 = Adodc3.Recordset(12)
        Adodc2.Recordset(13) = Adodc3.Recordset(12)
        Adodc2.Recordset(14) = Adodc3.Recordset(12)
        Adodc2.Recordset(15) = Adodc3.Recordset(12)
        Adodc2.Recordset(15) = Adodc3.Recordset(12)
        End If
    Adodc3.Recordset.MoveNext
    Loop
ElseIf (Adodc1.Recordset(49) = c) Then
d = Val(Adodc1.Recordset(13))
e = d + 7
Adodc2.Recordset(e) = Adodc1.Recordset(31)
If d = "1" Then
Text9.Text = Adodc2.Recordset(e)
ElseIf d = "2" Then
Text10.Text = Adodc2.Recordset(e)
ElseIf d = "3" Then
Text11.Text = Adodc2.Recordset(e)
ElseIf d = "4" Then
Text12.Text = Adodc2.Recordset(e)
ElseIf d = "5" Then
Text13.Text = Adodc2.Recordset(e)
ElseIf d = "6" Then
Text14.Text = Adodc2.Recordset(d)
ElseIf d = "7" Then
Text15.Text = Adodc2.Recordset(e)
ElseIf d = "8" Then
Text16.Text = Adodc2.Recordset(e)
End If
End If
End Sub
