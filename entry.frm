VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00008000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Timetable Management System - Entry Form"
   ClientHeight    =   10185
   ClientLeft      =   330
   ClientTop       =   1035
   ClientWidth     =   14670
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Show Routine"
      Height          =   495
      Left            =   240
      TabIndex        =   131
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Goto Home"
      Height          =   495
      Left            =   240
      TabIndex        =   130
      Top             =   4590
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Height          =   495
      Left            =   240
      TabIndex        =   129
      Top             =   3900
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   240
      TabIndex        =   128
      Top             =   3210
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   240
      TabIndex        =   127
      Top             =   2520
      Width           =   2175
   End
   Begin VB.ComboBox Combo32 
      Height          =   315
      Left            =   13320
      TabIndex        =   126
      Top             =   7560
      Width           =   975
   End
   Begin VB.ComboBox Combo31 
      Height          =   315
      Left            =   13320
      TabIndex        =   125
      Top             =   7080
      Width           =   975
   End
   Begin VB.ComboBox Combo30 
      Height          =   315
      Left            =   13320
      TabIndex        =   124
      Top             =   6480
      Width           =   975
   End
   Begin VB.ComboBox Combo22 
      Height          =   315
      Left            =   13320
      TabIndex        =   123
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox Combo15 
      Height          =   315
      Left            =   10080
      TabIndex        =   122
      Top             =   8040
      Width           =   1695
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   10080
      TabIndex        =   121
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text19 
      Height          =   315
      Left            =   1440
      TabIndex        =   120
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text18 
      Height          =   315
      Left            =   7320
      TabIndex        =   119
      Top             =   9600
      Width           =   855
   End
   Begin VB.TextBox Text17 
      Height          =   315
      Left            =   7320
      TabIndex        =   118
      Top             =   9120
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   10080
      TabIndex        =   117
      Top             =   1320
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   735
      Left            =   2040
      Top             =   10080
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
      Height          =   855
      Left            =   -120
      Top             =   10080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
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
      Height          =   975
      Left            =   4200
      Top             =   10080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1720
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
   Begin VB.TextBox Text16 
      Height          =   315
      Left            =   7320
      TabIndex        =   114
      Text            =   " "
      Top             =   8640
      Width           =   855
   End
   Begin VB.ComboBox Combo0 
      Height          =   315
      Left            =   1440
      TabIndex        =   113
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      Height          =   315
      Left            =   7320
      TabIndex        =   100
      Top             =   8040
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Height          =   315
      Left            =   7320
      TabIndex        =   99
      Top             =   7560
      Width           =   855
   End
   Begin VB.ComboBox Combo36 
      Height          =   315
      Left            =   13320
      TabIndex        =   98
      Top             =   9600
      Width           =   975
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7320
      TabIndex        =   97
      Top             =   7080
      Width           =   855
   End
   Begin VB.ComboBox Combo35 
      Height          =   315
      Left            =   13320
      TabIndex        =   96
      Top             =   9120
      Width           =   975
   End
   Begin VB.ComboBox Combo34 
      Height          =   315
      Left            =   13320
      TabIndex        =   95
      Top             =   8640
      Width           =   975
   End
   Begin VB.ComboBox Combo33 
      Height          =   315
      Left            =   13320
      TabIndex        =   94
      Top             =   8040
      Width           =   975
   End
   Begin VB.ComboBox comDay6 
      Height          =   315
      Left            =   4680
      TabIndex        =   92
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      Height          =   315
      Left            =   7320
      TabIndex        =   82
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   315
      Left            =   7320
      TabIndex        =   81
      Top             =   6000
      Width           =   855
   End
   Begin VB.ComboBox Combo28 
      Height          =   315
      Left            =   13320
      TabIndex        =   80
      Top             =   5520
      Width           =   975
   End
   Begin VB.ComboBox Combo27 
      Height          =   315
      Left            =   13320
      TabIndex        =   79
      Top             =   4920
      Width           =   975
   End
   Begin VB.ComboBox Combo26 
      Height          =   315
      Left            =   13320
      TabIndex        =   78
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7320
      TabIndex        =   77
      Top             =   5520
      Width           =   855
   End
   Begin VB.ComboBox Combo25 
      Height          =   315
      Left            =   13320
      TabIndex        =   76
      Top             =   3960
      Width           =   975
   End
   Begin VB.ComboBox Combo24 
      Height          =   315
      Left            =   13320
      TabIndex        =   75
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox Combo23 
      Height          =   315
      Left            =   13320
      TabIndex        =   74
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox comDay5 
      Height          =   315
      Left            =   4680
      TabIndex        =   72
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Height          =   315
      Left            =   7320
      TabIndex        =   62
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   315
      Left            =   7320
      TabIndex        =   61
      Top             =   4440
      Width           =   855
   End
   Begin VB.ComboBox Combo21 
      Height          =   315
      Left            =   13320
      TabIndex        =   60
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox Combo20 
      Height          =   315
      Left            =   13320
      TabIndex        =   59
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox Combo19 
      Height          =   315
      Left            =   13320
      TabIndex        =   58
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7320
      TabIndex        =   57
      Top             =   3960
      Width           =   855
   End
   Begin VB.ComboBox Combo18 
      Height          =   315
      Left            =   10080
      TabIndex        =   56
      Top             =   9600
      Width           =   1695
   End
   Begin VB.ComboBox Combo17 
      Height          =   315
      Left            =   10080
      TabIndex        =   55
      Top             =   9120
      Width           =   1695
   End
   Begin VB.ComboBox Combo16 
      Height          =   315
      Left            =   10080
      TabIndex        =   54
      Top             =   8640
      Width           =   1695
   End
   Begin VB.ComboBox comDay4 
      Height          =   315
      Left            =   4680
      TabIndex        =   52
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   7320
      TabIndex        =   42
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   7320
      TabIndex        =   41
      Top             =   2880
      Width           =   855
   End
   Begin VB.ComboBox Combo14 
      Height          =   315
      Left            =   10080
      TabIndex        =   40
      Top             =   7560
      Width           =   1695
   End
   Begin VB.ComboBox Combo13 
      Height          =   315
      Left            =   10080
      TabIndex        =   39
      Top             =   7080
      Width           =   1695
   End
   Begin VB.ComboBox Combo12 
      Height          =   315
      Left            =   10080
      TabIndex        =   38
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7320
      TabIndex        =   37
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   10080
      TabIndex        =   36
      Top             =   6000
      Width           =   1695
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      Left            =   10080
      TabIndex        =   35
      Top             =   5520
      Width           =   1695
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      ItemData        =   "entry.frx":0000
      Left            =   10080
      List            =   "entry.frx":0002
      TabIndex        =   34
      Top             =   4920
      Width           =   1695
   End
   Begin VB.ComboBox comDay3 
      Height          =   315
      Left            =   4680
      TabIndex        =   32
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   7320
      TabIndex        =   22
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   7320
      TabIndex        =   21
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   10080
      TabIndex        =   20
      Top             =   3960
      Width           =   1695
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   10080
      TabIndex        =   19
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   10080
      TabIndex        =   18
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7320
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   10080
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox comDay2 
      Height          =   315
      Left            =   4680
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox Combo29 
      Height          =   315
      Left            =   13320
      TabIndex        =   3
      Top             =   6000
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   10080
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox comDay1 
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label65 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   " Faculty Id"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   116
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label64 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   " Staff Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   115
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Type the hour between 1 to 8"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5400
      TabIndex        =   112
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label Label62 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   " Select the Subject"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   9240
      TabIndex        =   111
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   " Select the class"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   12480
      TabIndex        =   110
      Top             =   240
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   6
      X1              =   3840
      X2              =   14520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   5
      X1              =   3840
      X2              =   14520
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   4
      X1              =   3960
      X2              =   14640
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   3
      X1              =   3960
      X2              =   14640
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   2
      X1              =   3960
      X2              =   14640
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   1
      X1              =   3960
      X2              =   14640
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   0
      X1              =   3840
      X2              =   14520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label60 
      Caption         =   "Hour 1"
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
      Left            =   6480
      TabIndex        =   109
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label59 
      Caption         =   "Hour 2"
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
      Left            =   6480
      TabIndex        =   108
      Top             =   9120
      Width           =   615
   End
   Begin VB.Label Label58 
      Caption         =   "Hour 3"
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
      Left            =   6480
      TabIndex        =   107
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label57 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   106
      Top             =   9120
      Width           =   615
   End
   Begin VB.Label Label56 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   105
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label55 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   104
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label54 
      Caption         =   "Subject 1"
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
      Left            =   8760
      TabIndex        =   103
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label Label53 
      Caption         =   "Subject 2"
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
      Left            =   8760
      TabIndex        =   102
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label Label52 
      Caption         =   "Subject 3/Lab"
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
      Left            =   8760
      TabIndex        =   101
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Label Label51 
      Caption         =   "Day6"
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
      Left            =   3960
      TabIndex        =   93
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label Label50 
      Caption         =   "Hour 1"
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
      Left            =   6480
      TabIndex        =   91
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label49 
      Caption         =   "Hour 2"
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
      Left            =   6480
      TabIndex        =   90
      Top             =   7560
      Width           =   615
   End
   Begin VB.Label Label48 
      Caption         =   "Hour 3"
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
      Left            =   6480
      TabIndex        =   89
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label47 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   88
      Top             =   7560
      Width           =   615
   End
   Begin VB.Label Label44 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   87
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label42 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   86
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label41 
      Caption         =   "Subject 1"
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
      Left            =   8760
      TabIndex        =   85
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label40 
      Caption         =   "Subject 2"
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
      Left            =   8760
      TabIndex        =   84
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label39 
      Caption         =   "Subject 3/Lab"
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
      Left            =   8760
      TabIndex        =   83
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label38 
      Caption         =   "Day5"
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
      Left            =   3960
      TabIndex        =   73
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label Label37 
      Caption         =   "Hour 1"
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
      Left            =   6480
      TabIndex        =   71
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label36 
      Caption         =   "Hour 2"
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
      Left            =   6480
      TabIndex        =   70
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label35 
      Caption         =   "Hour 3"
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
      Left            =   6480
      TabIndex        =   69
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label34 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   68
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label33 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   67
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label32 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   66
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label31 
      Caption         =   "Subject 1"
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
      Left            =   8760
      TabIndex        =   65
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label30 
      Caption         =   "Subject 2"
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
      Left            =   8760
      TabIndex        =   64
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label29 
      Caption         =   "Subject 3/Lab"
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
      Left            =   8760
      TabIndex        =   63
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label28 
      Caption         =   "Day4"
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
      Left            =   3960
      TabIndex        =   53
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label27 
      Caption         =   "Hour 1"
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
      Left            =   6480
      TabIndex        =   51
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label26 
      Caption         =   "Hour 2"
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
      Left            =   6480
      TabIndex        =   50
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label25 
      Caption         =   "Hour 3"
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
      Left            =   6480
      TabIndex        =   49
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label24 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   48
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label23 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   47
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label22 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   46
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label21 
      Caption         =   "Subject 1"
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
      Left            =   8760
      TabIndex        =   45
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label20 
      Caption         =   "Subject 2"
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
      Left            =   8760
      TabIndex        =   44
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   "Subject 3/Lab"
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
      Left            =   8760
      TabIndex        =   43
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "Day3"
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
      Left            =   3960
      TabIndex        =   33
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Hour 1"
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
      Left            =   6480
      TabIndex        =   31
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Hour 2"
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
      Left            =   6480
      TabIndex        =   30
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "Hour 3"
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
      Left            =   6480
      TabIndex        =   29
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   28
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   27
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   26
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Subject 1"
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
      Left            =   8760
      TabIndex        =   25
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Subject 2"
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
      Left            =   8760
      TabIndex        =   24
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Subject 3/Lab"
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
      Left            =   8760
      TabIndex        =   23
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Day2"
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
      Left            =   3960
      TabIndex        =   14
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Hour 1"
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
      Left            =   6480
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Hour 2"
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
      Left            =   6480
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Hour 3"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label46 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label45 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label43 
      Caption         =   "Class"
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
      Left            =   12480
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Subject 1"
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
      Left            =   8760
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Subject 2"
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
      Left            =   8760
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Subject 3/Lab"
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
      Left            =   8760
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub clear()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Text12 = ""
Text13 = ""
Text14 = ""
Text15 = ""
Text16 = ""
Text17 = ""
Text18 = ""
Text19 = ""
End Sub



Private Sub Command1_Click()
Dim a As Integer

On Error GoTo soln
clear
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
Combo0.AddItem Adodc2.Recordset(1)
If Combo0.Text = Adodc2.Recordset(1) Then
Text19.Text = Adodc2.Recordset(0)
End If
Adodc2.Recordset.MoveNext
Loop
init
Adodc1.Recordset.AddNew
Exit Sub
soln:
a = MsgBox("Empty Row cannot be inserted", vbOKOnly, "Error Message")
End Sub
Private Sub Command2_Click()
Dim a As Integer
On Error GoTo soln
assign
Adodc1.Recordset.Update
MsgBox "Record updated"
Exit Sub

soln:
a = MsgBox("Hours already assigned, Please Try next", vbOKCancel, "Error Message")
clear
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

Private Sub Command4_Click()
Unload Me
Form2.Show
End Sub

Private Sub Command5_Click()
Unload Me
Form7.Show

End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "select * from entry;"
Adodc1.Refresh
Adodc2.RecordSource = "select * from staffentry;"
Adodc2.Refresh
Adodc3.RecordSource = "select * from labentry;"
Adodc3.Refresh

End Sub
Public Sub assign()
Adodc1.Recordset(0) = Val(Text19)
Adodc1.Recordset(1) = Combo0.Text
Adodc1.Recordset(2) = comDay1.Text
Adodc1.Recordset(3) = comDay2.Text
Adodc1.Recordset(4) = comDay3.Text
Adodc1.Recordset(5) = comDay4.Text
Adodc1.Recordset(6) = comDay5.Text
Adodc1.Recordset(7) = comDay6.Text

Adodc1.Recordset(8) = Val(Text1)
Adodc1.Recordset(9) = Val(Text2)
Adodc1.Recordset(10) = Val(Text3)
Adodc1.Recordset(11) = Val(Text4)
Adodc1.Recordset(12) = Val(Text5)
Adodc1.Recordset(13) = Val(Text6)
Adodc1.Recordset(14) = Val(Text7)
Adodc1.Recordset(15) = Val(Text8)
Adodc1.Recordset(16) = Val(Text9)
Adodc1.Recordset(17) = Val(Text10)
Adodc1.Recordset(18) = Val(Text11)
Adodc1.Recordset(19) = Val(Text12)
Adodc1.Recordset(20) = Val(Text13)
Adodc1.Recordset(21) = Val(Text14)
Adodc1.Recordset(22) = Val(Text15)
Adodc1.Recordset(23) = Val(Text16)
Adodc1.Recordset(24) = Val(Text17)
Adodc1.Recordset(25) = Val(Text18)

Adodc1.Recordset(26) = Combo1.Text
Adodc1.Recordset(27) = Combo2.Text
Adodc1.Recordset(28) = Combo3.Text
Adodc1.Recordset(29) = Combo4.Text
Adodc1.Recordset(30) = Combo5.Text
Adodc1.Recordset(31) = Combo6.Text
Adodc1.Recordset(32) = Combo7.Text
Adodc1.Recordset(33) = Combo8.Text
Adodc1.Recordset(34) = Combo9.Text
Adodc1.Recordset(35) = Combo10.Text
Adodc1.Recordset(36) = Combo11.Text
Adodc1.Recordset(37) = Combo12.Text
Adodc1.Recordset(38) = Combo13.Text
Adodc1.Recordset(39) = Combo14.Text
Adodc1.Recordset(40) = Combo15.Text

Adodc1.Recordset(41) = Combo16.Text
Adodc1.Recordset(42) = Combo17.Text
Adodc1.Recordset(43) = Combo18.Text
Adodc1.Recordset(44) = Combo19.Text
Adodc1.Recordset(45) = Combo20.Text
Adodc1.Recordset(46) = Combo21.Text
Adodc1.Recordset(47) = Combo22.Text
Adodc1.Recordset(48) = Combo23.Text
Adodc1.Recordset(49) = Combo24.Text
Adodc1.Recordset(50) = Combo25.Text
Adodc1.Recordset(51) = Combo26.Text
Adodc1.Recordset(52) = Combo27.Text
Adodc1.Recordset(53) = Combo28.Text
Adodc1.Recordset(54) = Combo29.Text
Adodc1.Recordset(55) = Combo30.Text

Adodc1.Recordset(56) = Combo31.Text
Adodc1.Recordset(57) = Combo32.Text
Adodc1.Recordset(58) = Combo33.Text
Adodc1.Recordset(59) = Combo34.Text
Adodc1.Recordset(60) = Combo35.Text
Adodc1.Recordset(61) = Combo36.Text
End Sub
Public Sub init()
comDay1.AddItem "Sunday"
comDay1.AddItem "Monday"
comDay1.AddItem "Tuesday"
comDay1.AddItem "Wednesday"
comDay1.AddItem "Thursday"
comDay1.AddItem "Friday"

comDay2.AddItem "Sunday"
comDay2.AddItem "Monday"
comDay2.AddItem "Tuesday"
comDay2.AddItem "Wednesday"
comDay2.AddItem "Thursday"
comDay2.AddItem "Friday"

comDay3.AddItem "Sunday"
comDay3.AddItem "Monday"
comDay3.AddItem "Tuesday"
comDay3.AddItem "Wednesday"
comDay3.AddItem "Thursday"
comDay3.AddItem "Friday"

comDay4.AddItem "Sunday"
comDay4.AddItem "Monday"
comDay4.AddItem "Tuesday"
comDay4.AddItem "Wednesday"
comDay4.AddItem "Thursday"
comDay4.AddItem "Friday"

comDay5.AddItem "Sunday"
comDay5.AddItem "Monday"
comDay5.AddItem "Tuesday"
comDay5.AddItem "Wednesday"
comDay5.AddItem "Thursday"
comDay5.AddItem "Friday"

comDay6.AddItem "Sunday"
comDay6.AddItem "Monday"
comDay6.AddItem "Tuesday"
comDay6.AddItem "Wednesday"
comDay6.AddItem "Thursday"
comDay6.AddItem "Friday"

' adding class number to combo 19 - 36
Combo19.AddItem "101"
Combo19.AddItem "102"
Combo19.AddItem "103"

Combo20.AddItem "101"
Combo20.AddItem "102"
Combo20.AddItem "103"

Combo21.AddItem "101"
Combo21.AddItem "102"
Combo21.AddItem "103"

Combo22.AddItem "101"
Combo22.AddItem "102"
Combo22.AddItem "103"

Combo23.AddItem "101"
Combo23.AddItem "102"
Combo23.AddItem "103"

Combo24.AddItem "101"
Combo24.AddItem "102"
Combo24.AddItem "103"

Combo25.AddItem "101"
Combo25.AddItem "102"
Combo25.AddItem "103"

Combo26.AddItem "101"
Combo26.AddItem "102"
Combo26.AddItem "103"

Combo27.AddItem "101"
Combo27.AddItem "102"
Combo27.AddItem "103"

Combo28.AddItem "101"
Combo28.AddItem "102"
Combo28.AddItem "103"

Combo29.AddItem "101"
Combo29.AddItem "102"
Combo29.AddItem "103"

Combo30.AddItem "101"
Combo30.AddItem "102"
Combo30.AddItem "103"

Combo31.AddItem "101"
Combo31.AddItem "102"
Combo31.AddItem "103"

Combo32.AddItem "101"
Combo32.AddItem "102"
Combo32.AddItem "103"

Combo33.AddItem "101"
Combo33.AddItem "102"
Combo33.AddItem "103"

Combo34.AddItem "101"
Combo34.AddItem "102"
Combo34.AddItem "103"

Combo35.AddItem "101"
Combo35.AddItem "102"
Combo35.AddItem "103"

Combo36.AddItem "101"
Combo36.AddItem "102"
Combo36.AddItem "103"


End Sub


Private Sub comDay1_GotFocus()
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
If Combo0.Text = Adodc2.Recordset(1) Then
Text19.Text = Adodc2.Recordset(0)
If Adodc2.Recordset(3) = "HOD" Then
Combo5.Enabled = False
Combo6.Enabled = False
Combo9.Enabled = False
Combo12.Enabled = False
Combo15.Enabled = False
Combo18.Enabled = False
End If
End If
Adodc2.Recordset.MoveNext
Loop
init1
If ((Combo0.Text = Adodc3.Recordset(0)) And ((comDay1.Text = Adodc3.Recordset(1)) Or (comDay1.Text = Adodc3.Recordset(5)) Or (comDay1.Text = Adodc3.Recordset(9)))) Then
Text3.Enabled = False
Combo3.Enabled = False '
Adodc1.Recordset(8) = "0"
End If
If ((Combo0.Text = Adodc3.Recordset(0)) And ((comDay2.Text = Adodc3.Recordset(1)) Or (comDay2.Text = Adodc3.Recordset(5)) Or (comDay2.Text = Adodc3.Recordset(9)))) Then
Text6.Enabled = False
Combo6.Enabled = False '
Adodc1.Recordset(11) = "0"
End If
If ((Combo0.Text = Adodc3.Recordset(0)) And ((comDay3.Text = Adodc3.Recordset(1)) Or (comDay3.Text = Adodc3.Recordset(5)) Or (comDay3.Text = Adodc3.Recordset(9)))) Then
Text9.Enabled = False
Combo9.Enabled = False '
Adodc1.Recordset(14) = "0"
End If
If ((Combo0.Text = Adodc3.Recordset(0)) And ((comDay4.Text = Adodc3.Recordset(1)) Or (comDay4.Text = Adodc3.Recordset(5)) Or (comDay4.Text = Adodc3.Recordset(9)))) Then
Text12.Enabled = False
Combo12.Enabled = False '
Adodc1.Recordset(17) = "0"
End If
If ((Combo0.Text = Adodc3.Recordset(0)) And ((comDay5.Text = Adodc3.Recordset(1)) Or (comDay5.Text = Adodc3.Recordset(5)) Or (comDay5.Text = Adodc3.Recordset(9)))) Then
Text15.Enabled = False
Combo15.Enabled = False '
Adodc1.Recordset(20) = "0"
End If
If ((Combo0.Text = Adodc3.Recordset(0)) And ((comDay6.Text = Adodc3.Recordset(1)) Or (comDay6.Text = Adodc3.Recordset(5)) Or (comDay6.Text = Adodc3.Recordset(9)))) Then
Text18.Enabled = False
Combo18.Enabled = False '
Adodc1.Recordset(23) = "0"
End If
End Sub

Private Sub Combo19_LostFocus()
If ((Text1 = Adodc1.Recordset(8)) And (Combo19.Text = Adodc1.Recordset(44))) Then
MsgBox "please select some other hour! already alloted"
Text1 = "" '
End If

End Sub
Private Sub Combo20_LostFocus()
If ((Text2 = Adodc1.Recordset(9)) And (Combo20.Text = Adodc1.Recordset(45))) Then
MsgBox "please select some other hour! already alloted"
Text2 = "" '
End If

End Sub
Private Sub Combo21_LostFocus()
If ((Text3 = Adodc1.Recordset(10)) And (Combo21.Text = Adodc1.Recordset(46))) Then
MsgBox "please select some other hour! already alloted"
Text3 = "" '
End If
End Sub
Private Sub Combo22_LostFocus()
If ((Text4 = Adodc1.Recordset(11)) And (Combo22.Text = Adodc1.Recordset(47))) Then
MsgBox "please select some other hour! already alloted"
Text4 = ""
End If
End Sub
Private Sub Combo23_LostFocus()
If ((Text5 = Adodc1.Recordset(12)) And (Combo23.Text = Adodc1.Recordset(48))) Then
MsgBox "please select some other hour! already alloted"
Text5 = ""
End If
End Sub
Private Sub Combo24_LostFocus()
If ((Text6 = Adodc1.Recordset(13)) And (Combo24.Text = Adodc1.Recordset(49))) Then
MsgBox "please select some other hour! already alloted"
Text6 = ""
End If

End Sub
Private Sub Combo25_LostFocus()
If ((Text7 = Adodc1.Recordset(14)) And (Combo25.Text = Adodc1.Recordset(50))) Then
MsgBox "please select some other hour! already alloted"
Text7 = ""
End If
End Sub
Private Sub Combo26_LostFocus()
If ((Text8 = Adodc1.Recordset(15)) And (Combo26.Text = Adodc1.Recordset(51))) Then
MsgBox "please select some other hour! already alloted"
Text8 = ""
End If
End Sub
Private Sub Combo27_LostFocus()
If ((Text9 = Adodc1.Recordset(16)) And (Combo27.Text = Adodc1.Recordset(52))) Then
MsgBox "please select some other hour! already alloted"
Text9 = ""
End If
End Sub
Private Sub Combo28_LostFocus()
If ((Text10 = Adodc1.Recordset(17)) And (Combo28.Text = Adodc1.Recordset(53))) Then
MsgBox "please select some other hour! already alloted"
Text10 = ""
End If
End Sub
Private Sub Combo29_LostFocus()
If ((Text11 = Adodc1.Recordset(18)) And (Combo29.Text = Adodc1.Recordset(54))) Then
MsgBox "please select some other hour! already alloted"
Text11 = ""
End If
End Sub
Private Sub Combo30_LostFocus()
If ((Text12 = Adodc1.Recordset(19)) And (Combo30.Text = Adodc1.Recordset(55))) Then
MsgBox "please select some other hour! already alloted"
Text12 = ""
End If
End Sub
Private Sub Combo31_LostFocus()
If ((Text13 = Adodc1.Recordset(20)) And (Combo31.Text = Adodc1.Recordset(56))) Then
MsgBox "please select some other hour! already alloted"
Text13 = ""
End If
End Sub
Private Sub Combo32_LostFocus()
If ((Text14 = Adodc1.Recordset(21)) And (Combo32.Text = Adodc1.Recordset(57))) Then
MsgBox "please select some other hour! already alloted"
Text14 = ""
End If
End Sub
Private Sub Combo33_LostFocus()
If ((Text15 = Adodc1.Recordset(22)) And (Combo33.Text = Adodc1.Recordset(58))) Then
MsgBox "please select some other hour! already alloted"
Text15 = "" '
End If
End Sub
Private Sub Combo34_LostFocus()
If ((Text16 = Adodc1.Recordset(23)) And (Combo34.Text = Adodc1.Recordset(59))) Then
MsgBox "please select some other hour! already alloted"
Text16 = "" '
End If
End Sub
Private Sub Combo35_LostFocus()
If ((Text17 = Adodc1.Recordset(24)) And (Combo35.Text = Adodc1.Recordset(60))) Then
MsgBox "please select some other hour! already alloted"
Text17 = "" '
End If
End Sub
Private Sub Combo36_LostFocus()
If ((Text18 = Adodc1.Recordset(25)) And (Combo36.Text = Adodc1.Recordset(61))) Then
MsgBox "please select some other hour! already alloted"
Text18 = "" '
End If
End Sub


Private Sub init1()
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
If Val(Text19) = Adodc2.Recordset(0) Then
Combo1.AddItem Adodc2.Recordset(2)
Combo1.AddItem Adodc2.Recordset(3)
Combo1.AddItem Adodc2.Recordset(4)
'Combo1.AddItem Adodc2.Recordset(5)

Combo2.AddItem Adodc2.Recordset(2)
Combo2.AddItem Adodc2.Recordset(3)
Combo2.AddItem Adodc2.Recordset(4)
'Combo2.AddItem Adodc2.Recordset(5)

Combo3.AddItem Adodc2.Recordset(2)
Combo3.AddItem Adodc2.Recordset(3)
Combo3.AddItem Adodc2.Recordset(4)
'Combo3.AddItem Adodc2.Recordset(5)

Combo4.AddItem Adodc2.Recordset(2)
Combo4.AddItem Adodc2.Recordset(3)
Combo4.AddItem Adodc2.Recordset(4)
'Combo4.AddItem Adodc2.Recordset(5)

Combo5.AddItem Adodc2.Recordset(2)
Combo5.AddItem Adodc2.Recordset(3)
Combo5.AddItem Adodc2.Recordset(4)
'Combo5.AddItem Adodc2.Recordset(5)

Combo6.AddItem Adodc2.Recordset(2)
Combo6.AddItem Adodc2.Recordset(3)
Combo6.AddItem Adodc2.Recordset(4)
'Combo6.AddItem Adodc2.Recordset(5)

Combo7.AddItem Adodc2.Recordset(2)
Combo7.AddItem Adodc2.Recordset(3)
Combo7.AddItem Adodc2.Recordset(4)
'Combo7.AddItem Adodc2.Recordset(5)

Combo8.AddItem Adodc2.Recordset(2)
Combo8.AddItem Adodc2.Recordset(3)
Combo8.AddItem Adodc2.Recordset(4)
'Combo8.AddItem Adodc2.Recordset(5)

Combo9.AddItem Adodc2.Recordset(2)
Combo9.AddItem Adodc2.Recordset(3)
Combo9.AddItem Adodc2.Recordset(4)
'Combo9.AddItem Adodc2.Recordset(5)

Combo10.AddItem Adodc2.Recordset(2)
Combo10.AddItem Adodc2.Recordset(3)
Combo10.AddItem Adodc2.Recordset(4)
'Combo10.AddItem Adodc2.Recordset(5)

Combo11.AddItem Adodc2.Recordset(2)
Combo11.AddItem Adodc2.Recordset(3)
Combo11.AddItem Adodc2.Recordset(4)
'Combo11.AddItem Adodc2.Recordset(5)

Combo12.AddItem Adodc2.Recordset(2)
Combo12.AddItem Adodc2.Recordset(3)
Combo12.AddItem Adodc2.Recordset(4)
'Combo12.AddItem Adodc2.Recordset(5)

Combo13.AddItem Adodc2.Recordset(2)
Combo13.AddItem Adodc2.Recordset(3)
Combo13.AddItem Adodc2.Recordset(4)
'Combo13.AddItem Adodc2.Recordset(5)

Combo14.AddItem Adodc2.Recordset(2)
Combo14.AddItem Adodc2.Recordset(3)
Combo14.AddItem Adodc2.Recordset(4)
'Combo14.AddItem Adodc2.Recordset(5)

Combo15.AddItem Adodc2.Recordset(2)
Combo15.AddItem Adodc2.Recordset(3)
Combo15.AddItem Adodc2.Recordset(4)
'Combo15.AddItem Adodc2.Recordset(5)

Combo16.AddItem Adodc2.Recordset(2)
Combo16.AddItem Adodc2.Recordset(3)
Combo16.AddItem Adodc2.Recordset(4)
'Combo16.AddItem Adodc2.Recordset(5)

Combo17.AddItem Adodc2.Recordset(2)
Combo17.AddItem Adodc2.Recordset(3)
Combo17.AddItem Adodc2.Recordset(4)
'Combo17.AddItem Adodc2.Recordset(5)

Combo18.AddItem Adodc2.Recordset(2)
Combo18.AddItem Adodc2.Recordset(3)
Combo18.AddItem Adodc2.Recordset(4)
'Combo18.AddItem Adodc2.Recordset(5)

End If
Adodc2.Recordset.MoveNext
Loop
End Sub

