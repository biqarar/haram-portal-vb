VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form QeybatF 
   Caption         =   "”Ì” „  „œÌ—Ì  Õ÷Ê— Ê €Ì«» ﬁ—¬‰ ¬„Ê“«‰"
   ClientHeight    =   9375
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   18240
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "QeybatF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   18240
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo15 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatF.frx":08CA
      Left            =   9840
      List            =   "QeybatF.frx":08CC
      TabIndex        =   160
      Text            =   "1391"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   300
      Left            =   13200
      TabIndex        =   157
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   18120
      TabIndex        =   156
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Ê÷ÌÕ« "
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000003&
      ForeColor       =   &H80000007&
      Height          =   360
      ItemData        =   "QeybatF.frx":08CE
      Left            =   120
      List            =   "QeybatF.frx":08D0
      TabIndex        =   154
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Frame Frame9 
      Caption         =   "Motor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   7440
      TabIndex        =   122
      Top             =   7560
      Visible         =   0   'False
      Width           =   5775
      Begin MSAdodcLib.Adodc vadie 
         Height          =   330
         Left            =   2640
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":08D2
         OLEDBString     =   $"QeybatF.frx":095B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from vadie"
         Caption         =   "vadie"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Qeybat 
         Height          =   330
         Left            =   2640
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":09E4
         OLEDBString     =   $"QeybatF.frx":0A6D
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from qeybat"
         Caption         =   "Qeybat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc STU2CLASS 
         Height          =   375
         Left            =   360
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":0AF6
         OLEDBString     =   $"QeybatF.frx":0B7F
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from stu2class"
         Caption         =   "STU2CLASS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc mclass 
         Height          =   375
         Left            =   360
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":0C08
         OLEDBString     =   $"QeybatF.frx":0C91
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mclass"
         Caption         =   "mclass"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc teacher 
         Height          =   375
         Left            =   360
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":0D1A
         OLEDBString     =   $"QeybatF.frx":0DA3
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from teacher"
         Caption         =   "teacher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Student 
         Height          =   330
         Left            =   360
         Top             =   360
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":0E2C
         OLEDBString     =   $"QeybatF.frx":0EB5
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from student"
         Caption         =   "Student"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Tarhha 
         Height          =   375
         Left            =   360
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":0F3E
         OLEDBString     =   $"QeybatF.frx":0FC7
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tarhha"
         Caption         =   "Tarhha"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc ekhtar 
         Height          =   330
         Left            =   2640
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":1050
         OLEDBString     =   $"QeybatF.frx":10D9
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from ekhtar"
         Caption         =   "ekhtar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Setting 
         Height          =   330
         Left            =   2640
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":1162
         OLEDBString     =   $"QeybatF.frx":11EB
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from settingtable"
         Caption         =   "Setting"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc userprofiletable 
         Height          =   330
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":1274
         OLEDBString     =   $"QeybatF.frx":12FD
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from userprofiletable"
         Caption         =   "userprofiletable"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc tozih_table 
         Height          =   330
         Left            =   2640
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"QeybatF.frx":1386
         OLEDBString     =   $"QeybatF.frx":140F
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tozih_table"
         Caption         =   "tozih_table"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.ComboBox Combo14 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatF.frx":1498
      Left            =   9360
      List            =   "QeybatF.frx":149A
      TabIndex        =   149
      Text            =   "«‰ Œ«» ò‰Ìœ"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H000000FF&
      Caption         =   "Print Taahod"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Height          =   300
      Left            =   7200
      TabIndex        =   147
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command20 
      Caption         =   " «ÌÌœ"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   146
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Combo13 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatF.frx":149C
      Left            =   2640
      List            =   "QeybatF.frx":149E
      TabIndex        =   143
      Text            =   "01"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Combo12 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1920
      TabIndex        =   142
      Text            =   "01"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Combo11 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatF.frx":14A0
      Left            =   1080
      List            =   "QeybatF.frx":14A2
      TabIndex        =   141
      Text            =   "1390"
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo10 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatF.frx":14A4
      Left            =   5640
      List            =   "QeybatF.frx":14A6
      TabIndex        =   140
      Text            =   "01"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Combo9 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4920
      TabIndex        =   139
      Text            =   "01"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Combo8 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatF.frx":14A8
      Left            =   4080
      List            =   "QeybatF.frx":14AA
      TabIndex        =   138
      Text            =   "1390"
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H8000000A&
      Caption         =   " ‰ŸÌ„« "
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Combo7 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatF.frx":14AC
      Left            =   10680
      List            =   "QeybatF.frx":14AE
      TabIndex        =   135
      Text            =   " „«„Ì „«Â Â«"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12600
      Top             =   3960
   End
   Begin ComctlLib.ProgressBar Pb2 
      Height          =   135
      Left            =   9360
      TabIndex        =   128
      Top             =   3600
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "QeybatF.frx":14B0
      Height          =   4455
      Left            =   120
      TabIndex        =   127
      ToolTipText     =   "»—«Ì „‘«ÂœÂ «”ò‰ ›«Ì· »— —ÊÌ Å—Ê‰œÂ œÊ »«— ò·Ìò ò‰Ìœ"
      Top             =   4560
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
      DefColWidth     =   120
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«ÿ·«⁄«  ﬁ—¬‰ ¬„Ê“«‰"
      ColumnCount     =   27
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "‰«„"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Famil"
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Namepedar"
         Caption         =   "‰«„ Åœ—"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Tavalod"
         Caption         =   " «—ÌŒ  Ê·œ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Shsh"
         Caption         =   "‘„«—Â ‘‰«”‰«„Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Sadere"
         Caption         =   "’«œ—Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Meliyat"
         Caption         =   "„·Ì "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Mazhab"
         Caption         =   "„–Â»"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Kodmeli"
         Caption         =   "òœ „·Ì"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Gozarname"
         Caption         =   "‘„«—Â ê–“ ‰«„Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Taahol"
         Caption         =   "Ê÷⁄Ì   «Â·"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Farzand"
         Caption         =   " ⁄œ«œ ›—“‰œ«‰"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Tahsilat"
         Caption         =   " Õ’Ì·« "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Ostad"
         Caption         =   "‰«„ «” «œ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Tozih"
         Caption         =   " Ê÷ÌÕ« "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "Tell"
         Caption         =   " ·›‰ À«» "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "Mob"
         Caption         =   "Â„—«Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "Scan"
         Caption         =   "«”ò‰ ›«Ì·"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Clas1"
         Caption         =   "ò·«” 1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "Op"
         Caption         =   "«Å—« Ê—"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "D"
         Caption         =   " «—ÌŒ À»  «ÿ·«⁄« "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "Clas2"
         Caption         =   "ò·«”2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "Clas3"
         Caption         =   "ò·«” 3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "Clas4"
         Caption         =   "ò·«” 4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "Clas5"
         Caption         =   "ò·«” 5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "NF"
         Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column26 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Â„Â"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      TabIndex        =   24
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "”«Ì—"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   23
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   " ⁄Âœ"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   22
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   " «ŒÌ—"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   21
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "„—Œ’Ì"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      TabIndex        =   20
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "€Ì»  „ÊÃÂ"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6960
      TabIndex        =   19
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "€Ì»  €Ì— „ÊÃÂ"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8280
      TabIndex        =   18
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000002&
      Caption         =   "„Ãœœ"
      Height          =   300
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "À»  €Ì»  Â„—«Â »« òœ ò·«”"
      Height          =   540
      Left            =   16440
      TabIndex        =   116
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ﬁ—¬‰ ¬„Ê“«‰"
      Height          =   495
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      DisabledPicture =   "QeybatF.frx":14C6
      DownPicture     =   "QeybatF.frx":26140
      DragIcon        =   "QeybatF.frx":4ADBA
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9360
      Picture         =   "QeybatF.frx":6FA34
      Style           =   1  'Graphical
      TabIndex        =   113
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      DisabledPicture =   "QeybatF.frx":946AE
      DownPicture     =   "QeybatF.frx":B9328
      DragIcon        =   "QeybatF.frx":DDFA2
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9360
      Picture         =   "QeybatF.frx":102C1C
      Style           =   1  'Graphical
      TabIndex        =   112
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "À»  €Ì» "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame7 
      Caption         =   "Ê÷⁄Ì  €Ì  ﬁ—¬‰ ¬„Ê“"
      Height          =   3735
      Left            =   120
      TabIndex        =   88
      Top             =   0
      Width           =   5775
      Begin VB.ListBox List1 
         BackColor       =   &H80000003&
         ForeColor       =   &H80000007&
         Height          =   360
         ItemData        =   "QeybatF.frx":127896
         Left            =   120
         List            =   "QeybatF.frx":127898
         TabIndex        =   130
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox CHF 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã“∆Ì« "
         Height          =   420
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "‰„«Ì‘ Ã“∆Ì« "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Ê—Êœ"
         Height          =   300
         Left            =   5040
         TabIndex        =   162
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tshoro"
         DataSource      =   "stu2class"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3720
         TabIndex        =   161
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1320
         TabIndex        =   153
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   " «ŒÌ—"
         Height          =   300
         Left            =   1800
         TabIndex        =   152
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   " ⁄Âœ"
         Height          =   300
         Left            =   720
         TabIndex        =   125
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   360
         TabIndex        =   124
         Top             =   1440
         Width           =   120
      End
      Begin VB.Line Line11 
         X1              =   1800
         X2              =   1800
         Y1              =   2880
         Y2              =   3600
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ Å—œ«Œ  ÊœÌ⁄Â"
         ForeColor       =   &H00C000C0&
         Height          =   300
         Left            =   360
         TabIndex        =   123
         Top             =   2880
         Width           =   1170
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tozih"
         DataSource      =   "stu2class"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   480
         TabIndex        =   118
         Top             =   2280
         Width           =   120
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
         Height          =   300
         Left            =   2880
         TabIndex        =   117
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2400
         TabIndex        =   115
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "„—Œ’Ì"
         Height          =   300
         Left            =   2760
         TabIndex        =   114
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tpayan"
         DataSource      =   "stu2class"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3720
         TabIndex        =   111
         Top             =   2280
         Width           =   120
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Õ–›"
         Height          =   300
         Left            =   5040
         TabIndex        =   110
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Elat"
         DataSource      =   "stu2class"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   480
         TabIndex        =   109
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "⁄·  Õ–› «“ ò·«”"
         Height          =   300
         Left            =   2280
         TabIndex        =   108
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         DataField       =   "Clas5"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2160
         TabIndex        =   107
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         DataField       =   "Clas4"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2880
         TabIndex        =   106
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         DataField       =   "Clas3"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3600
         TabIndex        =   105
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         DataField       =   "Clas2"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   4440
         TabIndex        =   104
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         DataField       =   "Clas1"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5160
         TabIndex        =   103
         Top             =   3240
         Width           =   45
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   5640
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "ò·«”5"
         Height          =   300
         Left            =   2160
         TabIndex        =   102
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "ò·«”4"
         Height          =   300
         Left            =   2880
         TabIndex        =   101
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” 3"
         Height          =   300
         Left            =   3600
         TabIndex        =   100
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” 2"
         Height          =   300
         Left            =   4440
         TabIndex        =   99
         Top             =   2880
         Width           =   435
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” 1"
         Height          =   300
         Left            =   5160
         TabIndex        =   98
         Top             =   2880
         Width           =   405
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   5640
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3480
         TabIndex        =   97
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "€Ì— „ÊÃÂ"
         Height          =   300
         Left            =   3960
         TabIndex        =   96
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   4920
         TabIndex        =   95
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "„ÊÃÂ"
         Height          =   300
         Left            =   5280
         TabIndex        =   94
         Top             =   1440
         Width           =   315
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5640
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2280
         TabIndex        =   93
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò· €Ì  Â«Ì ﬁ—¬‰ ¬„Ê“ œ— «Ì‰ ò·«”"
         Height          =   300
         Left            =   3000
         TabIndex        =   92
         Top             =   960
         Width           =   2550
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "ò·«”"
         Height          =   300
         Left            =   2880
         TabIndex        =   91
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3480
         TabIndex        =   90
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò·«” Â«Ì ‘—ò  ò—œÂ"
         Height          =   300
         Left            =   3960
         TabIndex        =   89
         Top             =   360
         Width           =   1680
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "QeybatF.frx":12789A
      Height          =   4455
      Left            =   120
      TabIndex        =   87
      Top             =   4560
      Visible         =   0   'False
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
      DefColWidth     =   120
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "·Ì”  €Ì  Â«"
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "‰«„"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Famil"
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Clas"
         Caption         =   "ò·«”"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Tarh"
         Caption         =   "ÿ—Õ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Sal"
         Caption         =   "”«·"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Mah"
         Caption         =   "„«Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Rooz"
         Caption         =   "—Ê“"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Noe"
         Caption         =   "‰Ê⁄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Elat"
         Caption         =   "⁄· "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Tozih"
         Caption         =   " Ê÷ÌÕ« "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Clas"
         Caption         =   "ò·«”"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Op"
         Caption         =   "«Å—« Ê—"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "D"
         Caption         =   " «—ÌŒ À» "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Vazeyat"
         Caption         =   "Ê÷⁄Ì  »——”Ì"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Natije"
         Caption         =   "‰ ÌÃÂ »——”Ì"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "EMTAHANAT"
         Caption         =   "«„ Õ«‰« "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   3270.047
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   3735
      Left            =   6000
      TabIndex        =   66
      Top             =   0
      Width           =   3255
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   " «"
         Height          =   330
         Left            =   720
         TabIndex        =   86
         Top             =   1800
         Width           =   120
      End
      Begin VB.Label lzsho 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "ZamaneShoro"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   85
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label lmadras 
         AutoSize        =   -1  'True
         Caption         =   "-  "
         DataField       =   "Madras"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   84
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label lostad 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Ostad"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   83
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label lzpa 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "ZamanePayan"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   82
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label lmaqta 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Maqta"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   81
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label ltarh 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tarh"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   80
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lkodclass 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "KodClass"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   240
         TabIndex        =   79
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "„œ—”"
         Height          =   330
         Left            =   2040
         TabIndex        =   78
         Top             =   2160
         Width           =   405
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄ "
         Height          =   330
         Left            =   2040
         TabIndex        =   77
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
         Height          =   330
         Left            =   2040
         TabIndex        =   76
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "„ﬁÿ⁄"
         Height          =   330
         Left            =   2040
         TabIndex        =   75
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   330
         Left            =   2040
         TabIndex        =   74
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
         Height          =   330
         Index           =   0
         Left            =   2040
         TabIndex        =   73
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lroz 
         AutoSize        =   -1  'True
         Caption         =   "-  "
         DataField       =   "AyameHafte"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   72
         Top             =   3120
         Width           =   225
      End
      Begin VB.Label ltsho 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "TShoro"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   71
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "—Ê“ Â«Ì ò·«”"
         Height          =   345
         Left            =   2040
         TabIndex        =   70
         Top             =   3240
         Width           =   885
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ ‘—Ê⁄"
         Height          =   345
         Left            =   2040
         TabIndex        =   69
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label ltpa 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "TPayan"
         DataSource      =   "MClass"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   68
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å«Ì«‰"
         Height          =   345
         Left            =   2040
         TabIndex        =   67
         Top             =   2880
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "«’·«Õ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Õ–› €Ì» "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "«ÿ·«⁄«  €Ì» "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   13080
      TabIndex        =   39
      Top             =   1200
      Width           =   3255
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "QeybatF.frx":1278AF
         Left            =   1200
         List            =   "QeybatF.frx":1278B1
         TabIndex        =   134
         Text            =   "›—Ê—œÌ‰"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2400
         TabIndex        =   133
         Text            =   "01"
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "QeybatF.frx":1278B3
         Left            =   240
         List            =   "QeybatF.frx":1278B5
         TabIndex        =   132
         Text            =   "1391"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   131
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   2
         Text            =   "€Ì»  €Ì— „ÊÃÂ"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Text            =   " Ê÷ÌÕ« "
         Top             =   1800
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ã” ÊÃÊ œ—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   16440
      TabIndex        =   36
      Top             =   0
      Width           =   1455
      Begin VB.OptionButton Option7 
         Caption         =   "€Ì»  Â«"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "ﬁ—¬‰ ¬„Ê“«‰"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ﬁ»·Ì"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   35
      Top             =   11400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "»⁄œÌ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   34
      Top             =   11520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ã” ÊÃÊ »— «”«”"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   16440
      TabIndex        =   26
      Top             =   960
      Width           =   1455
      Begin VB.OptionButton Option10 
         Alignment       =   1  'Right Justify
         Caption         =   " Ê÷ÌÕ« "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton Option9 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ Åœ—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ Œ«‰Ê«œêÌ ‘„«—Â Å—Ê‰œÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         TabIndex        =   9
         Top             =   315
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂœ „·Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ «” «œ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂœ ﬂ·«”"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   129
      Top             =   9000
      Width           =   18240
      _ExtentX        =   32173
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "ò«—»— Ã«—Ì"
            TextSave        =   "ò«—»— Ã«—Ì"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   " «—ÌŒ «„—Ê“"
            TextSave        =   " «—ÌŒ «„—Ê“"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "„‘Œ’«  ﬁ—¬‰ ¬„Ê“"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   9360
      TabIndex        =   27
      Top             =   0
      Width           =   3615
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Xselect"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   3480
         TabIndex        =   151
         Top             =   2520
         Width           =   60
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   " ·›‰ À«» "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   50
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Â„—«Â"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   49
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tell"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   48
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Mob"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   47
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tozih"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   45
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "famil"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   33
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "name"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   32
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Parvande"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   480
         TabIndex        =   31
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   30
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "‰«„"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   29
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   28
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "„‘Œ’«  €Ì» "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   9360
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   65
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "‰«„"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   64
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   63
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "ﬂ·«”"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   62
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Parvande"
         DataSource      =   "Qeybat"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   480
         TabIndex        =   61
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "name"
         DataSource      =   "Qeybat"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   60
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "famil"
         DataSource      =   "Qeybat"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   59
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "-  "
         DataField       =   "Clas"
         DataSource      =   "Qeybat"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   58
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   57
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "tozih"
         DataSource      =   "Qeybat"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   56
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Noe"
         DataSource      =   "Qeybat"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   55
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "elat"
         DataSource      =   "Qeybat"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   54
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   53
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "⁄· "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   52
         Top             =   1680
         Width           =   270
      End
   End
   Begin VB.Label sen_lable 
      AutoSize        =   -1  'True
      Caption         =   "”‰"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   12600
      TabIndex        =   159
      Top             =   3720
      Width           =   210
   End
   Begin VB.Label maqta_lable 
      AutoSize        =   -1  'True
      Caption         =   "„ﬁÿ⁄"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   10920
      TabIndex        =   158
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "‰Ê⁄ €Ì» "
      Height          =   285
      Left            =   9720
      TabIndex        =   150
      Top             =   3645
      Width           =   585
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0C0C0&
      X1              =   9240
      X2              =   13080
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«“  «—ÌŒ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6480
      TabIndex        =   145
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      TabIndex        =   144
      Top             =   3960
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "€Ì»  Â«Ì „«Â"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      TabIndex        =   136
      Top             =   2880
      Width           =   870
   End
   Begin VB.Label Label78 
      AutoSize        =   -1  'True
      Caption         =   "òœ ò·«” —« »——”Ì ò‰Ìœ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   6600
      TabIndex        =   121
      Top             =   4200
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Line Line10 
      X1              =   7560
      X2              =   9120
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line9 
      X1              =   7560
      X2              =   7560
      Y1              =   4320
      Y2              =   3960
   End
   Begin VB.Line Line8 
      X1              =   7560
      X2              =   9120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line7 
      X1              =   9120
      X2              =   9120
      Y1              =   3960
      Y2              =   4320
   End
   Begin VB.Label Label77 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   7800
      TabIndex        =   120
      Top             =   3960
      Width           =   75
   End
   Begin VB.Label Label74 
      AutoSize        =   -1  'True
      Caption         =   " ⁄œ«œ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8640
      TabIndex        =   119
      Top             =   3960
      Width           =   360
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      X1              =   9240
      X2              =   9240
      Y1              =   3840
      Y2              =   4440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      X1              =   120
      X2              =   9240
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "«ÿ·«⁄«  œ— ”Ì” „ À»  ‘œ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11160
      TabIndex        =   46
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15600
      TabIndex        =   44
      Top             =   960
      Width           =   75
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      TabIndex        =   43
      Top             =   11520
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6960
      TabIndex        =   42
      Top             =   11520
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7320
      TabIndex        =   41
      Top             =   11520
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "„Ê—œ Ì«›  ‘œ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   14160
      TabIndex        =   40
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ã” ÃÊ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15600
      TabIndex        =   0
      Top             =   120
      Width           =   420
   End
   Begin VB.Menu mnuhome 
      Caption         =   "#"
   End
   Begin VB.Menu m9 
      Caption         =   " ‰ŸÌ„« "
      Begin VB.Menu mnujoz 
         Caption         =   "›⁄«· »Êœ‰ Ã“∆Ì« "
      End
      Begin VB.Menu mnuqeybatclass 
         Caption         =   "À»  €Ì»  »« òœ ò·«”"
      End
      Begin VB.Menu mnugotrei 
         Caption         =   "..."
         Shortcut        =   ^G
      End
      Begin VB.Menu m111 
         Caption         =   "Å«Ìê«Â ›⁄«·"
         Begin VB.Menu mnubank 
            Caption         =   "«ÿ·«⁄«  ﬁ—¬‰ ¬„Ê“«‰"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuqey 
            Caption         =   "€Ì»  Â«"
         End
      End
      Begin VB.Menu mnuxls 
         Caption         =   "«‰ ﬁ«· «ÿ·«⁄«  »Â »—‰«„Â «ò”·"
      End
      Begin VB.Menu sdfghsdfgfdsg 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnutaahodKatbi 
         Caption         =   " ‰ŸÌ„«   ⁄Âœ ò »Ì"
      End
      Begin VB.Menu mnuChaptaahod 
         Caption         =   "ç«Å  ⁄Âœ ‰«„Â"
      End
   End
   Begin VB.Menu mnq 
      Caption         =   "€Ì»  Â«"
      Begin VB.Menu mnunemayeshjoz 
         Caption         =   "‰„«Ì‘ Ã“∆Ì« "
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuyesterday 
         Caption         =   "·Ì”  €Ì»  Â«Ì œÌ—Ê“"
      End
   End
   Begin VB.Menu mnusuort 
      Caption         =   "„— » ”«“Ì"
      WindowList      =   -1  'True
      Begin VB.Menu mnuoarvande 
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuname 
         Caption         =   "‰«„"
         Shortcut        =   {F6}
      End
      Begin VB.Menu nufamil 
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         Shortcut        =   {F7}
      End
      Begin VB.Menu dsfgd 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuxselext 
         Caption         =   "«÷«›Â »Â ·Ì”  „‰ Œ»"
         Shortcut        =   ^S
      End
      Begin VB.Menu MNUDELETELIST 
         Caption         =   "Õ–› «“ ·Ì”  «‰ Œ«»Ì"
      End
      Begin VB.Menu MNUWIV 
         Caption         =   "‰„«Ì‘ ·Ì”  «‰ Œ«»Ì"
      End
      Begin VB.Menu MNUCLEAN 
         Caption         =   "Å«ò”«“Ì ·Ì”  «‰ Œ«»Ì"
      End
   End
   Begin VB.Menu mnubakhsh 
      Caption         =   "»Œ‘ Â«"
      Begin VB.Menu mnusabtquranAmooz 
         Caption         =   "À»  «ÿ·«⁄«  ﬁ—¬‰ ¬„Ê“"
      End
      Begin VB.Menu mnumodirclass 
         Caption         =   "„œÌ—Ì  ·Ì”  ò·«”Ì"
      End
      Begin VB.Menu mnugovahiname 
         Caption         =   "êÊ«ÂÌ ‰«„Â"
      End
      Begin VB.Menu mnusabtnom 
         Caption         =   "À»  ‰„—«  «„ Õ«‰"
      End
      Begin VB.Menu mnukarname 
         Caption         =   "ò«—‰«„Â"
      End
   End
End
Attribute VB_Name = "QeybatF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Ds As Date
 Dim ADOP As String
 




Private Sub Check2_Click()
If Check2.Value = 0 Then
Combo8.Enabled = False
Combo9.Enabled = False
Combo10.Enabled = False
Combo11.Enabled = False
Combo12.Enabled = False
Combo13.Enabled = False
Else
Combo8.Enabled = True
Combo9.Enabled = True
Combo10.Enabled = True
Combo11.Enabled = True
Combo12.Enabled = True
Combo13.Enabled = True
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Command1.Default = True
Else
Command1.Default = False
End If

End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If Len(Combo1.Text) = 2 Then
If Combo1.Text = "01" Then Combo1.Text = "01" & "-" & "›—Ê—œÌ‰"
If Combo1.Text = "02" Then Combo1.Text = "02" & "-" & "«—œÌ»Â‘ "
If Combo1.Text = "03" Then Combo1.Text = "03" & "-" & "Œ—œ«œ"
If Combo1.Text = "04" Then Combo1.Text = "04" & "-" & " Ì—"
If Combo1.Text = "05" Then Combo1.Text = ("05" & "-" & "„—œ«œ")
If Combo1.Text = "06" Then Combo1.Text = ("06" & "-" & "‘Â—ÌÊ—")
If Combo1.Text = "07" Then Combo1.Text = ("07" & "-" & "„Â—")
If Combo1.Text = "08" Then Combo1.Text = ("08" & "-" & "¬»«‰")
If Combo1.Text = "09" Then Combo1.Text = ("09" & "-" & "¬–—")
If Combo1.Text = "10" Then Combo1.Text = ("10" & "-" & "œÌ")
If Combo1.Text = "11" Then Combo1.Text = ("11" & "-" & "»Â„‰")
If Combo1.Text = "12" Then Combo1.Text = ("12" & "-" & "«”›‰œ")
Combo3.SetFocus


'”Ì»”»”Ì»”»

End If
End Sub


Private Sub Combo14_Click()
If Combo7.Text = " „«„Ì „«Â Â«" Then
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and noe like ('%" + Combo14.Text + "%')"
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
Else
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and mah like ('%" & Combo7.Text & "')and noe like ('%" + Combo14.Text + "%')"
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
End If

End Sub


Private Sub Combo15_Click()
If Combo7.Text = " „«„Ì „«Â Â«" Then
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')   and sal like ('%" + Combo15.Text + "%') "
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
Else
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and mah like ('%" & Combo7.Text & "')  and sal like ('%" + Combo15.Text + "%')"
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
End If
End Sub

Private Sub Combo2_Click()
On Error GoTo 9898
GoTo 9999

9898:
MsgBox "  ‰ŸÌ„«   ⁄Âœ —« »——”Ì ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub




9999:
If Combo2.Text = " «ŒÌ—" Then
Combo5.Clear
Combo5.AddItem ("œﬁÌﬁÂ 5")
Combo5.AddItem ("œﬁÌﬁÂ 10")
Combo5.AddItem ("œﬁÌﬁÂ 15")
Combo5.AddItem ("œﬁÌﬁÂ 20")
Combo5.Text = Combo5.List(0)



Else
Combo5.Clear
Combo5.AddItem ("⁄· ")
Combo5.Text = Combo5.List(0)
End If


If Combo2.Text = " ⁄Âœ ò »Ì" Then
Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + "QeybatF-TaahodKatbi-Text" + "%')"
Setting.Refresh

Setting.Recordset.Sort = "xsort"

Setting.Recordset.MoveFirst


Combo5.Clear
For I = 1 To Setting.Recordset.RecordCount

Combo5.AddItem (Setting.Recordset.Fields("xsort") & "-" & Setting.Recordset.Fields("xname"))
Setting.Recordset.MoveNext
Next I
Combo5.Text = Combo5.List(0)
End If


End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)

If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
If Len(Combo3.Text) = 2 Then
Combo1.SetFocus

End If
End Sub


Private Sub Combo4_Change()
On Error Resume Next
If CHF.Value = 0 Then Exit Sub
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Combo4.Text + "%')"
mclass.Refresh
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" + Combo4.Text + "%')and parvande like ('%" + Label8.Caption + "%')"
STU2CLASS.Refresh

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "€Ì— „ÊÃÂ" + "%')"
Qeybat.Refresh
Label62.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "€Ì»  „ÊÃÂ" + "%')"
Qeybat.Refresh
Label60.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "„—Œ’Ì" + "%')"
Qeybat.Refresh
Label73.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and clas like ('%" + Combo4.Text + "%') and noe like ('%" + " ⁄Âœ" + "%')"
Qeybat.Refresh
Label81.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%')"
Qeybat.Refresh
Label58.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%')and noe like ('%" + " «ŒÌ—" + "%')"
Qeybat.Refresh
Label54.Caption = Qeybat.Recordset.RecordCount

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')"
Qeybat.Refresh
End Sub

Private Sub Combo4_Click()

On Error Resume Next

If CHF.Value = 0 Then Exit Sub
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Combo4.Text + "%')"
mclass.Refresh
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" + Combo4.Text + "%')and parvande like ('%" + Label8.Caption + "%')"
STU2CLASS.Refresh

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "€Ì— „ÊÃÂ" + "%')"
Qeybat.Refresh
Label62.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "€Ì»  „ÊÃÂ" + "%')"
Qeybat.Refresh
Label60.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "„—Œ’Ì" + "%')"
Qeybat.Refresh
Label73.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and clas like ('%" + Combo4.Text + "%') and noe like ('%" + " ⁄Âœ" + "%')"
Qeybat.Refresh
Label81.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%')"
Qeybat.Refresh
Label58.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%')and noe like ('%" + " «ŒÌ—" + "%')"
Qeybat.Refresh
Label54.Caption = Qeybat.Recordset.RecordCount

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')"
Qeybat.Refresh

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Combo7_Click()

If Combo7.Text = " „«„Ì „«Â Â«" Then
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')   and sal like ('%" + Combo15.Text + "%') "
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
Else
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and mah like ('%" & Combo7.Text & "')  and sal like ('%" + Combo15.Text + "%')"
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
End If

End Sub

Private Sub Command1_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-newsabt" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Dim C, X As String

Dim TC As Integer

If lkodclass.Caption = "" Then
MsgBox "ò·«” ﬁ—¬‰ ¬„Ê“ —« «‰ Œ«» ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If






'⁄œ„  ò—«—Ì »Êœ‰ €Ì» 
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and sal like('%" + Combo6.Text + "%')and mah like('%" + Combo1.Text + "%')and rooz like('%" + Combo3.Text + "%')"
Qeybat.Refresh
If Qeybat.Recordset.BOF = True Or Qeybat.Recordset.EOF = True Then
GoTo 17
Else
MsgBox "  »—«Ì «Ì‰ ﬁ—¬‰ ¬„Ê“ œ— «Ì‰  «—ÌŒ " & Qeybat.Recordset.Fields("noe") & " À»  ‘œÂ «”  ", vbExclamation + vbOKOnly, "À»  €Ì» "

'MsgBox "‘„« ﬁ»·« «Ì‰ €Ì»  —« À»  ò—œÂ «Ìœ", vbInformation, "À»  €Ì» "
Exit Sub
End If
17: ' €Ì»   ò—«—Ì ‰Ì” 


'⁄œ„  ò—«—Ì »Êœ‰ €Ì» 

' çò ò—œ‰  «—ÌŒ  €Ì» 

If Combo1.DataChanged = False Then
MsgBox "·ÿ›«  «—ÌŒ À»  €Ì»  —« »——”Ì ﬂ‰Ìœ", vbExclamation, "”«„«‰Â À»  €Ì» "
Exit Sub
Else
GoTo 15
End If
15:  'Å«Ì«‰ çò ò—œ‰  «—ÌŒ €Ì» 





'À»   ⁄Âœ« 
'If Combo2.Text = " ⁄Âœ ò »Ì" Then
'Call Command21_Click
'End If


'«Ì‰ Ã« Ã«ÌÌ «”  òÂ »«Ìœ œ— »«—Â À»  òœ ò·«” ò«— ò—œ

' À»  €Ì  «“ ÿ—Ìﬁ òœ ò·«”Ì
If Check1.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Text1.Text + "%')"
mclass.Refresh

If mclass.Recordset.BOF = True Or mclass.Recordset.EOF = True Or mclass.Recordset.RecordCount > 1 Then
MsgBox "òœ ò·«” «‘ »«Â «”  ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
Else
'À»  €Ì»  »«  ÊÃÂ »Â òœ ò·«”



Qeybat.Refresh
Qeybat.Recordset.AddNew
Qeybat.Recordset.Fields("Parvande") = Label8.Caption
Qeybat.Recordset.Fields("name") = Label9.Caption
Qeybat.Recordset.Fields("famil") = Label10.Caption
'Qeybat.Recordset.Fields("Ostad") = Label12.Caption
'Qeybat.Recordset.Fields("tarh") = Label11.Caption
Qeybat.Recordset.Fields("sal") = Combo6
Qeybat.Recordset.Fields("mah") = Combo1.Text
Qeybat.Recordset.Fields("rooz") = Combo3.Text
Qeybat.Recordset.Fields("noe") = Combo2.Text
Qeybat.Recordset.Fields("elat") = Combo5
Qeybat.Recordset.Fields("tozih") = Text7
Qeybat.Recordset.Fields("clas") = Text1.Text

Qeybat.Recordset.Fields("vazeyat") = "0"
Qeybat.Recordset.Fields("EMTAHANAT") = "»——”Ì ‰‘œÂ"
Qeybat.Recordset.Fields("d") = Taqvim.Tarikh.Caption
Qeybat.Recordset.Fields("natije") = "‰œ«—œ"
Qeybat.Recordset.Fields("KodQeybat") = Combo6.Text & mid(Combo1.Text, 1, 2) & Combo3.Text

Qeybat.Recordset.Update
Qeybat.Refresh





'À»   ⁄Âœ« 
If Combo2.Text = " ⁄Âœ ò »Ì" Then
Call Command21_Click
End If



'»——”Ì 4  «ŒÌ— »—«Ì À»  €Ì» 
If Combo2.Text = " «ŒÌ—" Then

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and noe like ('%" + " «ŒÌ—" + "%')and CLAS like ('%" + Text1.Text + "%') and vazeyat like ('%" + "0" + "%')"
Qeybat.Refresh
If Qeybat.Recordset.RecordCount >= 4 Then


'À»  €Ì»  ŒÊœ ò«—


Qeybat.Refresh
Qeybat.Recordset.AddNew
Qeybat.Recordset.Fields("Parvande") = Label8.Caption
Qeybat.Recordset.Fields("name") = Label9.Caption
Qeybat.Recordset.Fields("famil") = Label10.Caption
'Qeybat.Recordset.Fields("Ostad") = Label12.Caption
'Qeybat.Recordset.Fields("tarh") = Label11.Caption
Qeybat.Recordset.Fields("sal") = Combo6
Qeybat.Recordset.Fields("mah") = Combo1.Text
Qeybat.Recordset.Fields("rooz") = Combo3.Text
Qeybat.Recordset.Fields("noe") = "€Ì»  €Ì— „ÊÃÂ"
Qeybat.Recordset.Fields("elat") = "çÂ«—  «ŒÌ—"
Qeybat.Recordset.Fields("tozih") = "À»   Ê”ÿ ”Ì” „"
Qeybat.Recordset.Fields("clas") = Text1.Text

Qeybat.Recordset.Fields("vazeyat") = "0"
Qeybat.Recordset.Fields("EMTAHANAT") = "»——”Ì ‰‘œÂ"
Qeybat.Recordset.Fields("d") = Taqvim.Tarikh.Caption
Qeybat.Recordset.Fields("KodQeybat") = Combo6.Text & mid(Combo1.Text, 1, 2) & Combo3.Text
Qeybat.Recordset.Fields("natije") = "‰œ«—œ"
Qeybat.Recordset.Update
Qeybat.Refresh


Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and noe like ('%" + " «ŒÌ—" + "%')and CLAS like ('%" + Text1.Text + "%') and vazeyat like ('%" + "0" + "%')"
Qeybat.Refresh

For w = 1 To Qeybat.Recordset.RecordCount
Qeybat.Recordset.Fields("vazeyat") = " «ŒÌ— »——”Ì ‘œ"
Qeybat.Recordset.Fields("natije") = "À»  ŒÊœò«— €Ì»  €Ì— „ÊÃÂ"
Qeybat.Recordset.Update
Qeybat.Recordset.MoveNext

Next w


MsgBox "À»  €Ì»  €Ì— „ÊÃÂ »Â ’Ê—  ŒÊœò«— »Â œ·Ì· 4  «ŒÌ— «‰Ã«„ ‘œ ", vbInformation + vbOKOnly, "À»  ŒÊœò«— €Ì» "

'Å«Ì«‰ À»  €Ì»  ŒÊœò«—
End If '»—«Ì »“—ê — «“ 4 »Êœ‰
End If

'Å«Ì«‰ »——”Ì 4  «ŒÌ— »—«Ì À»  €Ì» «








Label25.Caption = Qeybat.Recordset.RecordCount
DataGrid3.Refresh

Beep

Label27.Visible = True

End If
If Check3.Value = 1 Then Combo3.SetFocus

Exit Sub

End If


' €Ì»   »« ”«” ›«Âœ «“ òœ ò·”« Å«Ì«‰ À» 






'«Ì‰ òÂ ÂÌÃÌ ‰Ì”  ›ﬁÿ ⁄œœ «” 

X = 1
TC = 0
'«Ì‰ òÂ ÂÌÃÌ ‰Ì” 


For I = 1 To 5   ' Å‰Ã ò·«” —« çò „Ì ò‰œ

C = "clas" & X
If Student.Recordset.Fields(C) <> "‰œ«—œ" Then
TC = TC + 1 '‘„«—‘ ò·«” Â«

End If
X = X + 1
Next I
If TC > 1 Then  'ÿ—› »Ì” — «“ 1 ‘—ò  „Ì ò‰œ
Beep

If MsgBox("«Ì‰ ﬁ—¬‰ ¬„Ê“ Â„ «ò‰Ê‰ œ— »Ì‘ «“ Ìò ò·«” ‘—ò  „Ì ò‰œ   ¬Ì« «“ ’Õ  òœ ò·«” «ÿ„Ì‰«‰ œ«—Ìœ ", vbQuestion + vbYesNo, "À»  €Ì» ") = vbYes Then

GoTo 12 '«ÿ„Ì‰«‰ œ«—œ «“ òœ ò·«”


Else  ' «“ ’Õ  òœ ò·«” «ÿ„Ì‰«‰ ‰œ«—œ
If Check3.Value = 1 Then Combo3.SetFocus

Exit Sub
End If



Else 'ÿ—› Ìò ò·«” „Ì —ÊœÅ



GoTo 13  '«Ì‰ »—«Ì «Ì‰ «”  òÂ «“ »«·« Ê«—œ ‰‘Êœ Ê ›ﬁÿ «“ ÿ—› òœ 12 Ê«—œ ‘Êœ

12:



Qeybat.Refresh
Qeybat.Recordset.AddNew
Qeybat.Recordset.Fields("Parvande") = Label8.Caption
Qeybat.Recordset.Fields("name") = Label9.Caption
Qeybat.Recordset.Fields("famil") = Label10.Caption
'Qeybat.Recordset.Fields("Ostad") = Label12.Caption
'Qeybat.Recordset.Fields("tarh") = Label11.Caption
Qeybat.Recordset.Fields("sal") = Combo6
Qeybat.Recordset.Fields("mah") = Combo1.Text
Qeybat.Recordset.Fields("rooz") = Combo3.Text
Qeybat.Recordset.Fields("noe") = Combo2.Text
Qeybat.Recordset.Fields("elat") = Combo5
Qeybat.Recordset.Fields("tozih") = Text7
Qeybat.Recordset.Fields("clas") = lkodclass.Caption
Qeybat.Recordset.Fields("vazeyat") = "0"
Qeybat.Recordset.Fields("EMTAHANAT") = "»——”Ì ‰‘œÂ"
Qeybat.Recordset.Fields("d") = Taqvim.Tarikh.Caption
Qeybat.Recordset.Fields("KodQeybat") = Combo6.Text & mid(Combo1.Text, 1, 2) & Combo3.Text
Qeybat.Recordset.Fields("natije") = "‰œ«—œ"
Qeybat.Recordset.Update
Qeybat.Refresh








'À»   ⁄Âœ« 
If Combo2.Text = " ⁄Âœ ò »Ì" Then
Call Command21_Click
End If




'»——”Ì 4  «ŒÌ— »—«Ì À»  €Ì» 
If Combo2.Text = " «ŒÌ—" Then

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and noe like ('%" + " «ŒÌ—" + "%')and CLAS like ('%" + lkodclass.Caption + "%') and vazeyat like ('%" + "0" + "%')"
Qeybat.Refresh
If Qeybat.Recordset.RecordCount >= 4 Then


'À»  €Ì»  ŒÊœ ò«—


Qeybat.Refresh
Qeybat.Recordset.AddNew
Qeybat.Recordset.Fields("Parvande") = Label8.Caption
Qeybat.Recordset.Fields("name") = Label9.Caption
Qeybat.Recordset.Fields("famil") = Label10.Caption
'Qeybat.Recordset.Fields("Ostad") = Label12.Caption
'Qeybat.Recordset.Fields("tarh") = Label11.Caption
Qeybat.Recordset.Fields("sal") = Combo6
Qeybat.Recordset.Fields("mah") = Combo1.Text
Qeybat.Recordset.Fields("rooz") = Combo3.Text
Qeybat.Recordset.Fields("noe") = "€Ì»  €Ì— „ÊÃÂ"
Qeybat.Recordset.Fields("elat") = "çÂ«—  «ŒÌ—"
Qeybat.Recordset.Fields("tozih") = "À»   Ê”ÿ ”Ì” „"
Qeybat.Recordset.Fields("clas") = lkodclass.Caption
Qeybat.Recordset.Fields("vazeyat") = "0"
Qeybat.Recordset.Fields("EMTAHANAT") = "»——”Ì ‰‘œÂ"
Qeybat.Recordset.Fields("d") = Taqvim.Tarikh.Caption
Qeybat.Recordset.Fields("KodQeybat") = Combo6.Text & mid(Combo1.Text, 1, 2) & Combo3.Text
Qeybat.Recordset.Fields("natije") = "‰œ«—œ"
Qeybat.Recordset.Update
Qeybat.Refresh





Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and noe like ('%" + " «ŒÌ—" + "%')and CLAS like ('%" + lkodclass.Caption + "%') and vazeyat like ('%" + "0" + "%')"
Qeybat.Refresh

For w = 1 To Qeybat.Recordset.RecordCount
Qeybat.Recordset.Fields("vazeyat") = " «ŒÌ— »——”Ì ‘œ"
Qeybat.Recordset.Fields("natije") = "À»  ŒÊœò«— €Ì»  €Ì— „ÊÃÂ"
Qeybat.Recordset.Update
Qeybat.Recordset.MoveNext
Next w


MsgBox "À»  €Ì»  €Ì— „ÊÃÂ »Â ’Ê—  ŒÊœò«— »Â œ·Ì· 4  «ŒÌ— «‰Ã«„ ‘œ ", vbInformation + vbOKOnly, "À»  ŒÊœò«— €Ì» "


'Å«Ì«‰ À»  €Ì»  ŒÊœò«—
End If '»—«Ì »“—ê — «“ 4 »Êœ‰
End If

'Å«Ì«‰ »——”Ì 4  «ŒÌ— »—«Ì À»  €Ì» «













GoTo 18 ' »—«Ì «Ì‰òÂ ›ÀﬁÿÌ «“18 Ê« —œ «Ì‰ ‘Êœ


13:  'Ìò ò·«” „Ì —Êœ
 
Qeybat.Refresh
Qeybat.Recordset.AddNew
Qeybat.Recordset.Fields("Parvande") = Label8.Caption
Qeybat.Recordset.Fields("name") = Label9.Caption
Qeybat.Recordset.Fields("famil") = Label10.Caption
'Qeybat.Recordset.Fields("Ostad") = Label12.Caption
'Qeybat.Recordset.Fields("tarh") = Label11.Caption
Qeybat.Recordset.Fields("sal") = Combo6
Qeybat.Recordset.Fields("mah") = Combo1.Text
Qeybat.Recordset.Fields("rooz") = Combo3.Text
Qeybat.Recordset.Fields("noe") = Combo2.Text
Qeybat.Recordset.Fields("elat") = Combo5
Qeybat.Recordset.Fields("tozih") = Text7
Qeybat.Recordset.Fields("vazeyat") = "0"
Qeybat.Recordset.Fields("natije") = "‰œ«—œ"
Qeybat.Recordset.Fields("EMTAHANAT") = "»——‘Ì ‰‘œÂ"
Qeybat.Recordset.Fields("d") = Taqvim.Tarikh.Caption
Qeybat.Recordset.Fields("KodQeybat") = Combo6.Text & mid(Combo1.Text, 1, 2) & Combo3.Text

'«Ì‰ òÂ ÂÌÃÌ
X = 1
TC = 0
'«Ì‰ òÂ ÂÌÃÌ



' çÊ‰ Ìò ò·«” œ«—œ „Ì ŒÊ«Âœ Â«‰ òœ ò·«” —« Ê«—œ ò‰œ

For I = 1 To 5
C = "clas" & X
If Student.Recordset.Fields(C) = "‰œ«—œ" Then
GoTo 16

Else
Qeybat.Recordset.Fields("clas") = Student.Recordset.Fields(C)
GoTo 14

End If
16:
X = X + 1
 Next I
14:


Qeybat.Recordset.Update

Qeybat.Refresh



'À»   ⁄Âœ« 
If Combo2.Text = " ⁄Âœ ò »Ì" Then
Call Command21_Click
End If






'»——”Ì 4  «ŒÌ— »—«Ì À»  €Ì» 
If Combo2.Text = " «ŒÌ—" Then

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and noe like ('%" + " «ŒÌ—" + "%')and CLAS like ('%" + Student.Recordset.Fields(C) + "%') and vazeyat like ('%" + "0" + "%') "
Qeybat.Refresh
If Qeybat.Recordset.RecordCount >= 4 Then


'À»  €Ì»  ŒÊœ ò«—


Qeybat.Refresh
Qeybat.Recordset.AddNew
Qeybat.Recordset.Fields("Parvande") = Label8.Caption
Qeybat.Recordset.Fields("name") = Label9.Caption
Qeybat.Recordset.Fields("famil") = Label10.Caption
'Qeybat.Recordset.Fields("Ostad") = Label12.Caption
'Qeybat.Recordset.Fields("tarh") = Label11.Caption
Qeybat.Recordset.Fields("sal") = Combo6
Qeybat.Recordset.Fields("mah") = Combo1.Text
Qeybat.Recordset.Fields("rooz") = Combo3.Text
Qeybat.Recordset.Fields("noe") = "€Ì»  €Ì— „ÊÃÂ"
Qeybat.Recordset.Fields("elat") = "çÂ«—  «ŒÌ—"
Qeybat.Recordset.Fields("tozih") = "À»   Ê”ÿ ”Ì” „"
Qeybat.Recordset.Fields("clas") = Student.Recordset.Fields(C)
Qeybat.Recordset.Fields("EMTAHANAT") = "»——”Ì ‰‘œÂ"
Qeybat.Recordset.Fields("d") = Taqvim.Tarikh.Caption
Qeybat.Recordset.Fields("KodQeybat") = Combo6.Text & mid(Combo1.Text, 1, 2) & Combo3.Text
Qeybat.Recordset.Fields("vazeyat") = "0"

Qeybat.Recordset.Fields("natije") = "‰œ«—œ"
Qeybat.Recordset.Update
Qeybat.Refresh









Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and noe like ('%" + " «ŒÌ—" + "%')and CLAS like ('%" + Student.Recordset.Fields(C) + "%') and vazeyat like ('%" + "0" + "%')"
Qeybat.Refresh

For w = 1 To Qeybat.Recordset.RecordCount
Qeybat.Recordset.Fields("vazeyat") = " «ŒÌ— »——”Ì ‘œ"
Qeybat.Recordset.Fields("natije") = "À»  ŒÊœò«— €Ì»  €Ì— „ÊÃÂ"
Qeybat.Recordset.Update
Qeybat.Recordset.MoveNext
Next w


MsgBox "À»  €Ì»  €Ì— „ÊÃÂ »Â ’Ê—  ŒÊœò«— »Â œ·Ì· 4  «ŒÌ— «‰Ã«„ ‘œ ", vbInformation + vbOKOnly, "À»  ŒÊœò«— €Ì» "






'Å«Ì«‰ À»  €Ì»  ŒÊœò«—
End If '»—«Ì »“—ê — «“ 4 »Êœ‰
End If

'Å«Ì«‰ »——”Ì 4  «ŒÌ— »—«Ì À»  €Ì» «









18:



'MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbInformation, "À»  €Ì» "
Label25.Caption = Qeybat.Recordset.RecordCount
DataGrid3.Refresh

Beep

Label27.Visible = True
'Timer1.Enabled = True

End If

'End If
'Else
If Check3.Value = 1 Then Combo3.SetFocus

Exit Sub

 'À»   ò—«—Ì


'End If
If Check3.Value = 1 Then Combo3.SetFocus


10 End Sub

Private Sub Command10_Click()
CHF.Value = 0
PB2.Visible = True
PB2.Max = Student.Recordset.RecordCount

If Option6.Value = True Then '»«Ìœ ﬁ—¬‰ ¬„Ê“«‰ —« Ê—«—œ «ò”· ò‰œ



Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String

Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

Dim DataArray(1 To 10000, 1 To 29) As Variant

Dim r As Integer
Dim NumberOfRows As Integer
NumberOfRows = Student.Recordset.RecordCount
Student.Recordset.MoveFirst
For r = 1 To NumberOfRows
DataArray(r, 1) = Student.Recordset.Fields("parvande")
DataArray(r, 2) = Student.Recordset.Fields("name")
DataArray(r, 3) = Student.Recordset.Fields("famil")
DataArray(r, 4) = Student.Recordset.Fields("namepedar")
DataArray(r, 5) = Student.Recordset.Fields("tavalod")
DataArray(r, 6) = Student.Recordset.Fields("kodmeli")
DataArray(r, 7) = Student.Recordset.Fields("tahsilat")
DataArray(r, 8) = Student.Recordset.Fields("tozih")
DataArray(r, 9) = Student.Recordset.Fields("tell")
DataArray(r, 10) = Student.Recordset.Fields("mob")
DataArray(r, 11) = Student.Recordset.Fields("d")

DataArray(r, 12) = Student.Recordset.Fields("clas1")
DataArray(r, 13) = Student.Recordset.Fields("clas2")
DataArray(r, 14) = Student.Recordset.Fields("clas3")
DataArray(r, 15) = Student.Recordset.Fields("clas4")
DataArray(r, 16) = Student.Recordset.Fields("clas5")

vadie.Refresh
vadie.RecordSource = "select * from vadie where parvande like ('%" + Student.Recordset.Fields("parvande") + "%')"
vadie.Refresh
If vadie.Recordset.BOF = True Or vadie.Recordset.EOF = True Then
DataArray(r, 17) = "Å—œ«Œ  ‰‘œÂ" ' vadie.Recordset.Fields("mablaq")
DataArray(r, 18) = "‰œ«—œ" ' vadie.Recordset.Fields("dore")
DataArray(r, 19) = "‰œ«—œ" ' vadie.Recordset.Fields("d")

DataArray(r, 20) = "0" ' vadie.Recordset.RecordCount
DataArray(r, 21) = "0" ' vadie.Recordset.Fields("vazeyat")
DataArray(r, 22) = "0" ' vadie.Recordset.Fields("merja")
DataArray(r, 23) = "‰œ«—œ" ' vadie.Recordset.Fields("verja")
Else
DataArray(r, 17) = vadie.Recordset.Fields("mablaq")
DataArray(r, 18) = vadie.Recordset.Fields("dore")
DataArray(r, 19) = vadie.Recordset.Fields("d")
DataArray(r, 20) = vadie.Recordset.RecordCount
DataArray(r, 21) = vadie.Recordset.Fields("vazeyat")
DataArray(r, 22) = vadie.Recordset.Fields("merja")
DataArray(r, 23) = vadie.Recordset.Fields("verja")
End If

Student.Recordset.MoveNext
'PB2.Value = PB2.Value + 1
Next
Set oSheet = oBook.Worksheets(1)


oSheet.Range("A1:v1").Font.Bold = True


oSheet.Range("A1 :v1").Value = Array("parvande", "name", "famil", "name pedar", "tarikh tavalod", "kod meli", "tahsilat", "tozih", "tell", "mob", "tarikh sabt", "clas1", "clas2", "clas3", "clas4", "clas5", "mablaq vadie", "dore pardakht", "tarikhsabt vadie", "tedade pardakht", "vazeyat vadie", "mablaq erja", "vazwyar erja")


oSheet.Range("A2").Resize(NumberOfRows, 29).Value = DataArray
If Option1.Value = True Then ADOP = " ‘„«—Â Å—Ê‰œÂ "
If Option2.Value = True Then ADOP = " ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ "
If Option3.Value = True Then ADOP = " òœ „·Ì "
If Option4.Value = True Then ADOP = " «” «œ "
If Option5.Value = True Then ADOP = lmaqta.Caption & lostad.Caption
If Option9.Value = True Then ADOP = " ‰«„ Åœ— "
If Option10.Value = True Then ADOP = "  Ê÷ÌÕ«  "





AD = " ﬁ—¬‰ ¬„Ê“«‰ " & ADOP & " " & Text1.Text


oBook.SaveAs AD
'oBook.SaveAs "C:\Report.xls"
oExcel.quit
Student.Recordset.MoveFirst
MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "

End If

CHF.Value = 1

PB2.Visible = False

'qeybat  qeybat    qeybat  qeybat  qeybat  qeybat  qeybat  qeybat  qeybat  qeybat  qeybat  qeybat  qeybat



End Sub

Private Sub Command11_Click()
If Command11.Caption = "ﬁ—¬‰ ¬„Ê“«‰" Then
Command11.Caption = "€Ì»  Â«"
Option7.Value = True
Call Command9_Click

Else

Command11.Caption = "ﬁ—¬‰ ¬„Ê“«‰"
Option6.Value = True

End If

End Sub

Private Sub Command12_Click()
Call Text1_Change

End Sub

Private Sub Command13_Click()
Label78.Visible = False

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where clas like ('%" + Combo4.Text + "%') and Noe like ('%" + "„—Œ’Ì" + "%') and parvande like ('%" + Label8.Caption + "%') "
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
If Qeybat.Recordset.RecordCount = 0 Then
Label78.Visible = True
End If
End Sub

Private Sub Command14_Click()
Label78.Visible = False
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where clas like ('%" + Combo4.Text + "%') and Noe like ('%" + " «ŒÌ—" + "%') and parvande like ('%" + Label8.Caption + "%') "
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
If Qeybat.Recordset.RecordCount = 0 Then
Label78.Visible = True
End If
End Sub

Private Sub Command15_Click()
Label78.Visible = False
Qeybat.Refresh

Qeybat.RecordSource = "select * from qeybat where clas like ('%" + Combo4.Text + "%') and Noe like ('%" + " ⁄Âœ" + "%') and parvande like ('%" + Label8.Caption + "%') "
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
If Qeybat.Recordset.RecordCount = 0 Then
Label78.Visible = True
End If
End Sub

Private Sub Command16_Click()
Label77.Caption = Qeybat.Recordset.RecordCount

End Sub

Private Sub Command17_Click()
Label78.Visible = False
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')"

Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount

End Sub

Private Sub Command18_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-vadie-enter" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
VadieF.Show
VadieF.Text1.Text = Me.Label8.Caption
End Sub

Private Sub Command19_Click()
QeybatFilter.Show

End Sub

Private Sub Command2_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub

14082513

If Qeybat.Recordset.RecordCount = 0 Then
o = MsgBox("‘„« ÂÌç ê“Ì‰Â «Ì  »—«Ì Õ–› ‰œ«—Ìœ ", vbCritical, "Œÿ«")
Else
o = MsgBox(" ¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰  €Ì»  —« Õ–› ﬂ‰Ìœ ", vbYesNo + vbQuestion, "Õ–› €Ì» ")

If o = vbYes Then
Qeybat.Recordset.Delete
End If
End If






End Sub

Private Sub Command20_Click()

'Qeybat.Recordset.Filter = " kodqeybat <= 13900201 and parvande like ('%" & Label8.Caption & "%')"




'On Error GoTo 1
Dim SearchData, START, Ennd, SNahaee As String

GoTo 2

1:
MsgBox "??? ???? ??? ??? ?? ?? ???? ?? ????", vbCritical + vbOKOnly, "???"



Exit Sub


Me.Qeybat.Refresh
Label1.Caption = "?? ??? ?????"
Dim I As Double

SearchData = ""

If Val(START) > Val(Ennd) Then
MsgBox "????? ????? ??? ?? ?? ????? ???? ?? ????", vbCritical + vbOKOnly, "???"
Exit Sub
End If


For I = Val(START) To Val(Ennd)
If I = Val(START) Then
SearchData = " kodqeybat like ('" & I & "')"
Else
SearchData = SearchData & " or kodqeybat like ('" & I & "')"
End If
Next I
'MsgBox ""


2:
'SNahaee = "select * from qeybat where parvande like ('%" & Label8.Caption & "%') and " & SearchData
'Me.Qeybat.Refresh
'Me.Qeybat.RecordSource = "select * from qeybat where parvande like ('%" & Label8.Caption & "%') AND " & SearchData '& '" and parvande like ('%" & Label8.Caption & "%')"
'Start = Str(Combo6.Text) & Str(Combo3.Text) & Str(Combo1.Text)
'Ennd = Str(Combo8.Text) & Str(Combo7.Text) & Str(Combo4.Text)


START = Combo8.Text & "" & Combo9.Text & "" & Combo10.Text
Ennd = Combo11.Text & "" & Combo12.Text & "" & Combo13.Text

'Me.Qeybat.RecordSource = "select * from qeybat where" & SearchData '& " and parvande like ('%" & Label8.Caption & "%')"
Me.Qeybat.Refresh
'MsgBox ""

Qeybat.Recordset.Filter = " kodqeybat >= " & START & " and kodqeybat<= " & Ennd & " and parvande like ('%" & Label8.Caption & "%')"



'Me.Qeybat.RecordSource = "select * from qeybat where parvande like ('%" & Label8.Caption & "%')"
'Me.Qeybat.Refresh


'Label77.Caption = Me.Qeybat.Recordset.RecordCount




End Sub

Private Sub Command21_Click()

'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
Dim AD As String
On Error GoTo 1
GoTo 2
1: MsgBox " ‰ŸÌ„«  —« »——”Ì ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:
If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "TaahodKatbi.xlsx")

End If
If Entekhab.net.Checked = True Then
Set oExcel = GetObject("\\yafatemeh2-pc\F\Markaz Quran & Hadis\FORMXLS\TaahodKatbi.xlsx")

End If
'oExcel.ActiveSheet.Range("f1").Value = "«„Ê— ¬„Ê“‘Ì"
'oExcel.ActiveSheet.Range("j1").Value = Text10.Text
'oExcel.ActiveSheet.Range("M1").Value = Taqvim.Label1.Caption
'ekhtar.Recordset.MoveFirst

'Set oExcel = GetObject("\\yafatemeh2-pc\F\Markaz Quran & Hadis\FORMXLS\InformationClassFormool.xlsx")

'\\yafatemeh2-pc\F\Markaz Quran & Hadis\FORMXLS\InformationClassFormool.xlsx

Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + "QeybatF-TaahodKatbi-Text" + "%') and Xsort like('%" & mid(Combo5.Text, 1, 3) & "%')"
Setting.Refresh

oExcel.ActiveSheet.Range("B4").Value = Setting.Recordset.Fields("xname")


oExcel.ActiveSheet.Range("a7").Value = Setting.Recordset.Fields("xtext")

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where noe like ('%" + " ⁄Âœ" + "%')" ' or name like ('%" + Text1.Text + "%')"
Qeybat.Refresh
'ò·  ⁄Âœ Â«
Dim XrTaahoodKod As String

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where noe like ('%" + " ⁄Âœ" + "%')" ' or name like ('%" + Text1.Text + "%')"
Qeybat.Refresh
' ⁄Âœ«  ò »Ì
XrTaahoodKod = Qeybat.Recordset.RecordCount

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where noe like ('%" + " ⁄Âœ ò »Ì" + "%')" ' or name like ('%" + Text1.Text + "%')"
Qeybat.Refresh
' ⁄Âœ«  ‘›«ÂÌ
XrTaahoodKod = XrTaahoodKod & "-" & Qeybat.Recordset.RecordCount


Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where noe like ('%" + " ⁄Âœ ‘›«ÂÌ" + "%')" ' or name like ('%" + Text1.Text + "%')"
Qeybat.Refresh
'ò·  ⁄Âœ«  «Ì‰ ﬁ—¬‰ ¬„Ê“
XrTaahoodKod = XrTaahoodKod & "-" & Qeybat.Recordset.RecordCount


Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where  noe like ('%" + " ⁄Âœ" + "%')  and parvande like ('%" + Label8.Caption + "%')"
Qeybat.Refresh
' ⁄Âœ«  ò »Ì ﬁ—¬‰ ¬„Ê“
XrTaahoodKod = XrTaahoodKod & "-" & Qeybat.Recordset.RecordCount

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where  noe like ('%" + " ⁄Âœ ò »Ì" + "%')  and parvande like ('%" + Label8.Caption + "%')"
Qeybat.Refresh
' ⁄Âœ«  ‘›«ÂÌ ﬁ—–¬‰ ª¬Ê“
XrTaahoodKod = XrTaahoodKod & "-" & Qeybat.Recordset.RecordCount


Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where  noe like ('%" + " ⁄Âœ ‘›«ÂÌ" + "%')  and parvande like ('%" + Label8.Caption + "%')"
Qeybat.Refresh
XrTaahoodKod = XrTaahoodKod & "-" & Qeybat.Recordset.RecordCount



oExcel.ActiveSheet.Range("B10").Value = XrTaahoodKod


oExcel.ActiveSheet.Range("B8").Value = Me.stb1.Panels(4).Text
oExcel.ActiveSheet.Range("B9").Value = Combo4.Text





oExcel.ActiveSheet.Range("B5").Value = Label9.Caption & " " & Label10.Caption

oExcel.ActiveSheet.Range("B6").Value = Label8.Caption



'On Error Resume Next




MsgBox " ⁄Âœ ‰«„Â ¬„«œÂ ç«Å „Ì »«‘œ", 64, "ç«Å  ⁄Âœ"
AD = Label8.Caption & "Taahod" '& 'Me.stb1.Panels(4).Text
'oExcel.SaveAs AD
'oExcel.Application.Visible = True
'oExcel.Parent.Windows(1).Visible = True


oExcel.Application.Visible = True
On Error GoTo 722


oExcel.Parent.Windows(2).Visible = True
GoTo 910
722:

oExcel.Parent.Windows(1).Visible = True
910:
''''''

oExcel.SaveAs AD
'oExcel.Close
'
'
End Sub

Private Sub Command22_Click()
tozih_form.TE1.Text = Me.Label8.Caption
tozih_form.Show

End Sub

Private Sub Command23_Click()
tozih_form.TE1.Text = Me.Label8.Caption
tozih_form.Show

End Sub

Private Sub Command3_Click()
Label78.Visible = False
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where clas like ('%" + Combo4.Text + "%') and Noe like ('%" + "€Ì— „ÊÃÂ" + "%') and parvande like ('%" + Label8.Caption + "%') "
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
If Qeybat.Recordset.RecordCount = 0 Then
Label78.Visible = True
End If

 End Sub

Private Sub Command4_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513


userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-edit" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
If Command4.Caption = "«’·«Õ" Then

'Option7.Value = True
MsgBox " „«„ «ÿ·«⁄«  œ—Ê‰ ÃœÊ· ﬁ«»·  €ÌÌ— „Ì »«‘‰œ", vbInformation, "«’·«Õ €Ì» "

'DataGrid2.AllowUpdate = True
DataGrid3.AllowUpdate = True

Command4.Caption = "–ŒÌ—Â  €ÌÌ—« "
Else
Command4.Caption = "«’·«Õ"
DataGrid3.AllowUpdate = False
'DataGrid2.AllowUpdate = False
'Student.Recordset.Update
Qeybat.Recordset.Update

End If



End Sub

Private Sub Command5_Click()
On Error Resume Next



If Command6.Enabled = False Then
Command6.Enabled = True
Label24.Caption = Label24.Caption + 1
Else
Label24.Caption = Label24.Caption + 1
End If
If Label24.Caption = Label25.Caption Then Command5.Enabled = False
If Option6.Value = True Then ' Ã” Ê ÃÊ œ— ﬁ—¬‰ ¬„Ê“«‰ ‰Â œ— €Ì» Â«

If Student.Recordset.RecordCount = 0 Then
o = MsgBox("—òÊ—œ »⁄œÌ ÊÃÊœ ‰œ«—œ ", vbCritical, " ÊÃÂ")
Else
Student.Recordset.MoveNext
If Student.Recordset.BOF = False And Student.Recordset.EOF = True Then
Student.Recordset.MoveFirst
End If
End If


Else ' Ã” ÊÃÊ œ— €Ì  Â« ‰Â œ— ﬁ—¬‰ ¬„Ê“«‰
If Option7.Value = True Then
If Qeybat.Recordset.RecordCount = 0 Then
o = MsgBox("—òÊ—œ »⁄œÌ ÊÃÊœ ‰œ«—œ ", vbCritical, " ÊÃÂ")
Else
Qeybat.Recordset.MoveNext
If Qeybat.Recordset.BOF = False And Student.Recordset.EOF = True Then
Qeybat.Recordset.MoveFirst
End If
End If


End If
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
If Command5.Enabled = False Then
Command5.Enabled = True
Label24.Caption = Label24.Caption - 1
Else
Label24.Caption = Label24.Caption - 1
End If
If Label24.Caption = 1 Then Command6.Enabled = False




If Option6.Value = True Then ' Ã” Ê ÃÊ œ— ﬁ—¬‰ ¬„Ê“«‰ ‰Â œ— €Ì» Â«

If Student.Recordset.RecordCount = 0 Then
o = MsgBox("—ﬂÊ—œ ﬁ»·Ì ÊÃÊœ ‰œ«—œ ", vbCritical, " ÊÃÂ")
Else
Student.Recordset.MovePrevious

If Student.Recordset.BOF = False And Student.Recordset.EOF = True Then
Student.Recordset.MoveFirst
End If
End If


Else ' Ã” ÊÃÊ œ— €Ì  Â« ‰Â œ— ﬁ—¬‰ ¬„Ê“«‰
If Option7.Value = True Then
If Qeybat.Recordset.RecordCount = 0 Then
o = MsgBox("—ﬂÊ—œ ﬁ»·Ì ÊÃÊœ ‰œ«—œ", vbCritical, " ÊÃÂ")
Else
Qeybat.Recordset.MovePrevious

If Qeybat.Recordset.BOF = False And Student.Recordset.EOF = True Then
Qeybat.Recordset.MoveFirst
End If
End If



End If
End If
End Sub






Private Sub Command7_Click()
If Option7.Value = True Then '»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ


Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String

Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

Dim DataArray(1 To 2000, 1 To 11) As Variant

Dim r As Integer
Dim NumberOfRows As Integer
NumberOfRows = Qeybat.Recordset.RecordCount
Qeybat.Recordset.MoveFirst

For r = 1 To NumberOfRows
DataArray(r, 1) = Qeybat.Recordset.Fields("parvande")
DataArray(r, 2) = Qeybat.Recordset.Fields("name")
DataArray(r, 3) = Qeybat.Recordset.Fields("famil")
DataArray(r, 4) = Qeybat.Recordset.Fields("ostad")
DataArray(r, 5) = Qeybat.Recordset.Fields("sal")
DataArray(r, 6) = Qeybat.Recordset.Fields("mah")
DataArray(r, 7) = Qeybat.Recordset.Fields("rooz")
DataArray(r, 8) = Qeybat.Recordset.Fields("noe")
DataArray(r, 9) = Qeybat.Recordset.Fields("elat")
DataArray(r, 10) = Qeybat.Recordset.Fields("tozih")
DataArray(r, 11) = Qeybat.Recordset.Fields("clas")



Qeybat.Recordset.MoveNext
Next
Set oSheet = oBook.Worksheets(1)


oSheet.Range("A1:K1").Font.Bold = True


oSheet.Range("A1 :K1").Value = Array("‘„«—Â Å—Ê‰œÂ", "‰«„", "›«„Ì·", "«” «œ", "”«·", "„«Â", "—Ê“", "‰Ê⁄", "⁄· ", " Ê÷ÌÕ« ")


oSheet.Range("A2").Resize(NumberOfRows, 11).Value = DataArray
If Option1.Value = True Then ADOP = " ‘„«—Â Å—Ê‰œÂ "
If Option2.Value = True Then ADOP = " ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ "
If Option3.Value = True Then ADOP = " òœ „·Ì "
If Option4.Value = True Then ADOP = " «” «œ "
If Option5.Value = True Then ADOP = " òœ ò·«” "
If Option9.Value = True Then ADOP = " ‰«„ Åœ— "
If Option10.Value = True Then ADOP = "  Ê÷ÌÕ«  "





AD = " €Ì»  Â« " & ADOP & " " & Text1.Text


oBook.SaveAs AD
'oBook.SaveAs "C:\Report.xls"
oExcel.quit
Qeybat.Recordset.MoveFirst
MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "

End If
End Sub

Private Sub Command8_Click()
Label78.Visible = False
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where clas like ('%" + Combo4.Text + "%') and Noe like ('%" + "€Ì»  „ÊÃÂ" + "%') and parvande like ('%" + Label8.Caption + "%') "
Qeybat.Refresh
Label77.Caption = Qeybat.Recordset.RecordCount
If Qeybat.Recordset.RecordCount = 0 Then
Label78.Visible = True
End If
End Sub

Private Sub Command9_Click()
Option7.Value = True

Command11.Caption = "€Ì»  Â«"
'»ŒÀ ›Ì· — ò—œ‰ €Ì»  Â«
If Combo7.Text = " „«„Ì „«Â Â«" Then


Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')"
Qeybat.Refresh

Else



Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and mah like ('%" & Combo7.Text & "')"

Qeybat.Refresh

End If

If Check2.Value = 1 Then
Call Command20_Click

End If


End Sub








Private Sub DataGrid2_Click()
If Check3.Value = 1 Then Combo3.SetFocus

End Sub

Private Sub DataGrid2_DblClick()

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Dim A As String

On Error GoTo 7
GoTo 8
7:
MsgBox "›«Ì· ÅÌœ« ‰‘œ", vbCritical + vbOKOnly, "Œÿ«"
'Scan.Hide

Exit Sub

8:

If Entekhab.Pc.Checked = True Then
A = Label8.Caption
Scan.Text1.Text = A

Scan.Show
A = SettingF.ScanAdress.Caption & A & "\" & A & ".jpg"
'A = Student.Recordset.Fields("scan")
Scan.Im1.Picture = LoadPicture(A)

Exit Sub
End If

If Entekhab.net.Checked = True Then
A = Label8.Caption
Scan.Text1.Text = A

Scan.Show
'\\Yafatemeh2-pc\f\Markaz Quran & Hadis\FormScan\Pic\9020204\9020204.jpg
A = SettingF.NetScanAdress.Caption & A & "\" & A & ".jpg"
'A = Student.Recordset.Fields("scan")
Scan.Im1.Picture = LoadPicture(A)
Exit Sub
End If

Scan.Text1.Text = A


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



End Sub

Private Sub famil_Click()
Student.Recordset.Sort = "famil"

End Sub

Private Sub Form_DblClick()
Student.Refresh
Student.RecordSource = "select * from Student where parvande like ('%" + "" + "%')"
Student.Refresh
End Sub

Private Sub Form_Load()
'user.Hide
Combo1.AddItem ("01" & "-" & "›—Ê—œÌ‰")
Combo1.AddItem ("02" & "-" & "«—œÌ»Â‘ ")
Combo1.AddItem ("03" & "-" & "Œ—œ«œ")
Combo1.AddItem ("04" & "-" & " Ì—")
Combo1.AddItem ("05" & "-" & "„—œ«œ")
Combo1.AddItem ("06" & "-" & "‘Â—ÌÊ—")
Combo1.AddItem ("07" & "-" & "„Â—")
Combo1.AddItem ("08" & "-" & "¬»«‰")
Combo1.AddItem ("09" & "-" & "¬–—")
Combo1.AddItem ("10" & "-" & "œÌ")
Combo1.AddItem ("11" & "-" & "»Â„‰")
Combo1.AddItem ("12" & "-" & "«”›‰œ")
'€Ì»  Â«Ì „«Â
Combo7.AddItem (" „«„Ì „«Â Â«")
Combo7.AddItem ("›—Ê—œÌ‰")
Combo7.AddItem ("«—œÌ»Â‘ ")
Combo7.AddItem ("Œ—œ«œ")
Combo7.AddItem (" Ì—")
Combo7.AddItem ("„—œ«œ")
Combo7.AddItem ("‘Â—ÌÊ—")
Combo7.AddItem ("„Â—")
Combo7.AddItem ("¬»«‰")
Combo7.AddItem ("¬–—")
Combo7.AddItem ("œÌ")
Combo7.AddItem ("»Â„‰")
Combo7.AddItem ("«”›‰œ")


Dim I

'Å«Ì«‰ €Ì»  Â«Ì „«Â

For I = 1390 To 1408
Combo6.AddItem (I)
Combo15.AddItem (I)
Next I


For I = 1 To 31 Step 1  ' »—«Ì Ê«—œ ﬂ—œ‰ ‘„«—Â —Ê“ œ— ÃœÊ· «ÿ·«⁄«  €Ì 
If I < 10 Then
Combo3.AddItem ("0" & I)
Else
Combo3.AddItem (I)
End If
Next I

Combo2.AddItem ("€Ì»  €Ì— „ÊÃÂ")
Combo2.AddItem ("€Ì»  „ÊÃÂ")
Combo2.AddItem ("„—Œ’Ì")
Combo2.AddItem (" «ŒÌ—")
Combo2.AddItem (" ⁄Âœ ‘›«ÂÌ")
Combo2.AddItem (" ⁄Âœ ò »Ì")
Combo2.AddItem ("”«Ì—")



Combo14.AddItem ("")
Combo14.AddItem ("€Ì»  €Ì— „ÊÃÂ")
Combo14.AddItem ("€Ì»  „ÊÃÂ")
Combo14.AddItem ("„—Œ’Ì")
Combo14.AddItem (" «ŒÌ—")
Combo14.AddItem (" ⁄Âœ ‘›«ÂÌ")
Combo14.AddItem (" ⁄Âœ ò »Ì")
Combo14.AddItem ("”«Ì—")










Me.stb1.Panels(1).Text = user.OP.Text
Me.stb1.Panels(4).Text = Taqvim.Tarikh.Caption
Me.stb1.Panels(3).Text = Taqvim.Label1.Caption




'«“  «—ÌŒ » «  «—ÌŒ


For I = 1390 To 1408
Combo8.AddItem (I)
Combo11.AddItem (I)

Next I


For I = 1 To 12 Step 1  ' »—«Ì Ê«—œ ﬂ—œ‰ ‘„«—Â —Ê“ œ— ÃœÊ· «ÿ·«⁄«  €Ì 
If I < 10 Then
Combo9.AddItem ("0" & I)
Combo12.AddItem ("0" & I)

Else
Combo9.AddItem (I)
Combo12.AddItem (I)
End If
Next I


For I = 1 To 31 Step 1  ' »—«Ì Ê«—œ ﬂ—œ‰ ‘„«—Â —Ê“ œ— ÃœÊ· «ÿ·«⁄«  €Ì 
If I < 10 Then
Combo10.AddItem ("0" & I)
Combo13.AddItem ("0" & I)

Else
Combo10.AddItem (I)
Combo13.AddItem (I)
End If
Next I




'«“  «—ÌŒ  «  «Ì—“Œ






End Sub




'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label27.Visible = False

'End Sub

Private Sub Form_Resize()
On Error Resume Next
'If QeybatF.Width < 18480 Then
'QeybatF.Width = 18480
'Exit Sub
'Else
DataGrid2.Width = QeybatF.Width - 315
DataGrid2.Height = QeybatF.Height - 5658

'End If


'If QeybatF.Height < 101700 Then
'QeybatF.Height = 101700
'Exit Sub
'Else

DataGrid3.Width = QeybatF.Width - 315
DataGrid3.Height = QeybatF.Height - 5658
'End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show
QeybatF.Hide

End Sub

Private Sub Frame8_Click()
QeybatFilter.Show

End Sub

Private Sub Label13_Change()
'mclass.Refresh
'mclass.RecordSource = "seleCt * from mclass where kodclass like ('%" + Label13.Caption + "%')"
'mclass.Refresh

End Sub

Private Sub Label38_Change()
If CHF.Value = 1 Then
If Label38.Caption = "0" Then


Combo4.Clear
End If
End If

End Sub

Private Sub Label46_Change()
'mclass.Refresh
'mclass.RecordSource = "seleCt * from mclass where kodclass like ('%" + Label46.Caption + "%')"
'mclass.Refresh
End Sub

Private Sub Label68_Change()
mclass.Refresh
mclass.RecordSource = "seleCt * from mclass where kodclass like ('%" + Label68.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label68_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label68.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label69_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label69.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label70_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label70.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label71_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label71.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label72_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label72.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label8_Change()

'Barcode1.Text = Label8.Caption
If Student.Recordset.RecordCount = 0 Then Exit Sub



maqta_lable.Visible = True
On Error GoTo 8585

If Student.Recordset.Fields("tavalod") = "" Or Student.Recordset.Fields("tavalod") = "À»  „Êﬁ " Or Student.Recordset.Fields("tavalod") = Null Then
maqta_lable.Caption = "”«·  Ê·œ À»  ‰‘œÂ «” "
sen_lable.Caption = "!"
Else
Dim saL_tAvaloD As Integer
On Error Resume Next
nsaL_tAvaloD = Val(Student.Recordset.Fields("tavalod"))
saL_tAvaloD = 1391 - Val(nsaL_tAvaloD)
sen_lable.Caption = saL_tAvaloD
If saL_tAvaloD > 0 And saL_tAvaloD < 7 Then
maqta_lable.Caption = "Œ—œ”«·"
Else

If saL_tAvaloD >= 7 And saL_tAvaloD <= 12 Then
maqta_lable.Caption = "òÊœò"
Else
If saL_tAvaloD >= 13 And saL_tAvaloD <= 15 Then
maqta_lable.Caption = "‰ÊÃÊ«‰"
Else
If saL_tAvaloD >= 16 And saL_tAvaloD <= 25 Then
maqta_lable.Caption = "ÃÊ«‰"
Else
If saL_tAvaloD >= 26 And saL_tAvaloD <= 120 Then
maqta_lable.Caption = "»“—ê”«·"
Else
If saL_tAvaloD > 120 Then
maqta_lable.Caption = "Œÿ« œ— „Õ«”»Â"
sen_lable.Caption = "!"
Else
8585:

maqta_lable.Caption = "Œÿ« œ— „Õ«”»Â"
sen_lable.Caption = "!"
End If


End If

End If
End If
End If
End If





End If


If CHF.Value = 1 Then

tozih_table.Refresh
tozih_table.RecordSource = "SELECT * FROM TOZIH_TABLE WHERE PARVANDE LIKE ('%" & Label8.Caption & "%')"
tozih_table.Refresh
List2.Clear
For w = 1 To tozih_table.Recordset.RecordCount
List2.AddItem (tozih_table.Recordset.Fields("tozih"))
tozih_table.Recordset.MoveNext

Next w
Dim T As Integer

vadie.Refresh
vadie.RecordSource = "select * from vadie where  parvande like ('%" + Label8.Caption + "%') "
vadie.Refresh
Command18.Caption = vadie.Recordset.RecordCount
List1.Clear

For T = 1 To vadie.Recordset.RecordCount
List1.AddItem (vadie.Recordset.Fields("vazeyat"))
vadie.Recordset.MoveNext

Next T
'----------------------------------------------------------------

'--------------------------------------------------------
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%')"
Qeybat.Refresh
Label58.Caption = Qeybat.Recordset.RecordCount
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where parvande like ('%" + Label8.Caption + "%')"
STU2CLASS.Refresh
Label38.Caption = STU2CLASS.Recordset.RecordCount

Dim I
If STU2CLASS.Recordset.RecordCount >= 1 Then

Combo4.Clear
For I = 1 To STU2CLASS.Recordset.RecordCount


Combo4.AddItem (STU2CLASS.Recordset.Fields("kodclass"))
STU2CLASS.Recordset.MoveNext
Next I
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where parvande like ('%" + Label8.Caption + "%')"
STU2CLASS.Refresh

Combo4.Text = Combo4.List(0)
End If

Exit Sub

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "€Ì— „ÊÃÂ" + "%')"
Qeybat.Refresh
Label62.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "€Ì»  „ÊÃÂ" + "%')"
Qeybat.Refresh
Label60.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and clas like ('%" + Combo4.Text + "%') and noe like ('%" + "„—Œ’Ì" + "%')"
Qeybat.Refresh
Label73.Caption = Qeybat.Recordset.RecordCount
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')and clas like ('%" + Combo4.Text + "%') and noe like ('%" + " ⁄Âœ" + "%')"
Qeybat.Refresh
Label81.Caption = Qeybat.Recordset.RecordCount
End If

End Sub

Private Sub Label8_Click()
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%')"
Qeybat.Refresh
Label58.Caption = Qeybat.Recordset.RecordCount
End Sub

Private Sub List1_Click()
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" + List1 + "%')"
STU2CLASS.Refresh

End Sub





Private Sub m2_Click()
BankStudent.Show

End Sub

Private Sub m3_Click()
Beep

End Sub

Private Sub m4_Click()
Gozaresh.Show

End Sub

Private Sub m6_Click()
ModiriyatCLASS.Show

End Sub

Private Sub m7_Click()
FClassroom.Show

End Sub

Private Sub m8_Click()
FClassroom.Show

End Sub

Private Sub mnubank_Click()
If mnubank.Checked = True Then
Beep
Exit Sub
Else
Call Command11_Click
mnuqey.Checked = False
mnubank.Checked = True
End If

End Sub

Private Sub mnuChaptaahod_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-taahod-print" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Call Command21_Click

End Sub

Private Sub MNUCLEAN_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-list-alldelete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Student.Refresh
Student.RecordSource = "select * from student where XSELECT like ('%" + "1" + "%')"
Student.Refresh
CHF.Value = 0

For I = 1 To Student.Recordset.RecordCount


Student.Recordset.Fields("xselect") = "0"
Student.Recordset.Update
Student.Recordset.MoveNext
Next I


Beep
End Sub



Private Sub mnudel_Click()
Call Command2_Click

End Sub

Private Sub MNUDELETELIST_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-list-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Student.Recordset.Fields("xselect") = "0"
Student.Recordset.Update
Beep
End Sub

Private Sub mnuedit_Click()
Call Command4_Click

End Sub

Private Sub mnuexit2_Click()
End

End Sub

Private Sub mnugozaresh_Click()
Gozaresh.Show

End Sub

Private Sub mnustudent_Click()
BankStudent.Show

End Sub

Private Sub mnugotrei_Click()
qeybat_Gotrei.Show

End Sub

Private Sub mnugovahiname_Click()
Govahi.Text1.Text = Me.Label8.Caption
Govahi.Show

End Sub

Private Sub mnuhome_Click()
Entekhab.Show

End Sub

Private Sub mnujoz_Click()
If CHF.Value = 0 Then
CHF.Value = 1
mnujoz.Checked = True
Else
CHF.Value = 0
mnujoz.Checked = False
End If

End Sub

Private Sub mnukarname_Click()
Karname.Text2.Text = Me.Label8.Caption
Karname.Show

End Sub

Private Sub mnumodirclass_Click()
FClassroom.Text1.Text = Me.Label8.Caption
FClassroom.Show

End Sub

Private Sub mnuname_Click()
Student.Recordset.Sort = "name"

End Sub

Private Sub mnunemayeshjoz_Click()
Call Command11_Click

End Sub

Private Sub mnuoarvande_Click()
Student.Recordset.Sort = "parvande"


End Sub

Private Sub mnuqey_Click()
If mnuqey.Checked = True Then
Beep
Exit Sub
Else

Call Command11_Click
mnuqey.Checked = True
mnubank.Checked = False
End If

End Sub

Private Sub mnuqeybatclass_Click()
If Option5.Value = False Then
Option5.Value = True
mnuqeybatclass.Checked = True
Else
Option5.Value = False
mnuqeybatclass.Checked = False
End If

End Sub

Private Sub mnusabtnom_Click()
EmtahanF.Text2.Text = Me.Label8.Caption
EmtahanF.Show

End Sub

Private Sub mnusabtqeybat_Click()
Call Command1_Click

End Sub

Private Sub mnusabtquranAmooz_Click()
BankStudent.TEP.Text = Me.Label8.Caption
 BankStudent.chj.Value = 1
 BankStudent.Show

End Sub

Private Sub mnutaahodKatbi_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-taahod-setting" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
SettingTaahod.Show

End Sub

Private Sub MNUWIV_Click()
Student.Refresh
Student.RecordSource = "select * from student where XSELECT like ('%" + "1" + "%')"
Student.Refresh
End Sub

Private Sub mnuxselext_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "qeybat-list-add" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513

Student.Recordset.Fields("xselect") = "1"
Student.Recordset.Update
Beep
Text1.Text = ""
Text1.SetFocus

End Sub

Private Sub mnuyesterday_Click()
'MsgBox Mid(Me.stb1.Panels(4).Text, 11, 2)
'MsgBox Mid(Me.stb1.Panels(4).Text, 7, 2)



 GoTo 9898
 
Dim Sal, Rooz, Mah As String

If mid(Me.stb1.Panels(4).Text, 7, 2) = "01" Then Mah = "›—Ê—œÌ‰"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "02" Then Mah = "«—œÌ»Â‘ "
If mid(Me.stb1.Panels(4).Text, 7, 2) = "03" Then Mah = "Œ—œ«œ"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "04" Then Mah = " Ì—"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "05" Then Mah = "„—œ«œ"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "06" Then Mah = "‘Â—ÌÊ—"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "07" Then Mah = "„Â—"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "08" Then Mah = "¬»«‰"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "09" Then Mah = "¬–—"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "10" Then Mah = "œÌ"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "11" Then Mah = "»Â„‰"
If mid(Me.stb1.Panels(4).Text, 7, 2) = "12" Then Mah = "«”›‰œ"
Sal = mid(Me.stb1.Panels(4).Text, 1, 5)

Rooz = Str(Val(mid(Me.stb1.Panels(4).Text, 11, 2)) - 1)
MsgBox Sal & "/ " & Mah & " /" & Rooz
'MsgBox Mah
'MsgBox Rooz
GoTo 11
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where sal like ('%" & "" & "%')" ' and mah like ('%" + Mah + "%') and rooz like ('%" & Rooz & "%')"
Qeybat.Refresh
11
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where sal like ('%" & Sal & "%') and mah like ('%" + Mah + "%') and rooz like ('%" & Rooz & "%')"
Qeybat.Refresh





9898
sh = mid(Me.stb1.Panels(4).Text, 2, 4) & mid(Me.stb1.Panels(4).Text, 7, 2)

sd = Val((mid(Me.stb1.Panels(4).Text, 10, 3)) - 1)
kod = sg & sd
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where kodqeybat like ('" & kod & "') "
Qeybat.Refresh


End Sub

Private Sub name_Click()
Student.Recordset.Sort = "name"

End Sub

Private Sub nufamil_Click()
Student.Recordset.Sort = "famil"

End Sub

Private Sub Option1_Click()
Check1.Value = 0
End Sub

Private Sub Option10_Click()
Check1.Value = 0

End Sub

Private Sub Option2_Click()
Check1.Value = 0

End Sub

Private Sub Option3_Click()
Check1.Value = 0

End Sub

Private Sub Option4_Click()
Check1.Value = 0

End Sub

Private Sub Option5_Click()
Check1.Value = 1

End Sub

Private Sub Option6_Click()
DataGrid2.Visible = True
DataGrid3.Visible = False
DataGrid3.Visible = False
Command2.Enabled = False
Command4.Enabled = False

Command1.Enabled = True ' œﬂ„Â À»  €Ì» 

Label25.Caption = "0"
Label24.Caption = "0"
Label15.Caption = "0"
Command5.Enabled = False
Command6.Enabled = False
Frame6.Visible = False
Command10.Visible = True
Command7.Visible = False

End Sub

Private Sub Option7_Click()
DataGrid3.Visible = True
DataGrid2.Visible = False
Command4.Enabled = True
DataGrid3.Visible = True
Command2.Enabled = True
Command1.Enabled = False ' œﬂ„Â À»  €Ì» 
Label25.Caption = "0"
Label24.Caption = "0"
Label15.Caption = "0"
Command5.Enabled = False
Command6.Enabled = False
Frame6.Visible = True
Command10.Visible = False

Command7.Visible = True


End Sub

Private Sub TE1_DblClick()
TE1.Text = ""

End Sub

Private Sub TE2_DblClick()
TE2.Text = ""

End Sub

Private Sub Option9_Click()
Check1.Value = 0

End Sub

Private Sub Text1_Change()

'On Error Resume Next

 
 
 

Command6.Enabled = False  ' ﬂ«„«‰œ ﬁ»·Ì

If Option6.Value = True Then 'Ã” ÊÃÊ œ— «”«„Ì ﬁ—¬‰ ¬„Ê“«‰



If Option1.Value = True Then 'Ã” ÊÃÊ »— «”«” ‘„«—Â Å—Ê‰œÂ
Student.Refresh
Student.RecordSource = "select * from student where Parvande like ('%" + Text1.Text + "%')"
Student.Refresh
DataGrid3.Refresh
End If
If Option2.Value = True Then  ' Ã” ÊÃÊ »— «”«” ‰«„ Œ«‰Ê«œêÌ
Student.Refresh
Student.RecordSource = "select * from student where famil like ('%" + Text1.Text + "%')or parvande like ('%" + Text1.Text + "%') or name like ('%" + Text1.Text + "%')or nf like ('%" + Text1.Text + "%')"
Student.Refresh
'Student.Recordset.Sort = "famil"
DataGrid3.Refresh
End If
If Option3.Value = True Then  'Ã” ÊÃÊ »— «”«” ﬂœ „·Ì
Student.Refresh
Student.RecordSource = "select * from student where kodMeli like ('%" + Text1.Text + "%')"
Student.Refresh
DataGrid3.Refresh
End If
If Option4.Value = True Then  '»— «”«” ‰«„ «” «œ
Student.Refresh
Student.RecordSource = "select * from student where ostad like ('%" + Text1.Text + "%')"
Student.Refresh
DataGrid3.Refresh
End If
If Option5.Value = True Then  'ﬂœ ﬂ·«”
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + Text1.Text + "%') or clas2 like ('%" + Text1.Text + "%') or clas3 like ('%" + Text1.Text + "%') or clas4 like ('%" + Text1.Text + "%') or clas5 like ('%" + Text1.Text + "%')"
Student.Refresh
DataGrid3.Refresh
End If

If Option9.Value = True Then  '‰«„ Åœ—
Student.Refresh
Student.RecordSource = "select * from student where Namepedar like ('%" + Text1.Text + "%')"
Student.Refresh
DataGrid3.Refresh
End If

If Option10.Value = True Then  ' Ê÷ÌÕ« 
Student.Refresh
Student.RecordSource = "select * from student where tozih like ('%" + Text1.Text + "%')"
Student.Refresh
DataGrid3.Refresh
End If
Label15.Caption = Student.Recordset.RecordCount





Else ' Ã” ÊÃÊœ— €Ì»  Â«

If Option1.Value = True Then 'Ã” ÊÃÊ »— «”«” ‘„«—Â Å—Ê‰œÂ
Qeybat.RecordSource = "select * from qeybat where Parvande like ('%" + Text1.Text + "%')"
Qeybat.Refresh
DataGrid3.Refresh
End If
If Option2.Value = True Then  ' Ã” ÊÃÊ »— «”«” ‰«„ Œ«‰Ê«œêÌ

Qeybat.RecordSource = "select * from qeybat where famil like ('%" + Text1.Text + "%') or name like ('%" + Text1.Text + "%')"
Qeybat.Refresh
If Student.Recordset.RecordCount = 0 Then
Me.CHF.Value = 0
Else
Me.CHF.Value = 1

End If

DataGrid3.Refresh
End If
If Option5.Value = True Then  'Ã” ÊÃÊ œ— òœ ò·«” €Ì»  Â«
Qeybat.RecordSource = "select * from qeybat where clas like ('%" + Text1.Text + "%')"
Qeybat.Refresh
DataGrid3.Refresh
End If
If Option4.Value = True Then  '»— «”«” ‰«„ «” «œ
Qeybat.RecordSource = "select * from qeybat where Ostad like ('%" + Text1.Text + "%')"
Qeybat.Refresh
DataGrid3.Refresh
End If
Label15.Caption = Qeybat.Recordset.RecordCount
End If ' Ã” ÊÃÊ œ—





End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Text1_DblClick()
Text1.Text = ""

End Sub

Private Sub Text2_Change()
If Text2.Text = "gotre" Then
Call Command1_Click

End If

End Sub

Private Sub Text7_DblClick()
Text7.Text = ""
End Sub

