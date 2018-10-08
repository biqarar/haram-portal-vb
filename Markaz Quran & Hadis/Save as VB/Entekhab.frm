VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Entekhab 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„—ﬂ“ ﬁ—¬‰ Ê ÕœÌÀ ﬂ—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„"
   ClientHeight    =   4395
   ClientLeft      =   7215
   ClientTop       =   4995
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Entekhab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Palette         =   "Entekhab.frx":08CA
   ScaleHeight     =   4395
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command18 
      Caption         =   "À»  ‰«„ «Ê·ÌÂ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   600
      Width           =   1700
   End
   Begin VB.CommandButton Command17 
      Caption         =   "·Ì”  «‰ Ÿ«—"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton Command16 
      Caption         =   "„œÌ—Ì  „’«Õ»Â"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1800
      Width           =   1700
   End
   Begin VB.CommandButton Command15 
      Caption         =   "œÊ—Â Â«Ì ﬁ—«∆ "
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox useridtext 
      Height          =   735
      Left            =   9720
      TabIndex        =   26
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000002&
      Caption         =   "»Ì‘ —"
      Height          =   330
      Left            =   6360
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "«ÿ·«⁄«  ›«Ì· Œ—ÊÃÌ"
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«’·«Õ"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«’·«Õ"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adress:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   750
      End
      Begin VB.Label AdressLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F:\"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Network:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.Label NetAdresslabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F:\"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   840
         Width           =   225
      End
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
      Height          =   2415
      Left            =   9600
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   5175
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
         Connect         =   $"Entekhab.frx":D3F0
         OLEDBString     =   $"Entekhab.frx":D479
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
         Connect         =   $"Entekhab.frx":D502
         OLEDBString     =   $"Entekhab.frx":D58B
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
         Connect         =   $"Entekhab.frx":D614
         OLEDBString     =   $"Entekhab.frx":D69D
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
         Connect         =   $"Entekhab.frx":D726
         OLEDBString     =   $"Entekhab.frx":D7AF
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
         Connect         =   $"Entekhab.frx":D838
         OLEDBString     =   $"Entekhab.frx":D8C1
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
         Connect         =   $"Entekhab.frx":D94A
         OLEDBString     =   $"Entekhab.frx":D9D3
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
         Connect         =   $"Entekhab.frx":DA5C
         OLEDBString     =   $"Entekhab.frx":DAE5
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
         Connect         =   $"Entekhab.frx":DB6E
         OLEDBString     =   $"Entekhab.frx":DBF7
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
         Connect         =   $"Entekhab.frx":DC80
         OLEDBString     =   $"Entekhab.frx":DD09
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
         Left            =   2640
         Top             =   1800
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
         Connect         =   $"Entekhab.frx":DD92
         OLEDBString     =   $"Entekhab.frx":DE1B
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
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   4800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "êÊ«ÂÌ ‰«„Â"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   1700
   End
   Begin VB.CommandButton Command12 
      Caption         =   "¬„«—"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   6360
      TabIndex        =   15
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton Command10 
      Caption         =   "À»  ò·«” ÃœÌœ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1700
   End
   Begin VB.CommandButton Command9 
      Caption         =   "‰„«Ì‘ ‰„—« "
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton Command8 
      Caption         =   "œÊ—Â Â«Ì Õ›Ÿ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Å—œ«Œ  ÊœÌ⁄Â"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   1700
   End
   Begin ComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   4020
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "ò«—»— Ã«—Ì"
            TextSave        =   "ò«—»— Ã«—Ì"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3246
            MinWidth        =   3246
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "”«⁄  Ê  «—ÌŒ ›⁄·Ì"
            TextSave        =   "”«⁄  Ê  «—ÌŒ ›⁄·Ì"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FF80&
      Caption         =   "œ—»«—Â ‰—„ «›“«—"
      Height          =   490
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "„œÌ—Ì  ·Ì”  ò·«”Ì"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Œ—ÊÃ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ê÷⁄Ì  €Ì  Â«"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1700
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ÷Ê— Ê €Ì«»"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "À»  ﬁ—¬‰ ¬„Ê“"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1700
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À»  ‰«„"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8880
      TabIndex        =   32
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line23 
      X1              =   9120
      X2              =   9120
      Y1              =   600
      Y2              =   2040
   End
   Begin VB.Line Line22 
      X1              =   8760
      X2              =   9120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line21 
      X1              =   8640
      X2              =   9120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line20 
      X1              =   8640
      X2              =   9120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line19 
      X1              =   1500
      X2              =   1600
      Y1              =   650
      Y2              =   650
   End
   Begin VB.Line Line18 
      X1              =   1500
      X2              =   960
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line17 
      X1              =   1500
      X2              =   1000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line16 
      X1              =   1500
      X2              =   1500
      Y1              =   960
      Y2              =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰„—« "
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   28
      Top             =   450
      Width           =   510
   End
   Begin VB.Line Line4 
      X1              =   1800
      X2              =   2400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line15 
      X1              =   4440
      X2              =   4455
      Y1              =   2640
      Y2              =   2655
   End
   Begin VB.Line Line14 
      X1              =   4440
      X2              =   4320
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line13 
      X1              =   4440
      X2              =   4440
      Y1              =   2040
      Y2              =   2640
   End
   Begin VB.Line Line12 
      X1              =   4200
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line11 
      X1              =   4440
      X2              =   4440
      Y1              =   1680
      Y2              =   2040
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò·«” Â«"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   14
      Top             =   1320
      Width           =   705
   End
   Begin VB.Line Line10 
      X1              =   2280
      X2              =   2400
      Y1              =   650
      Y2              =   650
   End
   Begin VB.Line Line9 
      X1              =   1800
      X2              =   2400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line8 
      X1              =   2400
      X2              =   2400
      Y1              =   480
      Y2              =   2040
   End
   Begin VB.Line Line7 
      X1              =   6360
      X2              =   6840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line6 
      X1              =   6360
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line5 
      X1              =   6480
      X2              =   6840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   6840
      X2              =   6840
      Y1              =   600
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   4440
      X2              =   4200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   480
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«„ Õ«‰« "
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   13
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«„Ê— „«·Ì"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¬„Ê“‘"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      TabIndex        =   10
      Top             =   120
      Width           =   600
   End
   Begin VB.Menu mnuchuser 
      Caption         =   " €ÌÌ— ò«—»—"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuparvande 
      Caption         =   "Å—Ê‰œÂ"
      Begin VB.Menu FliePrrint 
         Caption         =   "›«Ì· Œ—ÊÃÌ"
         Begin VB.Menu Pc 
            Caption         =   "«“ Â«—œ"
         End
         Begin VB.Menu net 
            Caption         =   "«“ ‘»òÂ"
         End
      End
      Begin VB.Menu mnusetting 
         Caption         =   " ‰ŸÌ„« "
      End
      Begin VB.Menu mnuuserprofie 
         Caption         =   "ò«—»—"
         Begin VB.Menu PASSEDIT 
            Caption         =   " €ÌÌ— ò·„Â ⁄»Ê—"
         End
         Begin VB.Menu DSDSD 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu MNUMODIRIYATUSER 
            Caption         =   "„œÌ—Ì  ò«—»—«‰"
         End
      End
      Begin VB.Menu asdf 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuDsDore 
         Caption         =   "«’·«Õ œÊ—Â Â«"
      End
      Begin VB.Menu steww 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuend 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "Entekhab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
On Error Resume Next

If Entekhab.Height = 7000 Then
Entekhab.Height = 5000
Else
Entekhab.Height = 7000

End If



End Sub

Private Sub Command1_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-bankstudent-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount < 1 Then Exit Sub
14082513

BankStudent.Show
'Entekhab.Hide
'End If
End Sub

Private Sub Command10_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-mclass-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
ModiriyatCLASS.Show
'Entekhab.Hide

End Sub

Private Sub Command11_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-govahi-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Govahi.Show

End Sub

Private Sub Command12_Click()
Amar.Show

End Sub

Private Sub Command13_Click()
Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "netAdressXlsx" & "%') "
Setting.Refresh

'Me.AdressLabel.Caption = Setting.Recordset.Fields("xtext")

 Setting.Recordset.Fields("xtext") = InputBox("·ÿ›« ¬œ” ›Ê·œ— »—‰«„Â —« Ê«—œ ò‰Ìœ", "Adress")
Setting.Recordset.Update

Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "netAdressXlsx" & "%') "
Setting.Refresh

Me.NetAdresslabel.Caption = Setting.Recordset.Fields("xtext")



End Sub

Private Sub Command14_Click()
Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "AdressXlsx" & "%') "
Setting.Refresh

'Me.AdressLabel.Caption = Setting.Recordset.Fields("xtext")

 Setting.Recordset.Fields("xtext") = InputBox("·ÿ›« ¬œ” ›Ê·œ— »—‰«„Â —« Ê«—œ ò‰Ìœ", "Adress")
Setting.Recordset.Update

Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "AdressXlsx" & "%') "
Setting.Refresh

Me.AdressLabel.Caption = Setting.Recordset.Fields("xtext")



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Sub Command15_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-emtahan-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
sabt_nomre_omoomi.Show
End Sub

Private Sub Command16_Click()
mosahebe_settingf.Show

End Sub

Private Sub Command18_Click()

If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-entezar-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
first_paziresh.Show
End Sub

Private Sub Command2_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-qeybat-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513

QeybatF.Show


'Entekhab.Hide

End Sub

Private Sub Command3_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-gozaresh-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Gozaresh.Show
'Entekhab.Hide

End Sub

Private Sub Command4_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-fclass-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
FClassroom.Show
'Entekhab.Hide


End Sub

Private Sub Command5_Click()
Beep

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «“ ‰—„ «›“«— Œ«—Ã ‘ÊÌœ", vbQuestion + vbYesNo, "Œ—ÊÃ «“ ‰—„ «›“«—") = vbYes Then
End
End If

End Sub

Private Sub Command6_Click()
WE.Show

'MsgBox "»—‰«„Â ‰ÊÌ”: —÷« „ÕÌÿÌ øøøøøø›⁄·« Â„Ì‰  « »⁄œ«", vbInformation + vbOKOnly, "œ— »«—Â ‰—„ «›“«—"

End Sub

Private Sub Command7_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-vadie-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
VadieF.Show


End Sub

Private Sub Command8_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-emtahan-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
EmtahanF.Show

End Sub

Private Sub Command9_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "entekhab-karname-load" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Karname.Show

End Sub

Private Sub Form_DblClick()
CommonDialog1.ShowColor
Entekhab.BackColor = CommonDialog1.Color

End Sub

Private Sub Form_Load()

Me.SB.Panels(1).Text = user.OP.Text
Me.SB.Panels(3).Text = Taqvim.Tarikh.Caption



Dim ASD As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD, ParvandeQuranAmooZ As String

Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "AdressXlsx" & "%') "
Setting.Refresh

Me.AdressLabel.Caption = Setting.Recordset.Fields("xtext")

Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "netAdressXlsx" & "%') "
Setting.Refresh

Me.NetAdresslabel.Caption = Setting.Recordset.Fields("xtext")


On Error GoTo 1
GoTo 2
1:
net.Checked = True
Exit Sub
2:


Set oExcel = GetObject(Me.AdressLabel.Caption & "CopyofKarnameJadid.xlsx")
Pc.Checked = True






End Sub

Private Sub Form_Unload(Cancel As Integer)
Beep
If Me.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then End

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «“ ‰—„ «›“«— Œ«—Ã ‘ÊÌœ", vbQuestion + vbYesNo, "Œ—ÊÃ «“ ‰—„ «›“«—") = vbYes Then
End
Else
Cancel = 1
End If

End Sub

Private Sub Label5_Click()
Amar.Show

End Sub

Private Sub mnuexit_Click()
End

End Sub

Private Sub mnuozv1_Click()
ozviat.Show

End Sub

Private Sub mnuchuser_Click()

Entekhab.Hide
Asatid.Hide
BankStudent.Hide
Emtahan.Hide
FClassroom.Hide
Govahi.Hide
Gozaresh.Hide
Karname.Hide
ModiriyatCLASS.Hide
ozviat.Hide
VadieF.Hide
WE.Hide
QeybatF.Hide
user.Show
'Me.SB.Panels(1).Text = "„ÌÂ„«‰"

End Sub

Private Sub mnuDsDore_Click()

If Me.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then
DsDore.Show
Exit Sub

End If

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "dsdore-edit" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub


DsDore.Show

End Sub

Private Sub mnuend_Click()
End

End Sub

Private Sub MNUMODIRIYATUSER_Click()
If Me.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then UserProfile.Show

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "modiriyat" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub

UserProfile.Show
End Sub

Private Sub mnusetting_Click()
If Me.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then
SettingF.Show
Exit Sub
Else
MsgBox "809-25-31-4", vbExclamation + vbOKOnly, ""
End If

End Sub

Private Sub net_Click()
Pc.Checked = False
net.Checked = True
End Sub

Private Sub PASSEDIT_Click()
ozviat.Show
ozviat.Text1.Text = Me.useridtext.Text


End Sub

Private Sub pc_Click()
Pc.Checked = True
net.Checked = False
End Sub

Private Sub Timer1_Timer()
SB.Panels(4).Text = Format(Now, "long time")

End Sub


