VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form VadieF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "›—„ Å— œ«Œ  ÊœÌ⁄Â"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Vadie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Motor"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      Begin MSAdodcLib.Adodc vadie 
         Height          =   330
         Left            =   360
         Top             =   3360
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
         Connect         =   $"Vadie.frx":08CA
         OLEDBString     =   $"Vadie.frx":0953
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from vadie"
         Caption         =   "vadie"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Homa"
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
         Left            =   360
         Top             =   3000
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
         Connect         =   $"Vadie.frx":09DC
         OLEDBString     =   $"Vadie.frx":0A65
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
         Top             =   2160
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
         Connect         =   $"Vadie.frx":0AEE
         OLEDBString     =   $"Vadie.frx":0B77
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
         Top             =   1680
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
         Connect         =   $"Vadie.frx":0C00
         OLEDBString     =   $"Vadie.frx":0C89
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
         Top             =   1200
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
         Connect         =   $"Vadie.frx":0D12
         OLEDBString     =   $"Vadie.frx":0D9B
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
         Connect         =   $"Vadie.frx":0E24
         OLEDBString     =   $"Vadie.frx":0EAD
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
         Connect         =   $"Vadie.frx":0F36
         OLEDBString     =   $"Vadie.frx":0FBF
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
         Left            =   360
         Top             =   2520
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
         Connect         =   $"Vadie.frx":1048
         OLEDBString     =   $"Vadie.frx":10D1
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
      Begin MSAdodcLib.Adodc SettingUser 
         Height          =   330
         Left            =   3360
         Top             =   840
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
         Connect         =   $"Vadie.frx":115A
         OLEDBString     =   $"Vadie.frx":11E3
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from SettingUser"
         Caption         =   "SettingUser"
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
      Begin MSAdodcLib.Adodc Setting 
         Height          =   330
         Left            =   2880
         Top             =   1560
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
         Connect         =   $"Vadie.frx":126C
         OLEDBString     =   $"Vadie.frx":12F5
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
         Left            =   3240
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
         Connect         =   $"Vadie.frx":137E
         OLEDBString     =   $"Vadie.frx":1407
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Top             =   360
      Width           =   4815
   End
   Begin VB.Frame Frame4 
      Caption         =   "¬„«— ÊœÌ⁄Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   120
      Width           =   3615
      Begin VB.TextBox text_auto_sabt 
         BackColor       =   &H000000FF&
         Height          =   420
         Left            =   240
         TabIndex        =   82
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin ComctlLib.ProgressBar pb1 
         Height          =   135
         Left            =   240
         TabIndex        =   81
         Top             =   3480
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "»Â —Ê“ —”«‰Ì"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "À»  «—Ã«⁄ ÊœÌ⁄Â"
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label labelmojoodi 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         TabIndex        =   79
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label labelnoerja 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         TabIndex        =   78
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label labelerja 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         TabIndex        =   77
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ò· Ê—ÊœÌ"
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
         TabIndex        =   76
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "„ÊÃÊœÌ ›⁄·Ì"
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
         TabIndex        =   75
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label labelkol 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         TabIndex        =   74
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "⁄œ„ «” —œ«œ"
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
         TabIndex        =   73
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "«—Ã«⁄ œ«œÂ ‘œÂ"
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
         TabIndex        =   72
         Top             =   720
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "«—Ã«⁄ ÊœÌ⁄Â"
      Height          =   1575
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   2520
      Width           =   5895
      Begin VB.CommandButton Command8 
         DisabledPicture =   "Vadie.frx":1490
         DownPicture     =   "Vadie.frx":2610A
         DragIcon        =   "Vadie.frx":4AD84
         Height          =   330
         Left            =   120
         Picture         =   "Vadie.frx":6F9FE
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«—Ã«⁄ ÊœÌ⁄Â"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "À»  «—Ã«⁄ ÊœÌ⁄Â"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "⁄œ„ «—Ã«⁄"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "ÊœÌ⁄Â „” —œ ‰„Ì ê—œœ"
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox Combo6 
         Height          =   420
         ItemData        =   "Vadie.frx":94678
         Left            =   3000
         List            =   "Vadie.frx":94685
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Vadie.frx":946AD
         Left            =   1560
         List            =   "Vadie.frx":946B7
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Text            =   "Ê÷⁄Ì  «—Ã«⁄"
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo8 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Vadie.frx":946CC
         Left            =   120
         List            =   "Vadie.frx":946CE
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Text            =   "„»·€ «—Ã«⁄"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   840
         TabIndex        =   85
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ À» "
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
         Left            =   2040
         TabIndex        =   84
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "⁄·  «—Ã«⁄"
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
         Left            =   5040
         TabIndex        =   69
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "„»·€ «—Ã«⁄"
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
         Left            =   840
         TabIndex        =   68
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷⁄Ì  «—Ã«⁄"
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
         Left            =   1920
         TabIndex        =   67
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷⁄Ì "
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
         Left            =   5040
         TabIndex        =   66
         Top             =   1200
         Width           =   510
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Vadie.frx":946D0
      Height          =   3975
      Left            =   120
      TabIndex        =   57
      Top             =   4200
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
   Begin VB.TextBox Text3 
      Height          =   420
      Left            =   10800
      TabIndex        =   54
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   420
      Left            =   5160
      TabIndex        =   53
      Text            =   " Ê÷ÌÕ« "
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "‰„«Ì‘"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "‰„«Ì‘ ÊœÌ⁄Â Â«Ì Å—œ«Œ Ì  Ê”ÿ ﬁ—¬‰ ¬„Ê“"
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Vadie.frx":946E6
      Left            =   5040
      List            =   "Vadie.frx":946F0
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Text            =   "Å—œ«Œ  ò«„·"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ç«Å —”Ìœ"
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   1560
      Width           =   975
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
      ItemData        =   "Vadie.frx":94709
      Left            =   7080
      List            =   "Vadie.frx":9470B
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "«‰ Œ«» ò‰Ìœ"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "„‘ —ò"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "«‰ Œ«» ò‰Ìœ"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Vadie.frx":9470D
      Left            =   7080
      List            =   "Vadie.frx":94720
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "ÊœÌ⁄Â"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "À»  "
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Vadie.frx":9474A
      Height          =   3735
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12640511
      DefColWidth     =   80
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
      Caption         =   "„‘Œ’«  ÊœÌ⁄Â"
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "KOD"
         Caption         =   "‘„«—Â ÅÌê—Ì"
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
         DataField       =   "Parvande"
         Caption         =   "Å—Ê‰œÂ"
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
         DataField       =   "Mablaq"
         Caption         =   "„»·€ ÊœÌ⁄Â"
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
      BeginProperty Column04 
         DataField       =   "Op"
         Caption         =   "œ—Ì«›  ò‰‰œÂ"
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
         DataField       =   "Dore"
         Caption         =   "œÊ—Â"
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
         DataField       =   "Vazeyat"
         Caption         =   "Ê÷⁄Ì  Å—œ«Œ "
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
         DataField       =   "Terja"
         Caption         =   " —«—ÌŒ «—Ã«⁄"
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
         DataField       =   "Elat"
         Caption         =   "⁄·  «—Ã«⁄"
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
      BeginProperty Column10 
         DataField       =   "TTasvie"
         Caption         =   " «—ÌŒ  ”ÊÌÂ"
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
         DataField       =   "Bedehkar"
         Caption         =   "»œÂò«—"
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
         DataField       =   "MErja"
         Caption         =   "„»·€ «—Ã«⁄"
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
         DataField       =   "VErja"
         Caption         =   "Ê÷⁄Ì  «—Ã«⁄"
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
         DataField       =   "VarizHaram"
         Caption         =   "VarizHaram"
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
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   59
      Top             =   8175
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "ò«—»— Ã«—Ì"
            TextSave        =   "ò«—»— Ã«—Ì"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   " «—ÌŒ «„—Ê“"
            TextSave        =   " «—ÌŒ «„—Ê“"
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
   Begin VB.Frame Frame1 
      Caption         =   "„‘Œ’«  ÊœÌ⁄Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label Label50 
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
         TabIndex        =   87
         Top             =   3600
         Width           =   585
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tozih"
         DataSource      =   "vadie"
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
         TabIndex        =   86
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Terja"
         DataSource      =   "vadie"
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
         TabIndex        =   50
         Top             =   3240
         Width           =   135
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «—Ã«⁄"
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
         Top             =   3240
         Width           =   780
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Elat"
         DataSource      =   "vadie"
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
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "⁄·  «—Ã«⁄"
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
         TabIndex        =   47
         Top             =   2880
         Width           =   705
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ À» "
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
         TabIndex        =   40
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "D"
         DataSource      =   "vadie"
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
         TabIndex        =   39
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷⁄Ì "
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
         TabIndex        =   38
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Vazeyat"
         DataSource      =   "vadie"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   315
         Left            =   480
         TabIndex        =   37
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Op"
         DataSource      =   "vadie"
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
         TabIndex        =   36
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Dore"
         DataSource      =   "vadie"
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
         TabIndex        =   35
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Mablaq"
         DataSource      =   "vadie"
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
         TabIndex        =   34
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Parvande"
         DataSource      =   "vadie"
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
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "KOD"
         DataSource      =   "vadie"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   480
         TabIndex        =   32
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "À»  ‰«„ œ—  œÊ—Â"
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
         TabIndex        =   31
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   " ÕÊÌ· »Â ¬ﬁ«Ì"
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
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "„»·€ œ—Ì«› Ì"
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
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label22 
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
         Left            =   2280
         TabIndex        =   28
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â ÅÌêÌ—Ì"
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
         Index           =   0
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
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
      Height          =   4095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   3615
      Begin VB.Label Label43 
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
         TabIndex        =   52
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label42 
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
         TabIndex        =   51
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label29 
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
         TabIndex        =   46
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label28 
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
         TabIndex        =   45
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label17 
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
         TabIndex        =   44
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label14 
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
         TabIndex        =   43
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         TabIndex        =   42
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ Å—œ«Œ  "
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
         TabIndex        =   41
         Top             =   2880
         Width           =   945
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
         TabIndex        =   20
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label20 
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
         TabIndex        =   19
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label19 
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
         TabIndex        =   18
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label13 
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
         TabIndex        =   17
         Top             =   360
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
         TabIndex        =   16
         Top             =   720
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
         TabIndex        =   15
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label16 
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
         TabIndex        =   14
         Top             =   1440
         Width           =   585
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
         TabIndex        =   13
         Top             =   1440
         Width           =   135
      End
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
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
      Left            =   6120
      TabIndex        =   58
      Top             =   2760
      Width           =   45
   End
   Begin VB.Label Label45 
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
      Left            =   5040
      TabIndex        =   56
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   6240
      TabIndex        =   55
      Top             =   0
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ÃÂ  À»  ‰«„ œ— œÊ—Â"
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
      Left            =   8520
      TabIndex        =   11
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ÿ—Õ"
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
      Left            =   11760
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "„»·€ œ—Ì«› Ì"
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
      Left            =   9000
      TabIndex        =   5
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   " ÕÊÌ· »Â ¬ﬁ«Ì"
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
      Left            =   8880
      TabIndex        =   4
      Top             =   2160
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "—Ì«·"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ã” ÊÃÊ œ— ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
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
      Left            =   7920
      TabIndex        =   2
      Top             =   0
      Width           =   1845
   End
   Begin VB.Menu mnuhno 
      Caption         =   "#"
   End
   Begin VB.Menu mnufail 
      Caption         =   "Å—Ê‰œÂ"
      Begin VB.Menu changevadie 
         Caption         =   " €ÌÌ— „‘Œ’«  ÊœÌ⁄Â"
      End
      Begin VB.Menu mnudell 
         Caption         =   "Õ–› Å—œ«Œ  ÊœÌ⁄Â"
      End
   End
End
Attribute VB_Name = "VadieF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub changevadie_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "vadie-edit" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
If DataGrid1.AllowUpdate = False Then
changevadie.Checked = True

DataGrid1.AllowUpdate = True

Else



changevadie.Checked = False

DataGrid1.AllowUpdate = False

End If



End Sub

Private Sub Combo1_Change()



Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where NAME like ('%" + Combo3.Text + "%')" ' „«„ Å—Ê‰œÂ Â« —« „Ì ê—œœ Ê ò”«‰Ì òÂ ÿ— Õ «‰ Œ«» ‘œÂ —« œ«—‰œ „Ì ¬Ê—œÅ
Tarhha.Refresh
'On Error Resume Next
'Combo1.Text = Tarhha.Recordset.Fields("tozih")
If Combo1.Text < Tarhha.Recordset.Fields("tozih") Then
Combo5.Text = "»œÂò«—"
Else
Combo5.Text = "Å—œ«Œ  ò«„·"
End If


If Combo1.Text > Tarhha.Recordset.Fields("tozih") Then
Combo5.Text = "»” «‰ò«—"

End If






End Sub

Private Sub Combo1_Click()


Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where NAME like ('%" + Combo3.Text + "%')" ' „«„ Å—Ê‰œÂ Â« —« „Ì ê—œœ Ê ò”«‰Ì òÂ ÿ— Õ «‰ Œ«» ‘œÂ —« œ«—‰œ „Ì ¬Ê—œÅ
Tarhha.Refresh
'On Error Resume Next
'Combo1.Text = Tarhha.Recordset.Fields("tozih")
If Val(Combo1.Text) < Val(Tarhha.Recordset.Fields("tozih")) Then
Combo5.Text = "»œÂò«—"
Else
Combo5.Text = "Å—œ«Œ  ò«„·"
End If


If Val(Combo1.Text) > Val(Tarhha.Recordset.Fields("tozih")) Then
Combo5.Text = "»” «‰ò«—"

End If



End Sub

Private Sub Combo3_Change()
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where NAME like ('%" + Combo3.Text + "%')" ' „«„ Å—Ê‰œÂ Â« —« „Ì ê—œœ Ê ò”«‰Ì òÂ ÿ— Õ «‰ Œ«» ‘œÂ —« œ«—‰œ „Ì ¬Ê—œÅ
Tarhha.Refresh
On Error Resume Next
Combo1.Text = Tarhha.Recordset.Fields("tozih")

End Sub

Private Sub Combo3_Click()
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where NAME like ('%" + Combo3.Text + "%')" ' „«„ Å—Ê‰œÂ Â« —« „Ì ê—œœ Ê ò”«‰Ì òÂ ÿ— Õ «‰ Œ«» ‘œÂ —« œ«—‰œ „Ì ¬Ê—œÅ
Tarhha.Refresh
On Error Resume Next
Combo1.Text = Tarhha.Recordset.Fields("tozih")

Exit Sub

If Combo3.Text = "’Ê  Ê ·Õ‰" Then Text3.Text = "200,000"
If Combo3.Text = " ›”Ì—" Then Text3.Text = "200,000"
If Combo3.Text = "—Ê ŒÊ«‰Ì" Then Text3.Text = "100,000"
If Combo3.Text = "—Ê«‰ ŒÊ«‰Ì" Then Text3.Text = "100,000"
If Combo3.Text = " ÃÊÌœ" Then Text3.Text = "100,000"
If Combo3.Text = " ÃÊÌœ ”ÿÕ2" Then Text3.Text = "100,000"
If Combo3.Text = "—Ê ŒÊ«‰Ì Ê —Ê«‰ ŒÊ«‰Ì" Then Text3.Text = "100,000"
If Combo3.Text = " —Ã„Â" Then Text3.Text = "100,000"
If Combo3.Text = " ›”Ì—" Then Text3.Text = "200,000"
If Combo3.Text = "œ—”‰«„Â" Then Text3.Text = "200,000"






End Sub

Private Sub Combo4_Change()
'Combo3.Text = "„‘ —ò"
On Error GoTo 10
Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where goroh like ('" & Combo4.Text & "')"
Tarhha.Refresh

Tarhha.Recordset.Sort = "sortname"

For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("name"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)


10 Exit Sub


If Combo4.Text = "⁄„Ê„Ì" Then
Combo3.Clear

Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "1" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)

Text3.Text = "100,000"


End If


If Combo4.Text = "ò«—ê«Â Â«" Then
Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "3" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)

Text3.Text = "100,000"


End If

If Combo4.Text = " —»Ì  „—»Ì" Then
Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "4" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)

Text3.Text = "100,000"


End If


If Combo4.Text = "Õ›Ÿ ﬁ—¬‰ ò—Ì„" Then

Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "2" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)

Text3.Text = "200,000"


End If

If Combo4.Text = "„ÃÂÊ·" Then
Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "0" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)
End If

End Sub

Private Sub Combo4_Click()



On Error GoTo 10
Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where goroh like ('" & Combo4.Text & "')"
Tarhha.Refresh

Tarhha.Recordset.Sort = "sortname"

For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("name"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)


10 Exit Sub

If Combo4.Text = "⁄„Ê„Ì" Then
Combo3.Clear

Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "1" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)
Text3.Text = "100,000"
End If


If Combo4.Text = "ò«—ê«Â Â«" Then
Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "3" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)
Text3.Text = "100,000"
End If

If Combo4.Text = " —»Ì  „—»Ì" Then
Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "4" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)
Text3.Text = "100,000"
End If


If Combo4.Text = "Õ›Ÿ ﬁ—¬‰ ò—Ì„" Then

Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "2" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)
Text3.Text = "200,000"
End If

If Combo4.Text = "„ÃÂÊ·" Then
Combo3.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "0" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo3.Text = Combo3.List(0)
Text3.Text = "200,000"
End If
End Sub


Private Sub Command1_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "vadie-newsabt" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
'Dim VC As Double
Dim DDD As String
Dim ASD As String
Dim vc As String


'On Error Resume Next








If Combo2.Text = "«‰ Œ«» ò‰Ìœ" Then
MsgBox "‰«„ œ—Ì«›  ò‰‰œÂ —« «‰ Œ«» ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If



If Combo1.Text = "" Then
MsgBox "„»·€ œ—Ì«› Ì —« «‰ Œ«» ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If

If Combo3.Text = "" Then
'MsgBox "œÊ—Â „Ê—œ ‰Ÿ— —« «‰ Œ«» ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
'Exit Sub
Combo3.Text = "„‘ —ò"
Combo4.Text = ""

End If

Beep


If MsgBox("  „»·€  " & Combo1.Text & "  »—«Ì ¬ﬁ«Ì  " & Label10.Caption & "    À»  ŒÊ«Âœ ‘œ    ¬Ì« „ÿ„∆‰ Â” Ìœ", vbQuestion + vbYesNo, "") = vbYes Then
GoTo 14
Else
Exit Sub
End If
14:


DDD = Combo4.Text & " - " & Combo3.Text



vadie.Refresh
vadie.RecordSource = "select * from vadie where parvande like ('%" + Label13.Caption + "%')   and  dore like ('%" & DDD & "%')"
vadie.Refresh



If vadie.Recordset.BOF = True Or vadie.Recordset.EOF = True Then
GoTo 18
Else
If MsgBox("ﬁ»·« »—«Ì «Ì‰ ﬁ—¬‰ ¬„Ê“ œ— «Ì‰ œÊ—Â ÊœÌ⁄Â À»  ‘œÂ «” " & Chr(10) & "¬Ì« „Ì ŒÊ«ÂÌœ ÊœÌ⁄Â ÃœÌœ À»  ‘Êœ", vbQuestion + vbYesNo, "À»  ÊœÌ⁄Â") = vbYes Then
GoTo 18
Else

Exit Sub
End If
End If
18:


vadie.Refresh
vadie.RecordSource = "select * from vadie where parvande like ('%" + "" + "%') "
vadie.Refresh



Dim SaljariSTR, CodeBakhsh As String


SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "SalJari" + "%')"
SettingUser.Refresh

SaljariSTR = SettingUser.Recordset.Fields("value")

SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "CodeBakhsh" + "%')"
SettingUser.Refresh
CodeBakhsh = SettingUser.Recordset.Fields("value")



vc = Val("14082513" & SaljariSTR & CodeBakhsh & "001")


'vadie.Refresh
'VC = Val(vadie.Recordset.Fields("kod"))
'vadie.Recordset.MoveNext
'.>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'vadie.Refresh

'For I = 1 To vadie.Recordset.RecordCount
 


'If Val(vadie.Recordset.Fields("kod")) > VC Then
'VC = Val(vadie.Recordset.Fields("kod"))
'End If

'vadie.Recordset.MoveNext

'Next I


'VC = VC + 1

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'If Combo5.Text = "»Âœò«—" And Text2.Text = " «—ÌŒ „—«ÃÂ ÃÂ  Å—œ«Œ  »œÂÌ" Or Text2.Text = "" Then
'MsgBox " «—ÌŒ „—«Ã⁄Â ÃÂ   ’ÊÌÂ Õ”«» —« „‘Œ’ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
'Exit Sub
'Else

'GoTo 98

'End If

vc = Label13.Caption & Val(Label6.Caption) & Int(Rnd(100) * 100)
166:
vadie.Refresh
vadie.RecordSource = "select * from vadie where kod like ('%" & vc & "%') "
vadie.Refresh
If vadie.Recordset.BOF = True Or vadie.Recordset.EOF = True Then
GoTo 155
Else
vc = Label13.Caption & Val(Label6.Caption) & Int(Rnd(100) * 100)
GoTo 166

End If


155:
vadie.Refresh
vadie.Recordset.AddNew
vadie.Recordset.Fields("KOD") = vc ' Label13.Caption & Val(Label6.Caption) & Int(Rnd(100) * 100) '& Combo1.Text  ' VC
vadie.Recordset.Fields("PARVANDE") = Label13.Caption
vadie.Recordset.Fields("MABLAQ") = Combo1.Text
vadie.Recordset.Fields("D") = Label6.Caption
vadie.Recordset.Fields("op") = Combo2.Text
vadie.Recordset.Fields("DORE") = Combo4.Text & " - " & Combo3.Text
vadie.Recordset.Fields("VAZEYAT") = Combo5.Text
vadie.Recordset.Fields("tozih") = Text2.Text
vadie.Recordset.Update
vadie.Refresh
Beep
98



If MsgBox("Å—œ«Œ  ÊœÌ⁄Â À»  ‘œ" & Chr(10) & "òœ ÅÌê—Ì" & Chr(10) & vc & Chr(10) & "¬Ì« —”Ìœ „Ì ŒÊ«ÂÌœ", vbInformation + vbYesNo, "Å—œ«Œ  ÊœÌ⁄Â") = vbYes Then
GoTo 17
Else
Exit Sub
End If
17:

'vc = Label13.Caption & Label6.Caption & Combo1.Text
'vc = Label13.Caption & Val(Label6.Caption) & Int(Rnd(100) * 100)

vadie.Refresh
vadie.RecordSource = "select * from vadie where kod like ('%" & vc & "%') "
vadie.Refresh
'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
Dim AD As String
On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:



If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "vadiexls.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "vadiexls.xlsx")
End If


'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\vadiexls.xlsx")
'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("b3").Value = vadie.Recordset.Fields("kod")
oExcel.ActiveSheet.Range("f3").Value = vadie.Recordset.Fields("kod")


oExcel.ActiveSheet.Range("b4").Value = vadie.Recordset.Fields("parvande")
oExcel.ActiveSheet.Range("f4").Value = vadie.Recordset.Fields("parvande")




oExcel.ActiveSheet.Range("b6").Value = vadie.Recordset.Fields("mablaq") & " " & "—Ì«·"
oExcel.ActiveSheet.Range("f6").Value = vadie.Recordset.Fields("mablaq") & " " & "—Ì«·"


oExcel.ActiveSheet.Range("b7").Value = vadie.Recordset.Fields("op")
oExcel.ActiveSheet.Range("f7").Value = vadie.Recordset.Fields("op")


oExcel.ActiveSheet.Range("b8").Value = vadie.Recordset.Fields("d")
oExcel.ActiveSheet.Range("f8").Value = vadie.Recordset.Fields("d")


ASD = vadie.Recordset.Fields("parvande")
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + ASD + "%')"
Student.Refresh


oExcel.ActiveSheet.Range("b5").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("f5").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")


'MsgBox "—”Ìœ ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "ç«Å —”Ìœ"



'oExcel.SaveAs VC
'oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True

oExcel.Application.Visible = True
On Error GoTo 722


oExcel.Parent.Windows(2).Visible = True
GoTo 910
722:

oExcel.Parent.Windows(1).Visible = True
910:
''''''

oExcel.SaveAs vc
'oExcel.Close
'
'

Call Command7_Click



End Sub

Private Sub Command2_Click()
Dim ASD As String
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "vadie-print" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
vadie.Refresh
vadie.RecordSource = "select * from vadie where kod like ('%" + Label26.Caption + "%')  "
vadie.Refresh


'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:

If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "vadiexls.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "vadiexls.xlsx")
End If





Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "vadiexls.xlsx")
'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("b3").Value = vadie.Recordset.Fields("kod")
oExcel.ActiveSheet.Range("f3").Value = vadie.Recordset.Fields("kod")


oExcel.ActiveSheet.Range("b4").Value = vadie.Recordset.Fields("parvande")
oExcel.ActiveSheet.Range("f4").Value = vadie.Recordset.Fields("parvande")




oExcel.ActiveSheet.Range("b6").Value = vadie.Recordset.Fields("mablaq") & " " & "—Ì«·"
oExcel.ActiveSheet.Range("f6").Value = vadie.Recordset.Fields("mablaq") & " " & "—Ì«·"


oExcel.ActiveSheet.Range("b7").Value = vadie.Recordset.Fields("op")
oExcel.ActiveSheet.Range("f7").Value = vadie.Recordset.Fields("op")


oExcel.ActiveSheet.Range("b8").Value = vadie.Recordset.Fields("d")
oExcel.ActiveSheet.Range("f8").Value = vadie.Recordset.Fields("d")

ASD = vadie.Recordset.Fields("parvande")
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + ASD + "%')"
Student.Refresh

oExcel.ActiveSheet.Range("b5").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("f5").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")


MsgBox "—”Ìœ ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "ç«Å —”Ìœ"
X = vadie.Recordset.Fields("kod")

oExcel.SaveAs X
oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True


End Sub

Private Sub Command3_Click()
Call Command4_Click



vadie.Refresh
vadie.RecordSource = "select * from vadie where  parvande like ('%" + Label13.Caption + "%') "
vadie.Refresh




















End Sub

Private Sub Command4_Click()
If DataGrid2.Visible = False Then
DataGrid2.Visible = True
DataGrid1.Visible = False
Frame1.Visible = False
Frame2.Visible = True
Else
DataGrid2.Visible = False
DataGrid1.Visible = True
Frame1.Visible = True
Frame2.Visible = False

End If
vadie.Refresh
vadie.RecordSource = "select * from vadie where kod like ('%" + "" + "%') "
vadie.Refresh
vadie.Recordset.Sort = "merja"

End Sub

Private Sub Command5_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "vadie-erja" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
If Frame1.Visible = False Then

MsgBox "«» œ« »«Ìœ œò„Â ‰„«Ì‘ Ì« ÊœÌ⁄Â —« »“‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If
If Combo6.Text = "«‰ Œ«» ò‰Ìœ" Then
MsgBox "⁄·  «—Ã«⁄ —« «‰ Œ«» ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If

Beep


If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ ÊœÌ⁄Â «—Ã«⁄ œ«œÂ ‘Êœ", vbQuestion + vbYesNo, "«—Ã«⁄ ÊœÌ⁄Â") = vbYes Then

' vadie.Refresh
'vadie.RecordSource = "select * from vadie where kod like ('%" + Label26.Caption + "%') "
'vadie.Refresh
vadie.Recordset.Fields("terja") = Label6.Caption
vadie.Recordset.Fields("elat") = Combo6.Text

    vadie.Recordset.Fields("verja") = Combo7.Text
   vadie.Recordset.Fields("merja") = Combo8.Text
  

vadie.Recordset.Fields("vazeyat") = "«—Ã«⁄ ÊœÌ⁄Â"
vadie.Recordset.Update
vadie.Refresh




Call Command7_Click

End If

End Sub


Private Sub Command6_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "vadie-no-erja" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
If Frame1.Visible = False Then
MsgBox "«» œ« »«Ìœ œò„Â ‰„«Ì‘ Ì« ÊœÌ⁄Â —« »“‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If
Beep
If Combo6.Text = "«‰ Œ«» ò‰Ìœ" Then
MsgBox "⁄·  ⁄œ„ «—Ã«⁄ —« «‰ Œ«» ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If


If MsgBox("ÊœÌ⁄Â „” —œ ‰„Ì ê—œœ   ¬Ì« „ÿ„∆‰ Â” Ìœ", vbQuestion + vbYesNo, "⁄œ„ «—Ã«⁄ ÊœÌ⁄Â") = vbYes Then

' vadie.Refresh
'vadie.RecordSource = "select * from vadie where kod like ('%" + Label26.Caption + "%') "
'vadie.Refresh
vadie.Recordset.Fields("terja") = Label6.Caption
vadie.Recordset.Fields("elat") = Combo6.Text
vadie.Recordset.Fields("vazeyat") = "ÊœÌ⁄Â „” —œ ‰„Ì ê—œœ"
vadie.Recordset.Update
vadie.Refresh


Call Command7_Click
End If




End Sub


Private Sub Command7_Click()
Exit Sub

If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "vadie-amar" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513

vadie.Refresh
vadie.RecordSource = "select * from vadie where kod like ('%" + "" + "%')" ' or parvande like ('%" + Text1.Text + "%')or vazeyat like ('%" + Text1.Text + "%')or tozih like ('%" + Text1.Text + "%') "
vadie.Refresh
Dim Kolvorodi, Erja, Noerja, Feli As Double
 Kolvorodi = 0
 Erja = 0
 Noerja = 0
 Feli = 0
 PB1.Visible = True
 
PB1.Max = vadie.Recordset.RecordCount

For I = 1 To vadie.Recordset.RecordCount
Kolvorodi = Kolvorodi + Val(vadie.Recordset.Fields("mablaq"))
 
 vadie.Recordset.MoveNext
 PB1.Value = PB1.Value + 1
 
 Next I
 
labelkol.Caption = Kolvorodi & "00"

PB1.Value = 0

vadie.Refresh
vadie.RecordSource = "select * from vadie where vazeyat like ('%" + "«—Ã«⁄ ÊœÌ⁄Â" + "%')" ' or parvande like ('%" + Text1.Text + "%')or vazeyat like ('%" + Text1.Text + "%')or tozih like ('%" + Text1.Text + "%') "
vadie.Refresh

PB1.Max = vadie.Recordset.RecordCount


For I = 1 To vadie.Recordset.RecordCount
If vadie.Recordset.Fields("merja") <> "" Then GoTo 12

GoTo 13
12

Erja = Erja + Val(vadie.Recordset.Fields("merja"))
13

PB1.Value = PB1.Value + 1

 vadie.Recordset.MoveNext
 
 Next I
 
labelerja.Caption = Erja

PB1.Value = 0


vadie.Refresh
vadie.RecordSource = "select * from vadie where vazeyat like ('%" + "ÊœÌ⁄Â „” —œ ‰„Ì ê—œœ" + "%')"  ' or parvande like ('%" + Text1.Text + "%')or vazeyat like ('%" + Text1.Text + "%')or tozih like ('%" + Text1.Text + "%') "
vadie.Refresh
PB1.Max = vadie.Recordset.RecordCount

For I = 1 To vadie.Recordset.RecordCount
Noerja = Noerja + Val(vadie.Recordset.Fields("mablaq"))
 
 vadie.Recordset.MoveNext
 PB1.Value = PB1.Value + 1
 
 Next I
 labelnoerja.Caption = Noerja & "00"
 
labelmojoodi.Caption = Val(labelkol.Caption) - Val(labelerja.Caption)
PB1.Value = 0
PB1.Visible = False


End Sub

Private Sub Command8_Click()
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:

If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "kolvadie.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "kolvadie.xlsx")
End If



vadie.Refresh
vadie.RecordSource = "select * from vadie where kod like ('%" & "" & "%')"
vadie.Refresh

For I = T To vadie.Recordset.RecordCount


'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("a" & I + 1).Value = vadie.Recordset.Fields("mablaq")
oExcel.ActiveSheet.Range("b" & I + 1).Value = vadie.Recordset.Fields("merja")
oExcel.ActiveSheet.Range("c" & I + 1).Value = vadie.Recordset.Fields("vazeyat")

vadie.Recordset.MoveNext

Next I

oExcel.SaveAs "hhjhjgj"
oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True


End Sub

Private Sub DataGrid1_Click()
On Error Resume Next

vadie.Recordset.Update


End Sub

Private Sub DataGrid2_DblClick()

Dim A As String

On Error GoTo 7
GoTo 8
7:
MsgBox "›«Ì· ÅÌœ« ‰‘œ", vbCritical + vbOKOnly, "Œÿ«"
'Scan.Hide

Exit Sub

8:

If Entekhab.Pc.Checked = True Then
A = Student.Recordset.Fields("PARVANDE")

Scan.Text1.Text = A

Scan.Show
A = SettingF.ScanAdress.Caption & A & "\" & A & ".jpg"
'A = Student.Recordset.Fields("scan")
Scan.Im1.Picture = LoadPicture(A)

Exit Sub
End If

If Entekhab.net.Checked = True Then
A = Student.Recordset.Fields("PARVANDE")

Scan.Text1.Text = A

Scan.Show
'\\Yafatemeh2-pc\f\Markaz Quran & Hadis\FormScan\Pic\9020204\9020204.jpg
A = SettingF.NetScanAdress.Caption & A & "\" & A & ".jpg"
'A = Student.Recordset.Fields("scan")
Scan.Im1.Picture = LoadPicture(A)
Exit Sub
End If

Scan.Text1.Text = A


End Sub

Private Sub Form_Load()
Label6.Caption = Taqvim.Tarikh.Caption

'Combo2.AddItem ("„Õ„œ »«ﬁ— »«ﬁ—Ì")
'Combo2.AddItem ("„Õ„œ —”Ê·Ì")
'Combo2.AddItem ("Ã„«· «·œÌ‰ Õ”‰Ì")
'Combo2.AddItem ("⁄·Ì ‰Ê—Ê“Ì")
'Combo2.AddItem ("⁄·Ì—÷« «Ì—«‰ ‰é«œ")
'Combo2.AddItem ("«”„«⁄Ì· „—«œŒ«‰Ì")

'Combo2.AddItem ("")
'Combo2.AddItem ("")
'Combo2.AddItem ("")
'œ— «Ì‰ ﬁ”„  „»·€ »— ê—œ«‰œÂ ‘œ‰ »Â ﬁ—¬‰ ¬„Ê“ —Ê« Ê«—œ ò„Ì ò„œ

For I = 1000 To 20000 Step 1000

Combo8.AddItem (I)
Next I


Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "User-VadieF-DaryaftKonnande" & "%') "
Setting.Refresh
Combo2.Clear

 For I = 1 To Setting.Recordset.RecordCount
 Combo2.AddItem (Setting.Recordset.Fields("xtext"))
Setting.Recordset.MoveNext
Next I



Me.stb1.Panels(1).Text = user.OP.Text
Me.stb1.Panels(3).Text = Taqvim.Label1.Caption





Tarhha.Refresh
Tarhha.Recordset.Sort = "sortgoroh"


For I = 1 To Tarhha.Recordset.RecordCount
Combo4.AddItem (Tarhha.Recordset.Fields("goroh"))
    xsort = Tarhha.Recordset.Fields("sortgoroh")
     On Error GoTo 10
     
     While xsort = Tarhha.Recordset.Fields("sortgoroh")
Tarhha.Recordset.MoveNext
Wend

Next I
10
End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show

End Sub

Private Sub Label13_Change()
vadie.Refresh
vadie.RecordSource = "select * from vadie where  parvande like ('%" + Label13.Caption + "%') "
vadie.Refresh
Label12.Caption = vadie.Recordset.RecordCount
End Sub

Private Sub mnudell_Click()
On Error Resume Next
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "vadie-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Beep

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ Å—œ«Œ  ÊœÌ⁄Â —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–› Å—œ«Œ  ÊœÌ⁄Â") = vbYes Then
vadie.Recordset.Delete
End If

End Sub

Private Sub mnuhno_Click()
Entekhab.Show

End Sub

Private Sub text_auto_sabt_Change()
If text_auto_sabt = "auto_sabt_bank_stu" Then
Call Command1_Click
text_auto_sabt = ""
End If

End Sub

Private Sub Text1_Change()
Student.Refresh
Student.RecordSource = "select * from student where name like ('%" + Text1.Text + "%') or famil like ('%" + Text1.Text + "%') or nf like ('%" + Text1.Text + "%') or PARVANDE like ('%" + Text1.Text + "%') "
Student.Refresh

vadie.Refresh
vadie.RecordSource = "select * from vadie where kod like ('%" + Text1.Text + "%') or parvande like ('%" + Text1.Text + "%')or vazeyat like ('%" + Text1.Text + "%')or tozih like ('%" + Text1.Text + "%') "
vadie.Refresh



Label44.Caption = vadie.Recordset.RecordCount


End Sub


Private Sub Text2_Click()
Text2.Text = ""

End Sub

Private Sub Text2_DblClick()
Text2.Text = ""

End Sub

Private Sub Text3_Change()
Combo1.Text = Text3.Text

End Sub
