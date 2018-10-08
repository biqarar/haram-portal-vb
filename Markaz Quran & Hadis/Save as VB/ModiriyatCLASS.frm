VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ModiriyatCLASS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "À»  Ê «’·«Õ «ÿ·«⁄«  ò·«”"
   ClientHeight    =   10065
   ClientLeft      =   1425
   ClientTop       =   2595
   ClientWidth     =   14160
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ModiriyatCLASS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "„‘«ÂœÂ ·Ì”  ò·«”Ì"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "„œÌ—Ì  ·Ì”  ò·«”Ì"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3000
      Width           =   1575
   End
   Begin ComctlLib.ProgressBar PB1 
      Height          =   135
      Left            =   11760
      TabIndex        =   76
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10800
      TabIndex        =   73
      Text            =   "«‰ Œ«» ò‰Ìœ"
      Top             =   8520
      Width           =   3255
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "«‰ ﬁ«· »Â »—‰«„Â «ò”·"
      Height          =   465
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   9000
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Å«ò”«“Ì"
      Height          =   375
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13200
      Picture         =   "ModiriyatCLASS.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "—òÊ—œÌ òÂ „ÌŒÊ«ÂÌœ Õ–› ‰„«ÌÌœ —« «‰ Œ«» Ê «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10800
      Picture         =   "ModiriyatCLASS.frx":4E8F
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "ÃÂ  «÷«›Â ò—œ‰ —òÊ—œ «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   7200
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000002&
      Height          =   6615
      ItemData        =   "ModiriyatCLASS.frx":88EC
      Left            =   10800
      List            =   "ModiriyatCLASS.frx":88EE
      TabIndex        =   68
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Å«ò”«“Ì"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ã«Ìê“Ì‰Ì «ÿ·«⁄« "
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      DisabledPicture =   "ModiriyatCLASS.frx":88F0
      DownPicture     =   "ModiriyatCLASS.frx":2D56A
      DragIcon        =   "ModiriyatCLASS.frx":521E4
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
      Left            =   4560
      Picture         =   "ModiriyatCLASS.frx":76E5E
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   2400
      Width           =   495
   End
   Begin VB.Frame Frame9 
      Caption         =   "Motor"
      Height          =   375
      Left            =   3600
      TabIndex        =   64
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
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
         Connect         =   $"ModiriyatCLASS.frx":9BAD8
         OLEDBString     =   $"ModiriyatCLASS.frx":9BB61
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
         Connect         =   $"ModiriyatCLASS.frx":9BBEA
         OLEDBString     =   $"ModiriyatCLASS.frx":9BC73
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
         Connect         =   $"ModiriyatCLASS.frx":9BCFC
         OLEDBString     =   $"ModiriyatCLASS.frx":9BD85
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
         Connect         =   $"ModiriyatCLASS.frx":9BE0E
         OLEDBString     =   $"ModiriyatCLASS.frx":9BE97
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
         Connect         =   $"ModiriyatCLASS.frx":9BF20
         OLEDBString     =   $"ModiriyatCLASS.frx":9BFA9
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
         Connect         =   $"ModiriyatCLASS.frx":9C032
         OLEDBString     =   $"ModiriyatCLASS.frx":9C0BB
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
         Connect         =   $"ModiriyatCLASS.frx":9C144
         OLEDBString     =   $"ModiriyatCLASS.frx":9C1CD
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
         Connect         =   $"ModiriyatCLASS.frx":9C256
         OLEDBString     =   $"ModiriyatCLASS.frx":9C2DF
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
         Left            =   120
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
         Connect         =   $"ModiriyatCLASS.frx":9C368
         OLEDBString     =   $"ModiriyatCLASS.frx":9C3F1
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
      Begin MSAdodcLib.Adodc userprofiletable 
         Height          =   330
         Left            =   240
         Top             =   120
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
         Connect         =   $"ModiriyatCLASS.frx":9C47A
         OLEDBString     =   $"ModiriyatCLASS.frx":9C503
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
   Begin VB.CheckBox Check1 
      Caption         =   " €ÌÌ— «ÿ·«⁄« "
      Height          =   345
      Left            =   3720
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      DisabledPicture =   "ModiriyatCLASS.frx":9C58C
      DownPicture     =   "ModiriyatCLASS.frx":C1206
      DragIcon        =   "ModiriyatCLASS.frx":E5E80
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
      Left            =   4560
      Picture         =   "ModiriyatCLASS.frx":10AAFA
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   4095
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   3255
      Begin VB.Label ltaza 
         AutoSize        =   -1  'True
         Caption         =   "-  "
         DataField       =   "TedadJalasat"
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
         TabIndex        =   46
         Top             =   3240
         Width           =   225
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å«Ì«‰"
         Height          =   345
         Left            =   2040
         TabIndex        =   45
         Top             =   2760
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
         TabIndex        =   44
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ ‘—Ê⁄"
         Height          =   345
         Left            =   2040
         TabIndex        =   43
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ Ã·”« "
         Height          =   345
         Left            =   2040
         TabIndex        =   42
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "—Ê“ Â«Ì ò·«”"
         Height          =   345
         Left            =   2040
         TabIndex        =   41
         Top             =   3480
         Width           =   885
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
         TabIndex        =   40
         Top             =   2520
         Width           =   135
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
         TabIndex        =   39
         Top             =   3600
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
         Height          =   330
         Index           =   0
         Left            =   2040
         TabIndex        =   38
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   330
         Left            =   2040
         TabIndex        =   37
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "„ﬁÿ⁄"
         Height          =   330
         Left            =   2040
         TabIndex        =   36
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
         Height          =   330
         Left            =   2040
         TabIndex        =   35
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄ "
         Height          =   330
         Left            =   2040
         TabIndex        =   34
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "„œ—”"
         Height          =   330
         Left            =   2040
         TabIndex        =   33
         Top             =   2040
         Width           =   405
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
         TabIndex        =   32
         Top             =   360
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
         TabIndex        =   31
         Top             =   600
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
         TabIndex        =   30
         Top             =   960
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
         TabIndex        =   29
         Top             =   1680
         Width           =   135
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
         TabIndex        =   28
         Top             =   1320
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
         TabIndex        =   27
         Top             =   2160
         Width           =   225
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
         TabIndex        =   26
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   " «"
         Height          =   330
         Left            =   720
         TabIndex        =   25
         Top             =   1680
         Width           =   120
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "òœ ò·«”"
      Height          =   855
      Left            =   3480
      TabIndex        =   23
      Top             =   0
      Width           =   1575
      Begin VB.TextBox TEP 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "À»  ò·«” ÃœÌœ"
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   10200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox JostojoCh 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ã” ÃÊ"
      Height          =   345
      Left            =   3960
      TabIndex        =   17
      Top             =   4200
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "„‘Œ’«  ò·«” ÃœÌœ"
      Height          =   4575
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   360
         TabIndex        =   82
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   2760
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   2160
         TabIndex        =   16
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox Ostad 
         BackColor       =   &H00E0E0E0&
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
         Left            =   360
         TabIndex        =   4
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox Madras 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2760
         TabIndex        =   5
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   2760
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   2760
         TabIndex        =   10
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "ModiriyatCLASS.frx":12F774
         Left            =   360
         List            =   "ModiriyatCLASS.frx":12F776
         TabIndex        =   15
         Text            =   "Â„Â —Ê“Â"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton Am 
         Caption         =   "’»Õ"
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Pm 
         Caption         =   "⁄’—"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   2760
         TabIndex        =   12
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2760
         TabIndex        =   14
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   360
         TabIndex        =   13
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "ModiriyatCLASS.frx":12F778
         Left            =   2760
         List            =   "ModiriyatCLASS.frx":12F77A
         TabIndex        =   1
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "œÊ—Â"
         Height          =   345
         Left            =   1680
         TabIndex        =   83
         Top             =   3840
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ/ ‰«„ ò·«”"
         Height          =   345
         Left            =   4080
         TabIndex        =   61
         Top             =   3840
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   345
         Index           =   1
         Left            =   2160
         TabIndex        =   60
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "„ﬁÿ⁄"
         Height          =   345
         Left            =   4680
         TabIndex        =   59
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«” «œ"
         Height          =   345
         Left            =   2160
         TabIndex        =   58
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "„œ—”"
         Height          =   345
         Left            =   4680
         TabIndex        =   57
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄  ‘—Ê⁄"
         Height          =   345
         Left            =   4320
         TabIndex        =   56
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄  Å«Ì«‰"
         Height          =   345
         Left            =   1680
         TabIndex        =   55
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ ‘—Ê⁄"
         Height          =   345
         Left            =   4440
         TabIndex        =   54
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å«Ì«‰"
         Height          =   345
         Left            =   1800
         TabIndex        =   53
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "—Ê“ Â«Ì ò·«”Ì"
         Height          =   345
         Left            =   1680
         TabIndex        =   52
         Top             =   3360
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ Ã·”« "
         Height          =   345
         Left            =   4320
         TabIndex        =   51
         Top             =   2880
         Width           =   915
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "€Ì»  „Ã«“"
         Height          =   345
         Left            =   4560
         TabIndex        =   50
         Top             =   3360
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "„œ  “„«‰"
         Height          =   345
         Left            =   1800
         TabIndex        =   49
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "œÊ—Â"
         Height          =   345
         Left            =   4800
         TabIndex        =   48
         Top             =   480
         Width           =   270
      End
   End
   Begin MSDataGridLib.DataGrid DMClass 
      Bindings        =   "ModiriyatCLASS.frx":12F77C
      Height          =   4935
      Left            =   120
      TabIndex        =   62
      Top             =   4680
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
      DefColWidth     =   107
      HeadLines       =   1
      RowHeight       =   30
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "·Ì”  ò·«” Â«"
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "Kodclass"
         Caption         =   "òœ ò·«”"
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
      BeginProperty Column02 
         DataField       =   "Maqta"
         Caption         =   "„ﬁÿ⁄"
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
      BeginProperty Column04 
         DataField       =   "Ostad"
         Caption         =   "«” «œ"
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
         DataField       =   "Zamaneshoro"
         Caption         =   "”«⁄  ‘—Ê⁄"
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
         DataField       =   "Zamanepayan"
         Caption         =   "”«⁄  Å«Ì«‰"
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
         DataField       =   "Madras"
         Caption         =   "„œ—”"
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
         DataField       =   "Ayamehafte"
         Caption         =   "«Ì«„ Â› Â"
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
         DataField       =   "Tshoro"
         Caption         =   " «—ÌŒ ‘—Ê⁄"
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
         DataField       =   "Tpayan"
         Caption         =   " «—ÌŒ ÅÌ«Ì«‰"
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
         DataField       =   "Tedadjalasat"
         Caption         =   " ⁄œ«œ Ã·”« "
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
         DataField       =   "Sobh"
         Caption         =   "’»Õ"
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
         DataField       =   "Asr"
         Caption         =   "⁄’—"
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
      BeginProperty Column15 
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
      BeginProperty Column16 
         DataField       =   "QMojaz"
         Caption         =   " ⁄œ«œ €Ì»  „Ã«“"
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
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1620.284
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   63
      Top             =   9690
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
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
            Text            =   " «—ÌŒ «„—Ê“"
            TextSave        =   " «—ÌŒ «„—Ê“"
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
   Begin ComctlLib.ProgressBar PB2 
      Height          =   135
      Left            =   11760
      TabIndex        =   77
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   345
      Left            =   11880
      TabIndex        =   81
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label pleaswait 
      AutoSize        =   -1  'True
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11760
      TabIndex        =   78
      Top             =   9000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "·Ì”  «‰ Œ«»Ì"
      Height          =   345
      Left            =   13080
      TabIndex        =   75
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ò·«” Â«Ì"
      Height          =   345
      Left            =   13200
      TabIndex        =   74
      Top             =   8160
      Width           =   675
   End
   Begin VB.Label LJostojo 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "œ— Â—òœ«„ «“ „Ê«—œ »«·« ò·„Â «Ì »‰ÊÌ”Ìœ  « Ã” ÃÊ ‘Êœ"
      Height          =   345
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Menu mnuhome 
      Caption         =   "#"
   End
   Begin VB.Menu mnumenu 
      Caption         =   " ‰ŸÌ„« "
      Begin VB.Menu mnudelclas 
         Caption         =   "Õ–› ò·«”"
      End
      Begin VB.Menu SDG 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MNUEDITE 
         Caption         =   "«’·«Õ «ÿ·«⁄«  ò·«”"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mm 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuxls 
         Caption         =   "«‰ ﬁ«· «ÿ·«⁄«  ò·«” »Â »—‰«„Â «ò”·"
      End
      Begin VB.Menu mnuselect2select 
         Caption         =   "«÷«›Â ò—œ‰ ·Ì”  ò·«” Â«Ì „ÊÃÊœ »Â ·Ì”  «‰ Œ«»Ì"
      End
      Begin VB.Menu mnuadd_en 
         Caption         =   "«÷«›Â ò—œ‰ »Â ·Ì”  «‰ Œ«»Ì"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "ModiriyatCLASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EYVAL As String

Private Sub Check1_Click()
If DMClass.AllowUpdate = False Then
DMClass.AllowUpdate = True
'Label16.Caption = "«ÿ·«⁄«  ﬁ«»·  ⁄ÌÌ— „Ì »«‘‰œ"
Else
DMClass.AllowUpdate = False
'Label16.Caption = "«ÿ«⁄«  ﬁ«»·  €ÌÌ— ‰„Ì »«‘‰œ"
mclass.Recordset.Update

End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "œ— Õ«· «Ã—«" Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where tpayan  like ('%" + "" + "%')"
mclass.Refresh


For I = 1 To mclass.Recordset.RecordCount
If mclass.Recordset.Fields("tpayan") <> "" Then
GoTo 2
Else

List1.AddItem (mclass.Recordset.Fields("kodclass") & " - " & mclass.Recordset.Fields("maqta") & " - " & mclass.Recordset.Fields("ostad"))
2 mclass.Recordset.MoveNext
End If

Next I
End If
If Combo1.Text = "»Â « „«„ —”ÌœÂ" Then


mclass.Refresh
mclass.RecordSource = "select * from mclass where tozih like ('%" + "« „«„ ò·«”" + "%')"
mclass.Refresh
For I = 1 To mclass.Recordset.RecordCount

List1.AddItem (mclass.Recordset.Fields("kodclass") & " - " & mclass.Recordset.Fields("maqta") & " - " & mclass.Recordset.Fields("ostad"))
mclass.Recordset.MoveNext

Next I




End If
If Combo1.Text = "Â„Â ò·«” Â«" Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + "" + "%')"
mclass.Refresh
For I = 1 To mclass.Recordset.RecordCount

List1.AddItem (mclass.Recordset.Fields("kodclass") & " - " & mclass.Recordset.Fields("maqta") & " - " & mclass.Recordset.Fields("ostad"))
mclass.Recordset.MoveNext

Next I


End If

End Sub

Private Sub Combo2_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where tarh like ('%" + Combo2.Text + "%')"
mclass.Refresh
End If

End Sub

Private Sub Combo2_Click()
On Error Resume Next

If EYVAL = 1 Then Exit Sub



Dim A, B, KK, K, F, T, SSS As Long


Dim SaljariSTR, CodeBakhsh As String


SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "SalJari" + "%')"
SettingUser.Refresh

SaljariSTR = SettingUser.Recordset.Fields("value")

SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "CodeBakhsh" + "%')"
SettingUser.Refresh
CodeBakhsh = SettingUser.Recordset.Fields("value")


Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where NAME like ('%" + Combo2.Text + "%')" ' „«„ Å—Ê‰œÂ Â« —« „Ì ê—œœ Ê ò”«‰Ì òÂ ÿ— Õ «‰ Œ«» ‘œÂ —« œ«—‰œ „Ì ¬Ê—œÅ
Tarhha.Refresh
    KODTARH = Tarhha.Recordset.Fields("XkodDORE")
    A = Val(SaljariSTR & CodeBakhsh & KODTARH & "00")
    B = Val(SaljariSTR & CodeBakhsh & KODTARH & "99")
    KK = SaljariSTR & CodeBakhsh & KODTARH
    Text1.Text = KK
mclass.Refresh
mclass.RecordSource = "select * from MClass where kodclass like ('%" + Text1.Text + "%')"
mclass.Refresh
mclass.Recordset.MoveFirst
mclass.Refresh

K = mclass.Recordset.Fields("kodclass")


For J = 1 To mclass.Recordset.RecordCount
F = Val(mclass.Recordset.Fields("kodclass"))
    If F > A Then '»—«Ì  «Ì‰òÂ ⁄œœ œÌê—Ì ﬁ«ÿÌ òœ ‰‘Êœ
    If F < B Then
   
   
             If F > SSS Then
            SSS = F ' »“—ê —Ì‰ —« ÅÌœ« „Ì ò‰œ
             Else
             GoTo 14
             End If
    End If
    End If
14     mclass.Recordset.MoveNext
'PB1.Value = PB1.Value + 1
Next J
TEP = SSS + 1
If TEP.Text = "1" Then
TEP.Text = Text1.Text & "01"
Else

TEP = SSS + 1 ''‰ ÌçÂ ‰Â«ÌÌ
End If


End Sub

Private Sub Combo3_Click()

On Error GoTo 10
Combo2.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where goroh like ('" & Combo3.Text & "')"
Tarhha.Refresh

Tarhha.Recordset.Sort = "sortname"

For I = 1 To Tarhha.Recordset.RecordCount
Combo2.AddItem (Tarhha.Recordset.Fields("name"))
Tarhha.Recordset.MoveNext
Next I
Combo2.Text = Combo2.List(0)


10 Exit Sub
If Combo3.Text = "⁄„Ê„Ì" Then
Combo2.Clear

Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "1" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo2.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo2.Text = Combo2.List(0)
End If


If Combo3.Text = "ò«—ê«Â Â«" Then
Combo2.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "3" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo2.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo2.Text = Combo2.List(0)
End If

If Combo3.Text = " —»Ì  „—»Ì" Then
Combo2.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "4" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo2.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo2.Text = Combo2.List(0)
End If


If Combo3.Text = "Õ›Ÿ ﬁ—¬‰ ò—Ì„" Then

Combo2.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "2" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo2.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo2.Text = Combo2.List(0)
End If

If Combo3.Text = "„ÃÂÊ·" Then
Combo2.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "0" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo2.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo4_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where ayamehafte like ('%" + Combo4.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub Command1_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "mclass-newsabt" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + TEP.Text + "%')"
mclass.Refresh
If mclass.Recordset.BOF = True Or mclass.Recordset.EOF = True Then
GoTo 1
Else
MsgBox "òœ ò·«”  ò—«—Ì «” ", vbCritical + vbOKOnly, "Œÿ«"

Exit Sub
End If
Exit Sub
1:
If Text9.Text = "" Then


MsgBox " ⁄œ«œ €Ì»  „Ã«“ »—«Ì «Ì‰ ò·«” —« Ê«—œ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"

Exit Sub
End If

'naboodane kelas dar haman safhe
mclass.Refresh
mclass.RecordSource = "select * from mclass where zamaneshoro like ('%" + Text4.Text + "%') and zamanepayan like ('%" + Text5.Text + "%') and madras like ('%" + Madras.Text + "%')  "
mclass.Refresh
'or I = 1 To mclass.Recordset.RecordCount
'f mclass.Recordset.Fields("tozih") = "« „«„ ò·«”" Then


'If Text9.Text = "" Then


'MsgBox " ⁄œ«œ €Ì»  „Ã«“ »—«Ì «Ì‰ ò·«” —« Ê«—œ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"

'Exit Sub
'End If



mclass.Refresh
mclass.Recordset.AddNew
mclass.Recordset.Fields("kodclass") = TEP.Text
mclass.Recordset.Fields("tarh") = Combo2.Text
mclass.Recordset.Fields("maqta") = Text2.Text
mclass.Recordset.Fields("tozih") = Text3.Text
mclass.Recordset.Fields("ostad") = Ostad.Text
mclass.Recordset.Fields("madras") = Madras.Text
mclass.Recordset.Fields("zamaneshoro") = Text4.Text
mclass.Recordset.Fields("zamanepayan") = Text5.Text
mclass.Recordset.Fields("qmojaz") = Text9.Text
mclass.Recordset.Fields("ayamehafte") = Combo4.Text
mclass.Recordset.Fields("tshoro") = Text6.Text
mclass.Recordset.Fields("tpayan") = Text7.Text
mclass.Recordset.Fields("tedadjalasat") = Text8.Text
If Am.Value = True Then mclass.Recordset.Fields("sobh") = Am.Caption
If Pm.Value = True Then mclass.Recordset.Fields("asr") = Pm.Caption
mclass.Recordset.Fields("d") = Taqvim.Tarikh.Caption
mclass.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
mclass.Recordset.Fields("dore") = Text11.Text
mclass.Recordset.Update
mclass.Refresh


End Sub

Private Sub Command10_Click()
On Error Resume Next

List1.RemoveItem (List1.ListIndex)
mclass.Recordset.MoveNext



End Sub

Private Sub Command11_Click()
If MsgBox("·Ì”  «‰ Œ«»Ì ò·«” Â« Å«ò ŒÊ«Âœ ‘œ Ê ·Ì”  ò·«” œ— Â„«‰ ·Ì”  «‰ Œ«»Ì Ê«—œ „Ì ‘œ!!! ¬Ì« „ÿ„∆‰ Â” Ìœ", vbQuestion + vbYesNo, "Â‘œ«—") = vbYes Then
List1.Clear
Dim L_klass As String
L_klass = mclass.Recordset.Fields("kodclass")

'ﬁ—¬ ‰¬„Ê“«‰ Õ«÷—
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + L_klass + "%') or clas2 like ('%" + L_klass + "%') or clas3 like ('%" + L_klass + "%') or clas4 like ('%" + L_klass + "%') or clas5 like ('%" + L_klass + "%')"
Student.Refresh
List1.Clear
Label20.Caption = Student.Recordset.RecordCount

For I = 1 To Student.Recordset.RecordCount
List1.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
Student.Recordset.MoveNext
Next I


End If

End Sub

Private Sub Command2_Click()

'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
Dim AD As String
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:
If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "InformationClassFormool.xlsx")

End If
If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "InformationClassFormool.xlsx")

End If
'oExcel.ActiveSheet.Range("f1").Value = "«„Ê— ¬„Ê“‘Ì"
'oExcel.ActiveSheet.Range("j1").Value = Text10.Text
'oExcel.ActiveSheet.Range("M1").Value = Taqvim.Label1.Caption
'ekhtar.Recordset.MoveFirst

'Set oExcel = GetObject("\\yafatemeh2-pc\F\Markaz Quran & Hadis\FORMXLS\InformationClassFormool.xlsx")

'\\yafatemeh2-pc\F\Markaz Quran & Hadis\FORMXLS\InformationClassFormool.xlsx

Dim NumberOfRows As Integer
NumberOfRows = mclass.Recordset.RecordCount
For r = 3 To NumberOfRows + 2

oExcel.ActiveSheet.Range("B" & r).Value = mclass.Recordset.Fields("KODCLASS")
oExcel.ActiveSheet.Range("c" & r).Value = mclass.Recordset.Fields("TARH")
oExcel.ActiveSheet.Range("D" & r).Value = mclass.Recordset.Fields("MAQTA")
oExcel.ActiveSheet.Range("E" & r).Value = mclass.Recordset.Fields("OSTAD")
oExcel.ActiveSheet.Range("F" & r).Value = mclass.Recordset.Fields("zamaneshoro")
oExcel.ActiveSheet.Range("G" & r).Value = mclass.Recordset.Fields("zamanePAYAN")
oExcel.ActiveSheet.Range("H" & r).Value = mclass.Recordset.Fields("MADRAS")
oExcel.ActiveSheet.Range("I" & r).Value = mclass.Recordset.Fields("TSHORO")
oExcel.ActiveSheet.Range("J" & r).Value = mclass.Recordset.Fields("TPAYAN")
oExcel.ActiveSheet.Range("K" & r).Value = mclass.Recordset.Fields("AYAMEHAFTE")
oExcel.ActiveSheet.Range("L" & r).Value = mclass.Recordset.Fields("tedadjalasat")

Me.STU2CLASS.Refresh
Me.STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" & mclass.Recordset.Fields("kodclass") & "%')"
Me.STU2CLASS.Refresh

'hazfi = Me.STU2CLASS.Recordset.RecordCount

oExcel.ActiveSheet.Range("M" & r).Value = STU2CLASS.Recordset.RecordCount

'Student.Refresh
'Student.RecordSource = "select * from student where clas1 like ('%" + mclass.Recordset.Fields("kodclass") + "%') or clas2 like ('%" + mclass.Recordset.Fields("kodclass") + "%') or clas3 like ('%" + mclass.Recordset.Fields("kodclass") + "%') or clas4 like ('%" + mclass.Recordset.Fields("kodclass") + "%') or clas5 like ('%" + mclass.Recordset.Fields("kodclass") + "%')"
'Student.Refresh

'oExcel.ActiveSheet.Range("N" & r).Value = Student.Recordset.RecordCount



STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + mclass.Recordset.Fields("kodclass") + "%') and elat like ('%" + "€Ì» " + "%') "
STU2CLASS.Refresh

oExcel.ActiveSheet.Range("n" & r).Value = Me.STU2CLASS.Recordset.RecordCount
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + mclass.Recordset.Fields("kodclass") + "%') and elat like ('%" + "«‰’—«›" + "%') "
STU2CLASS.Refresh
oExcel.ActiveSheet.Range("n" & r).Value = Val(oExcel.ActiveSheet.Range("n" & r).Value) + Me.STU2CLASS.Recordset.RecordCount

'STU2CLASS.Refresh
'STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + mclass.Recordset.Fields("kodclass") + "%') and elat like ('%" + "« „«„" + "%') "
'STU2CLASS.Refresh

'oExcel.ActiveSheet.Range("Q" & r).Value = Me.STU2CLASS.Recordset.RecordCount
'etmam = Me.STU2CLASS.Recordset.RecordCount


'oExcel.ActiveSheet.Range("O" & r).Value = Val(hazfi) - Val(etmam)

mclass.Recordset.MoveNext


'On Error Resume Next

Next r



MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = "«ÿ·«⁄«  ò«„· ò·«” Â«"
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

Private Sub Command3_Click()
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String

Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

Dim DataArray(1 To 2000, 1 To 20) As Variant

Dim r As Integer
Dim NumberOfRows As Integer
NumberOfRows = mclass.Recordset.RecordCount
mclass.Recordset.MoveFirst

For r = 1 To NumberOfRows
DataArray(r, 1) = mclass.Recordset.Fields("kodclass")
DataArray(r, 2) = mclass.Recordset.Fields("tarh")
DataArray(r, 3) = mclass.Recordset.Fields("maqta")
DataArray(r, 4) = mclass.Recordset.Fields("tozih")
DataArray(r, 5) = mclass.Recordset.Fields("ostad")
DataArray(r, 6) = mclass.Recordset.Fields("zamaneshoro")
DataArray(r, 7) = mclass.Recordset.Fields("zamanepayan")
DataArray(r, 8) = mclass.Recordset.Fields("madras")
DataArray(r, 9) = mclass.Recordset.Fields("tshoro")
DataArray(r, 10) = mclass.Recordset.Fields("tpayan")
DataArray(r, 11) = mclass.Recordset.Fields("ayamehafte")
DataArray(r, 12) = mclass.Recordset.Fields("sobh")
DataArray(r, 13) = mclass.Recordset.Fields("asr")
DataArray(r, 14) = mclass.Recordset.Fields("tedadjalasat")
DataArray(r, 15) = mclass.Recordset.Fields("op")
DataArray(r, 16) = mclass.Recordset.Fields("d")
On Error Resume Next
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" + mclass.Recordset.Fields("kodclass") + "%')"
STU2CLASS.Refresh
DataArray(r, 17) = STU2CLASS.Recordset.RecordCount



mclass.Recordset.MoveNext
Next
Set oSheet = oBook.Worksheets(1)
'oSheet.Range("A1:E1").Font.Bold = True

oSheet.Range("A1:P1").Font.Bold = True


oSheet.Range("A1 :P1").Value = Array("òœ ò·«”", "ÿ—Õ", "„ﬁÿ⁄", " Ê÷ÌÕ", "«” «œ", "”«⁄  ‘—Ê⁄", "”«⁄  Å«Ì«‰", "„œ—”", "", "", "", "", "", "", "", "")





oSheet.Range("A2").Resize(NumberOfRows, 20).Value = DataArray
AD = "«ÿ·«⁄«  ò·«”"
oBook.SaveAs AD
'oBook.SaveAs "C:\Report.xls"
oExcel.quit
mclass.Recordset.MoveFirst
MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â ‘œ‰œ", 64, "À»  «ÿ·«⁄« "

End Sub

Private Sub Command4_Click()
If EYVAL = 0 Then
MsgBox "«» œ« »« œÊ »«— ò·Ìò ò—œ‰ »— —ÊÌ „‘Œ’«  ò·«” ° ¬‰ —« ¬„«œÂ «’·«Õ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
Else
GoTo 1
End If
Exit Sub
1:
EYVAL = 0


mclass.Recordset.Fields("kodclass") = TEP.Text
mclass.Recordset.Fields("tarh") = Combo2.Text
mclass.Recordset.Fields("maqta") = Text2.Text
mclass.Recordset.Fields("tozih") = Text3.Text
mclass.Recordset.Fields("ostad") = Ostad.Text
mclass.Recordset.Fields("madras") = Madras.Text
mclass.Recordset.Fields("zamaneshoro") = Text4.Text
mclass.Recordset.Fields("zamanepayan") = Text5.Text
mclass.Recordset.Fields("qmojaz") = Text9.Text
mclass.Recordset.Fields("ayamehafte") = Combo4.Text
mclass.Recordset.Fields("tshoro") = Text6.Text
mclass.Recordset.Fields("tpayan") = Text7.Text
mclass.Recordset.Fields("tedadjalasat") = Text8.Text
If Am.Value = True Then mclass.Recordset.Fields("sobh") = Am.Caption
If Pm.Value = True Then mclass.Recordset.Fields("asr") = Pm.Caption
mclass.Recordset.Fields("d") = Taqvim.Tarikh.Caption
mclass.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text


mclass.Recordset.Update


MsgBox "«ÿ·«⁄«  Ã«Ìê“Ì‰ ‘œ", vbInformation, "Ã«Ìê“Ì‰Ì «ÿ·«⁄« "

JostojoCh.Visible = True
Command1.Enabled = True
Command6.Enabled = True


End Sub

Private Sub Command5_Click()
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ ·Ì”  —« Å«ò”«“Ì ò‰Ìœ", vbQuestion + vbYesNo, "Å«ò”«“Ì") = vbYes Then



List1.Clear
End If

End Sub

Private Sub Command6_Click()
TEP.Text = ""
 Combo2.Text = ""
 Text2.Text = ""
 Text3.Text = ""
Ostad.Text = ""
 Madras.Text = ""
Text4.Text = ""
Text5.Text = ""
Text9.Text = ""
Combo4.Text = ""
Text6.Text = ""
Text7.Text = ""
 Text8.Text = ""
 Am.Value = False
 Pm.Value = False
Combo3.Text = "«‰ Œ«» ò‰Ìœ"
Combo2.Text = ""


End Sub

Private Sub Command7_Click()
List1.AddItem (mclass.Recordset.Fields("kodclass") & " - " & mclass.Recordset.Fields("maqta") & " - " & mclass.Recordset.Fields("ostad"))
On Error Resume Next

mclass.Recordset.MoveNext

End Sub

Private Sub Command8_Click()
'On Error Resume Next
'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
Dim AD As String
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:
If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "ClassSTUDENT.xlsx")

End If
If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "ClassSTUDENT.xlsx")

End If

pleaswait.Visible = True

PB1.Visible = True
Pb2.Visible = True
 

Pb2.Max = List1.ListCount

'entekhab list clasi
Dim klassCount, AfradToXLSX, Edame As Integer
Dim kodclassPrint, parvandePrint As String
Edame = 3
For klassCount = 1 To List1.ListCount
'Shroe List classs-Shroe List class-Shroe List class-Shroe List class-Shroe List class-Shroe List class-Shroe List class-

mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Mid(List1.List(klassCount - 1), 1, 6) + "%')"
mclass.Refresh
'À»  òœ ò·«” œ— Õ«›ŸÂ œ«Œ·Ì
kodclassPrint = mclass.Recordset.Fields("kodclass")
'payanne sabt class dar hafeze
Me.STU2CLASS.Refresh
Me.STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" & kodclassPrint & "%')"
Me.STU2CLASS.Refresh

If STU2CLASS.Recordset.RecordCount <> 0 Then
PB1.Max = STU2CLASS.Recordset.RecordCount
End If


'vared kardane afrad dar klass-vared kardane afrad dar klass-vared kardane afrad dar klass-vared kardane afrad dar klass-
For AfradToXLSX = Edame To Me.STU2CLASS.Recordset.RecordCount + (Edame - 1)
'start-start-start-start-start-start-start-start-start-start-start-start-start-
parvandePrint = Me.STU2CLASS.Recordset.Fields("parvande")

Me.Student.Refresh
Me.Student.RecordSource = "select * from student where parvande like ('%" & parvandePrint & "%')"
Me.Student.Refresh

'shoro amaliyat sabt dar xlsx-shoro amaliyat sabt dar xlsx-shoro amaliyat sabt dar xlsx-shoro amaliyat sabt dar xlsx-
'sabt etlaate klass-sabt etlaate klass-sabt etlaate klass-sabt etlaate klass-sabt etlaate klass-
oExcel.ActiveSheet.Range("B" & AfradToXLSX).Value = mclass.Recordset.Fields("KODCLASS")
oExcel.ActiveSheet.Range("D" & AfradToXLSX).Value = mclass.Recordset.Fields("MAQTA")
oExcel.ActiveSheet.Range("c" & AfradToXLSX).Value = mclass.Recordset.Fields("OSTAD")
oExcel.ActiveSheet.Range("e" & AfradToXLSX).Value = mclass.Recordset.Fields("TSHORO")
oExcel.ActiveSheet.Range("f" & AfradToXLSX).Value = mclass.Recordset.Fields("TPAYAN")
oExcel.ActiveSheet.Range("G" & AfradToXLSX).Value = mclass.Recordset.Fields("zamaneshoro") & " - " & mclass.Recordset.Fields("zamanePAYAN")
oExcel.ActiveSheet.Range("h" & AfradToXLSX).Value = mclass.Recordset.Fields("tarh")
oExcel.ActiveSheet.Range("i" & AfradToXLSX).Value = mclass.Recordset.Fields("ayamehafte")
'end of sabt etelaate klass-end of sabt etelaate klass-end of sabt etelaate klass-end of sabt etelaate klas
'sabt etelaate qoran Amoooz-sabt etelaate qoran Amoooz-sabt etelaate qoran Amoooz-sabt etelaate qoran Amoooz-


oExcel.ActiveSheet.Range("j" & AfradToXLSX).Value = Student.Recordset.Fields("tavalod")

oExcel.ActiveSheet.Range("k" & AfradToXLSX).Value = Student.Recordset.Fields("parvande")
oExcel.ActiveSheet.Range("l" & AfradToXLSX).Value = Student.Recordset.Fields("name")
oExcel.ActiveSheet.Range("m" & AfradToXLSX).Value = Student.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("n" & AfradToXLSX).Value = Student.Recordset.Fields("namepedar")
oExcel.ActiveSheet.Range("o" & AfradToXLSX).Value = Student.Recordset.Fields("tell")
oExcel.ActiveSheet.Range("p" & AfradToXLSX).Value = Student.Recordset.Fields("mob")
oExcel.ActiveSheet.Range("s" & AfradToXLSX).Value = Student.Recordset.Fields("tozih")
'end of sabt quran Amozan-end of sabt quran Amozan-end of sabt quran Amozan-end of sabt quran Amozan-
'start of stu2class-start of stu2class-start of stu2class-start of stu2class-start of stu2class-
oExcel.ActiveSheet.Range("t" & AfradToXLSX).Value = Me.STU2CLASS.Recordset.Fields("tshoro")
oExcel.ActiveSheet.Range("u" & AfradToXLSX).Value = Me.STU2CLASS.Recordset.Fields("tpayan")
oExcel.ActiveSheet.Range("v" & AfradToXLSX).Value = Me.STU2CLASS.Recordset.Fields("elat")
'end of sabt stu2class-end of sabt stu2class-end of sabt stu2class-
'sabt 5 class-sabt 5 class-sabt 5 class-sabt 5 class-sabt 5 class-sabt 5 class-
oExcel.ActiveSheet.Range("w" & AfradToXLSX).Value = Student.Recordset.Fields("clas1") & "-" & Student.Recordset.Fields("clas2") & "-" & Student.Recordset.Fields("clas3") & "-" & Student.Recordset.Fields("clas4") & "-" & Student.Recordset.Fields("clas5")
'end of sabt 5 classs-end of sabt 5 classs-end of sabt 5 classs-end of sabt 5 classs-


'end of amaliyat sabt dar xlsx-end of amaliyat sabt dar xlsx-end of amaliyat sabt dar xlsx-end of amaliyat sabt dar xlsx-
STU2CLASS.Recordset.MoveNext
PB1.Value = PB1.Value + 1
'End Of vared kardan-End Of vared kardan-End Of vared kardan-End Of vared kardan-End Of vared kardan-
Next AfradToXLSX
Edame = Edame + STU2CLASS.Recordset.RecordCount

'klass badi ro vared mikone
PB1.Value = 0
Pb2.Value = Pb2.Value + 1
Next klassCount



PB1.Value = 0
Pb2.Value = 0

pleaswait.Visible = False
PB1.Visible = False
Pb2.Visible = False



MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = "«ÿ·«⁄«  ò«„· ò·«” Â«"
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

Private Sub Command9_Click()
FClassroom.Show
FClassroom.Text3.Text = Me.lkodclass.Caption


End Sub

Private Sub DMClass_DblClick()
List1.AddItem (mclass.Recordset.Fields("kodclass") & " - " & mclass.Recordset.Fields("maqta") & " - " & mclass.Recordset.Fields("ostad"))

End Sub

Private Sub Form_DblClick()
On Error Resume Next
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + "" + "%')"
mclass.Refresh
End Sub

Private Sub Form_Load()
EYVAL = 0

Combo1.AddItem ("œ— Õ«· «Ã—«")
Combo1.AddItem ("»Â « „«„ —”ÌœÂ")
Combo1.AddItem ("Â„Â ò·«” Â«")

For J = 1 To 10
Madras.AddItem (J)
Next J


Me.stb1.Panels(1).Text = user.OP.Text
Combo4.AddItem ("Â„Â —Ê“Â")
Combo4.AddItem ("Â„Â —Ê“Â / ç—Œ‘Ì")
Combo4.AddItem ("—Ê“ Â«Ì “ÊÃ / ç—Œ‘Ì")
Combo4.AddItem ("—Ê“ Â«Ì ›—œ / ç—Œ‘Ì")

Combo4.AddItem ("—Ê“ Â«Ì “ÊÃ")
Combo4.AddItem ("—Ê“ Â«Ì ›—œ")
Combo4.AddItem ("‘‰»Â Â«")
Combo4.AddItem ("Ìò ‘‰»Â Â«")
Combo4.AddItem ("œÊ ‘‰»Â Â«")
Combo4.AddItem ("”Â ‘‰»Â Â«")
Combo4.AddItem ("çÂ«— ‘‰»Â Â«")
Combo4.AddItem ("Å‰ç ‘‰»Â Â«")
Combo4.AddItem ("Ã„⁄Â Â«")
Combo4.AddItem ("Â› êÌ")
Combo4.AddItem ("„«ÂÌ«‰Â")
Combo4.AddItem ("”«·«‰Â")
Combo4.AddItem ("‰«„⁄·Ê„")
'Combo4.AddItem ("")

Me.stb1.Panels(1).Text = user.OP.Text
Me.stb1.Panels(3).Text = Taqvim.Tarikh.Caption




Tarhha.Refresh
Tarhha.Recordset.Sort = "sortgoroh"


For I = 1 To Tarhha.Recordset.RecordCount
Combo3.AddItem (Tarhha.Recordset.Fields("goroh"))
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

Private Sub JostojoCh_Click()
If JostojoCh.Value = 1 Then
TEP.Enabled = True

LJostojo.Visible = True
Else
LJostojo.Visible = False
TEP.Enabled = False

End If

End Sub

Private Sub List1_Click()
On Error Resume Next
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Mid(List1.Text, 1, 6) + "%')"
mclass.Refresh
End Sub

Private Sub m10_Click()
EmtahanF.Show
End Sub

Private Sub m11_Click()

Karname.Show

End Sub

Private Sub m2_Click()
BankStudent.Show

End Sub

Private Sub m3_Click()
QeybatF.Show

End Sub

Private Sub m4_Click()
Gozaresh.Show

End Sub

Private Sub m6q_Click()
Beep

End Sub

Private Sub m7_Click()
FClassroom.Show

End Sub

Private Sub m8_Click()
FClassroom.Show

End Sub

Private Sub List1_DblClick()
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
A = Mid(List1.Text, 1, 7)





Scan.Text1.Text = A

Scan.Show
A = SettingF.ScanAdress.Caption & A & "\" & A & ".jpg"
'A = Student.Recordset.Fields("scan")
Scan.Im1.Picture = LoadPicture(A)

Exit Sub
End If

If Entekhab.net.Checked = True Then
A = Mid(List1.Text, 1, 7)

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

Private Sub Madras_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where madras like ('%" + Madras.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub mnuadd_en_Click()
List1.AddItem (mclass.Recordset.Fields("kodclass") & " - " & mclass.Recordset.Fields("maqta") & " - " & mclass.Recordset.Fields("ostad"))

End Sub

Private Sub mnudelclas_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "mclass-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where parvande like ('%" + Me.lkodclass.Caption + "%')"
STU2CLASS.Refresh

If STU2CLASS.Recordset.BOF = False Or STU2CLASS.Recordset.EOF = False Then

MsgBox "«„ò«‰ Õ–› ò·«” Â«ÌÌ òÂ ﬁ—¬‰ ¬„Ê“ œ— ¬‰ ‘—ò  ò—œÂ «”  ÊÃÊœ ‰œ«—œ", vbCritical + vbOKOnly, "Œÿ«"

Exit Sub
End If

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ò·«” —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–› ò·«”") = vbYes Then

mclass.Recordset.Delete
Else
Exit Sub
End If

End Sub

Private Sub MNUEDITE_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "mclass-edit" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Beep
On Error Resume Next

JostojoCh.Value = 0
JostojoCh.Visible = False
Command6.Enabled = False



Command1.Enabled = False
Beep
EYVAL = 1


'mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + lkodclass.Caption + "%')"
mclass.Refresh

TEP.Text = mclass.Recordset.Fields("kodclass")
 Combo2.Text = mclass.Recordset.Fields("tarh")
 Text2.Text = mclass.Recordset.Fields("maqta")
 Text3.Text = mclass.Recordset.Fields("tozih")
Ostad.Text = mclass.Recordset.Fields("ostad")
 Madras.Text = mclass.Recordset.Fields("madras")
Text4.Text = mclass.Recordset.Fields("zamaneshoro")
Text5.Text = mclass.Recordset.Fields("zamanepayan")
Text9.Text = mclass.Recordset.Fields("qmojaz")
Combo4.Text = mclass.Recordset.Fields("ayamehafte")
Text6.Text = mclass.Recordset.Fields("tshoro")
Text7.Text = mclass.Recordset.Fields("tpayan")
 Text8.Text = mclass.Recordset.Fields("tedadjalasat")
If mclass.Recordset.Fields("sobh") = Am.Caption Then Am.Value = True
If mclass.Recordset.Fields("asr") = Pm.Caption Then Pm.Value = True




End Sub

Private Sub mnuhome_Click()
Entekhab.Show

End Sub

Private Sub mnuselect2select_Click()
For I = 1 To mclass.Recordset.RecordCount
List1.AddItem (mclass.Recordset.Fields("kodclass") & " - " & mclass.Recordset.Fields("maqta") & " - " & mclass.Recordset.Fields("ostad"))
mclass.Recordset.MoveNext
Next I

End Sub

Private Sub mnuxls_Click()
Call Command2_Click
End Sub

Private Sub Ostad_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where ostad like ('%" + Ostad.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub net_Click()
Pc.Checked = False
net.Checked = True
End Sub

Private Sub pc_Click()
Pc.Checked = True
net.Checked = False
End Sub

Private Sub stb1_PanelClick(ByVal Panel As ComctlLib.Panel)
Text6.Text = Me.stb1.Panels(3).Text

End Sub

Private Sub TEP_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + TEP.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub TEP_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Text11_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where dore like ('%" + Text11.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub Text2_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where maqta like ('%" + Text2.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub Text3_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where tozih like ('%" + Text3.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub Text4_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where zamaneshoro like ('%" + Text4.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub Text5_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where zamanepayan like ('%" + Text5.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub Text6_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where tshoro like ('%" + Text6.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub Text7_Change()
If JostojoCh.Value = 1 Then
mclass.Refresh
mclass.RecordSource = "select * from mclass where tpayan like ('%" + Text7.Text + "%')"
mclass.Refresh
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
