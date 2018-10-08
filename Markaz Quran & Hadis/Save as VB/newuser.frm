VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ozviat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " €ÌÌ— ò·„Â ⁄»Ê—"
   ClientHeight    =   3585
   ClientLeft      =   8100
   ClientTop       =   3435
   ClientWidth     =   4905
   Icon            =   "newuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6120
      Picture         =   "newuser.frx":08CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   2655
      Left            =   6480
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   5295
      Begin MSAdodcLib.Adodc Qeybat 
         Height          =   330
         Left            =   2760
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
         Connect         =   $"newuser.frx":0E0C
         OLEDBString     =   $"newuser.frx":0E95
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from qeybat"
         Caption         =   "Qeybat"
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
      Begin MSAdodcLib.Adodc Emtahan 
         Height          =   330
         Left            =   360
         Top             =   2160
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
         Connect         =   $"newuser.frx":0F1E
         OLEDBString     =   $"newuser.frx":0FA7
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from emtahan"
         Caption         =   "Emtahan"
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
         Connect         =   $"newuser.frx":1030
         OLEDBString     =   $"newuser.frx":10B9
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from stu2class"
         Caption         =   "STU2CLASS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
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
         Connect         =   $"newuser.frx":1142
         OLEDBString     =   $"newuser.frx":11CB
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mclass"
         Caption         =   "mclass"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Connect         =   $"newuser.frx":1254
         OLEDBString     =   $"newuser.frx":12DD
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from teacher"
         Caption         =   "teacher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
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
         Connect         =   $"newuser.frx":1366
         OLEDBString     =   $"newuser.frx":13EF
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from student"
         Caption         =   "Student"
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
         Connect         =   $"newuser.frx":1478
         OLEDBString     =   $"newuser.frx":1501
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tarhha"
         Caption         =   "Tarhha"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Setting 
         Height          =   330
         Left            =   2760
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
         Connect         =   $"newuser.frx":158A
         OLEDBString     =   $"newuser.frx":1613
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
         Left            =   2760
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
         Connect         =   $"newuser.frx":169C
         OLEDBString     =   $"newuser.frx":1725
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
      Begin MSAdodcLib.Adodc DataUser 
         Height          =   330
         Left            =   2760
         Top             =   360
         Width           =   2040
         _ExtentX        =   3598
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
         Connect         =   $"newuser.frx":17AE
         OLEDBString     =   $"newuser.frx":1837
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from User1"
         Caption         =   "DataUser"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc bakhshhatable 
         Height          =   330
         Left            =   2760
         Top             =   720
         Width           =   2040
         _ExtentX        =   3598
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
         Connect         =   $"newuser.frx":18C0
         OLEDBString     =   $"newuser.frx":1949
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from bakhshhatable"
         Caption         =   "bakhshhatable"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   2500
      End
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "newuser.frx":19D2
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         Top             =   1920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "newuser.frx":1F14
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         Top             =   1440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   2500
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0080FF80&
         Caption         =   "À» "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2500
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1800
         Width           =   2500
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   480
         TabIndex        =   21
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "ò·„Â ⁄»Ê— ›⁄·Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3165
         TabIndex        =   29
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ ò«—»—Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   3165
         TabIndex        =   28
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "ò·„Â ⁄»Ê— ÃœÌœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3165
         TabIndex        =   27
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   " ﬂ—«— ﬂ·„Â ⁄»Ê— ÃœÌœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   26
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "«Å—« Ê—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   345
         Left            =   3120
         TabIndex        =   25
         Top             =   2400
         Width           =   405
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6000
      TabIndex        =   11
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   10
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9000
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      BOFAction       =   1
      EOFAction       =   2
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"newuser.frx":2456
      OLEDBString     =   $"newuser.frx":24DF
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from user1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "·€Ê"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "⁄÷ÊÌ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5640
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6000
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "«Å—« Ê—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   8280
      TabIndex        =   19
      Top             =   7560
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      Visible         =   0   'False
      X1              =   5520
      X2              =   9960
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ﬂ·„Â ⁄»Ê— „œÌ—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   8325
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "‰«„ „œÌ—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   8280
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   " ﬂ—«— ﬂ·„Â ⁄»Ê—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8325
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ﬂ·„Â ⁄»Ê—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8325
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "‰«„ ﬂ«—»—Ì "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   8325
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "»Â ”«„«‰Â ⁄÷ÊÌ  œ— ‰—„ «›“«— ŒÊ‘ ¬„œÌœ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   3600
   End
End
Attribute VB_Name = "ozviat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'If Entekhab.SB.Panels(1).Text = "„ÌÂ„«‰" Then
Exit Sub

'lllllllllllllllllllllllllllllllllllllllllllllllllllllll
If Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text6.Text = "»—‰«„Â ‰ÊÌ”" Or Combo1.Text = "" Then
Exit Sub
End If


If Text4 = "qoM" And Text5 = "251" Then
If Text2 = Text3 Then

'›ﬁÿ ‰«„ ﬂ«—»—Ì Ê ﬂ·„Â ⁄»Ê— —« »Â œÌ « »Ì” «÷«›Â „Ì ﬂ‰œ
'‘—Ê⁄
Adodc1.Refresh
Adodc1.Recordset.AddNew
'Adodc1.Recordset.Fields("id") = "1"
Adodc1.Recordset.Fields("username") = Me.Text1.Text
Adodc1.Recordset.Fields("password") = Me.Text2.Text
Adodc1.Recordset.Fields("pu") = Me.Combo1.Text
Adodc1.Recordset.Fields("op") = Me.Text6.Text
Adodc1.Recordset.Update
Adodc1.Refresh
MsgBox "⁄„·Ì«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbInformation, "⁄÷ÊÌ  œ— ‰—„ «›“«—"
user.Show
ozviat.Hide

'Å«Ì«‰
Else 'ÿ—› —„“ œÊ„ —« «‘ »«Â Ê«—œ ﬂ—œÂ «” 
MsgBox "ﬂ·„Â ⁄»Ê— «‘ »«Â «” ", vbInformation, "Œÿ«"
End If

Else 'Ì⁄‰Ì ÿ—› ‰«„ „œÌ— —« Ê«—œ ‰ﬂ—œÂ «”  Ì« «‘ »«Â Ê«—œ ﬂ—œÂ «” 
MsgBox "‰«„ „œÌ— Ì« ﬂ·„Â ⁄»Ê— „œÌ— «‘ »«Â «” ", vbCritical, "Œÿ«"
End If




End Sub

Private Sub Command11_Click()
If Picture1.Visible = False Or Picture8.Visible = False Then Exit Sub



DataUser.Recordset.Fields("password") = Text8.Text
DataUser.Recordset.Update
MsgBox " €ÌÌ— ò·„Â ⁄»Ê— »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbInformation, " €ÌÌ— ò·„Â ⁄»Ê—"





End Sub

Private Sub Command2_Click()
ozviat.Hide
user.Show



End Sub

Private Sub Form_Load()
user.Hide
For I = 1 To 10
Combo1.AddItem (I)
Next I

End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show


End Sub

Private Sub Text1_Change()
DataUser.Refresh
DataUser.RecordSource = "select * from user1 where userid like ('%" & Text1.Text & "%')"
DataUser.Refresh

Text11.Text = DataUser.Recordset.Fields("USERNAME")
Text10.Text = DataUser.Recordset.Fields("OP")








End Sub

Private Sub Text7_Change()
If Text7.Text = DataUser.Recordset.Fields("password") Then
Picture1.Visible = True
Else
Picture1.Visible = False
End If





End Sub

Private Sub Text9_Change()
If Text9.Text = Text8.Text Then
Picture8.Visible = True
Picture7.Visible = True


Else
Picture8.Visible = False
Picture7.Visible = False


End If
End Sub
