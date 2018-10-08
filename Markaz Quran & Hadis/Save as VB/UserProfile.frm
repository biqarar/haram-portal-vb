VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form UserProfile 
   Caption         =   " ‰ŸÌ„«  ò«»—"
   ClientHeight    =   8760
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UserProfile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      TabIndex        =   46
      Top             =   120
      Width           =   5145
   End
   Begin VB.TextBox bedoonepassword 
      Height          =   615
      Left            =   6120
      TabIndex        =   45
      Top             =   9120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0080C0FF&
      Caption         =   "»«“Ì«»Ì ò·„Â ⁄»Ê—"
      Height          =   420
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000000FF&
      Caption         =   "Õ–› ò«»—»—"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H000000FF&
      Caption         =   "«÷«›Â ò—œ‰ Å—Ê›«Ì·"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "À»  œò„Â"
      Height          =   615
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   8880
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
         Connect         =   $"UserProfile.frx":08CA
         OLEDBString     =   $"UserProfile.frx":0953
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
         Connect         =   $"UserProfile.frx":09DC
         OLEDBString     =   $"UserProfile.frx":0A65
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
         Connect         =   $"UserProfile.frx":0AEE
         OLEDBString     =   $"UserProfile.frx":0B77
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
         Connect         =   $"UserProfile.frx":0C00
         OLEDBString     =   $"UserProfile.frx":0C89
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
         Connect         =   $"UserProfile.frx":0D12
         OLEDBString     =   $"UserProfile.frx":0D9B
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
         Connect         =   $"UserProfile.frx":0E24
         OLEDBString     =   $"UserProfile.frx":0EAD
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
         Connect         =   $"UserProfile.frx":0F36
         OLEDBString     =   $"UserProfile.frx":0FBF
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
         Connect         =   $"UserProfile.frx":1048
         OLEDBString     =   $"UserProfile.frx":10D1
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
         Connect         =   $"UserProfile.frx":115A
         OLEDBString     =   $"UserProfile.frx":11E3
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
         Connect         =   $"UserProfile.frx":126C
         OLEDBString     =   $"UserProfile.frx":12F5
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
         Connect         =   $"UserProfile.frx":137E
         OLEDBString     =   $"UserProfile.frx":1407
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
   Begin VB.Frame Frame1 
      Caption         =   "ò«—»— ÃœÌœ"
      Height          =   3855
      Left            =   6240
      TabIndex        =   29
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— ›⁄«·"
         Height          =   300
         Left            =   480
         TabIndex        =   6
         Top             =   2880
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "›⁄«·"
         Height          =   300
         Left            =   1800
         TabIndex        =   5
         Top             =   2880
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "UserProfile.frx":1490
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox Picture5 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "UserProfile.frx":19D2
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   38
         Top             =   2400
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "UserProfile.frx":1F14
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
         Top             =   1920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "UserProfile.frx":2456
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ã«Ìê“Ì‰Ì «ÿ·«⁄« "
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   34
         Top             =   360
         Width           =   2500
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3360
         Width           =   1575
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
         Height          =   435
         Left            =   480
         TabIndex        =   1
         Top             =   840
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
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
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
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1800
         Width           =   2500
      End
      Begin VB.TextBox Text6 
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
         TabIndex        =   4
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò«—»—"
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
         Index           =   2
         Left            =   3165
         TabIndex        =   35
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ ﬂ«—»—Ì "
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
         Index           =   1
         Left            =   3165
         TabIndex        =   33
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ﬂ·„Â ⁄»Ê—"
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
         TabIndex        =   32
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " ﬂ—«— ﬂ·„Â ⁄»Ê—"
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
         TabIndex        =   31
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label7 
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
         TabIndex        =   30
         Top             =   2400
         Width           =   405
      End
   End
   Begin VB.CheckBox JostojoCh 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ã” ÃÊ"
      Height          =   345
      Left            =   4560
      TabIndex        =   22
      Top             =   9960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "À»  ò·«” ÃœÌœ"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   735
      Left            =   1800
      TabIndex        =   16
      Top             =   3720
      Width           =   4335
      Begin VB.Label ltarh 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "op"
         DataSource      =   "DataUser"
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
         TabIndex        =   20
         Top             =   240
         Width           =   135
      End
      Begin VB.Label labeluserid 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "userid"
         DataSource      =   "DataUser"
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
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "«Å—« Ê—"
         Height          =   300
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò«»—Ì"
         Height          =   300
         Index           =   0
         Left            =   3720
         TabIndex        =   17
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ã«Ìê“Ì‰Ì «ÿ·«⁄« "
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000002&
      Height          =   6360
      ItemData        =   "UserProfile.frx":2998
      Left            =   10680
      List            =   "UserProfile.frx":299A
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   9600
      Picture         =   "UserProfile.frx":299C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "ÃÂ  «÷«›Â ò—œ‰ —òÊ—œ «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   5400
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
      Left            =   9600
      Picture         =   "UserProfile.frx":63F9
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "—òÊ—œÌ òÂ „ÌŒÊ«ÂÌœ Õ–› ‰„«ÌÌœ —« «‰ Œ«» Ê «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   6360
      Width           =   855
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
      Left            =   6240
      TabIndex        =   14
      Text            =   "«‰ Œ«» ò‰Ìœ"
      Top             =   3960
      Width           =   3255
   End
   Begin ComctlLib.ProgressBar PB1 
      Height          =   135
      Left            =   11640
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DMClass 
      Bindings        =   "UserProfile.frx":A9BE
      Height          =   4095
      Left            =   120
      TabIndex        =   23
      Top             =   4560
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7223
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
      Caption         =   "„Ì“«‰ œ”—”Ì ò«—»—"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "userid"
         Caption         =   "userid"
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
         DataField       =   "commandname"
         Caption         =   "commandname"
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
         DataField       =   "status"
         Caption         =   "Ê÷⁄Ì "
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
         DataField       =   "farsiname"
         Caption         =   "‰«„ ⁄„·Ì« "
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
         DataField       =   "goroh"
         Caption         =   "ê—ÊÂ"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3630.047
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3764.977
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar PB2 
      Height          =   135
      Left            =   11640
      TabIndex        =   24
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "UserProfile.frx":A9DD
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5318
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
      Caption         =   "·Ì”  ò«»—«‰"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "username"
         Caption         =   "‰«„ ò«—»—Ì"
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
         DataField       =   "password"
         Caption         =   "password"
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
         DataField       =   "pu"
         Caption         =   "pu"
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
         DataField       =   "op"
         Caption         =   "‰«„ «Å—« Ê—"
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
         DataField       =   "userid"
         Caption         =   "òœ ò«—»—Ì"
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
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1620.284
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "UserProfile.frx":A9F4
      Height          =   2055
      Left            =   6600
      TabIndex        =   41
      Top             =   9720
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
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
      Caption         =   "œò„Â Â« Ê œ” Ê— Â«Ì „ÊÃÊœ œ— ‰—„ «›“«—"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "name"
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
         DataField       =   "bakhshkod"
         Caption         =   "bakhshkod"
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
         DataField       =   "goroh"
         Caption         =   "goroh"
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
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ã” ÃÊ"
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
      Left            =   5445
      TabIndex        =   47
      Top             =   120
      Width           =   465
   End
   Begin VB.Label LJostojo 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "œ— Â—òœ«„ «“ „Ê«—œ »«·« ò·„Â «Ì »‰ÊÌ”Ìœ  « Ã” ÃÊ ‘Êœ"
      Height          =   345
      Left            =   480
      TabIndex        =   28
      Top             =   10080
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ê—ÊÂ Â«"
      Height          =   300
      Left            =   9840
      TabIndex        =   27
      Top             =   3960
      Width           =   420
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "»Œ‘ Â«Ì ‰—„ «›“«—"
      Height          =   300
      Left            =   12720
      TabIndex        =   26
      Top             =   120
      Width           =   1110
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
      Left            =   11640
      TabIndex        =   25
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnudelluser 
      Caption         =   "„œÌ—Ì "
      Begin VB.Menu mnuprofile 
         Caption         =   " ‰ŸÌ„ „Ì“«‰ œ” —”Ì"
      End
      Begin VB.Menu mnuedit 
         Caption         =   " €ÌÌ— „‘Œ’«  ò«—»—"
      End
      Begin VB.Menu dfgdsfg 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Õ–› ò«—»—"
      End
   End
End
Attribute VB_Name = "UserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where goroh like ('" & Combo1.Text & "') and userid like ('%" & labeluserid.Caption & "%')"
userprofiletable.Refresh
End Sub

Private Sub Command10_Click()
On Error GoTo 9898
GoTo 9999
9898:
MsgBox "„Ê—œÌ «‰ Œ«» ‰‘œÂ «” ", vbCritical + vbOKOnly, "Œÿ«"

Exit Sub
9999:
'userprofiletable.Refresh
'userprofiletable.RecordSource = "select * from userprofiletable where farsiname like ('" & List1.Text & "') and userid like ('%" & labeluserid.Caption & "%')"
'userprofiletable.Refresh

userprofiletable.Recordset.Fields("status") = "off"
userprofiletable.Recordset.Update


userprofiletable.Recordset.MoveNext

End Sub

Private Sub Command2_Click()
'If Entekhab.SB.Panels(1).Text = "„ÌÂ„«‰" Then
If Text6.Text = "»—‰«„Â ‰ÊÌ”" Then Exit Sub

If Picture2.Visible = False Or Picture5.Visible = False Then Exit Sub

'DataUser.Refresh
'DataUser.RecordSource = "select * from datauser where userid like ('%" & "" & "%')"
'DataUser.Refresh
'DataUser.Recordset.Sort = "userid"
' DataUser.Recordset.MovePrevious
 
' DataUser.Recordset.MoveLast
'Dim datauSEridin As String
 
 'datauSEridin = DataUser.Recordset.Fields("userid")
'lllllllllllllllllllllllllllllllllllllllllllllllllllllll
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text6.Text = "" Or Text6.Text = "»—‰«„Â ‰ÊÌ”" Then
Exit Sub
End If



If Text2 = Text3 Then

'›ﬁÿ ‰«„ ﬂ«—»—Ì Ê ﬂ·„Â ⁄»Ê— —« »Â œÌ « »Ì” «÷«›Â „Ì ﬂ‰œ
'‘—Ê⁄
DataUser.Refresh
DataUser.Recordset.AddNew
DataUser.Recordset.Fields("userid") = "RMUBQ-" & Text1.Text
DataUser.Recordset.Fields("username") = Me.Text1.Text
DataUser.Recordset.Fields("password") = Me.Text2.Text
'DataUser.Recordset.Fields("pu") = Me.Combo1.Text
DataUser.Recordset.Fields("op") = Me.Text6.Text
DataUser.Recordset.Update
DataUser.Refresh
Text4.Text = "RMUBQ-" & Text1.Text
Call Command6_Click

MsgBox "⁄„·Ì«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbInformation, "⁄÷ÊÌ  œ— ‰—„ «›“«—"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text6.Text = ""
Text4.Text = ""

'Å«Ì«‰
Else 'ÿ—› —„“ œÊ„ —« «‘ »«Â Ê«—œ ﬂ—œÂ «” 
MsgBox "ﬂ·„Â ⁄»Ê— «‘ »«Â «” ", vbInformation, "Œÿ«"
End If



End Sub

Private Sub Command5_Click()
bakhshhatable.Refresh
bakhshhatable.Recordset.AddNew

bakhshhatable.Recordset.Fields("name") = Text1.Text
bakhshhatable.Recordset.Fields("bakhshkod") = Text2.Text
bakhshhatable.Recordset.Fields("goroh") = Text3.Text
bakhshhatable.Recordset.Fields("sort") = Text6.Text

bakhshhatable.Recordset.Update
 bakhshhatable.Refresh
 


End Sub

Private Sub Command6_Click()

bakhshhatable.Refresh
bakhshhatable.RecordSource = "select * from bakhshhatable where name like ('%" & "" & "%')"
bakhshhatable.Refresh

For I = 1 To bakhshhatable.Recordset.RecordCount




userprofiletable.Refresh
userprofiletable.Recordset.AddNew

userprofiletable.Recordset.Fields("userid") = "RMUBQ-" & Text1.Text

userprofiletable.Recordset.Fields("commandname") = bakhshhatable.Recordset.Fields("bakhshkod")
userprofiletable.Recordset.Fields("farsiname") = bakhshhatable.Recordset.Fields("name")
userprofiletable.Recordset.Fields("goroh") = bakhshhatable.Recordset.Fields("goroh")

 If bakhshhatable.Recordset.Fields("bakhshkod") = "modiriyat" Then GoTo 1
 
If Option1.Value = True Then userprofiletable.Recordset.Fields("status") = "on"

If Option2.Value = True Then
1

userprofiletable.Recordset.Fields("status") = "off"

End If

userprofiletable.Recordset.Update
userprofiletable.Refresh



bakhshhatable.Recordset.MoveNext

Next I

End Sub

Private Sub Command7_Click()


On Error GoTo 9898
GoTo 9999
9898:
MsgBox "„Ê—œÌ «‰ Œ«» ‰‘œÂ «” ", vbCritical + vbOKOnly, "Œÿ«"

Exit Sub
9999:
'userprofiletable.Refresh
'userprofiletable.RecordSource = "select * from userprofiletable where farsiname like ('" & List1.Text & "') and userid like ('%" & labeluserid.Caption & "%')"
'userprofiletable.Refresh

userprofiletable.Recordset.Fields("status") = "on"
userprofiletable.Recordset.Update
userprofiletable.Recordset.MoveNext



End Sub

Private Sub Command8_Click()
DataUser.Recordset.Delete
End Sub

Private Sub Command9_Click()


If MsgBox("»« «‰ Œ«» «Ì‰ ⁄„·Ì«  ”Ì” „ ò·„Â ⁄»Ê— ÃœÌœÌ »—«Ì «Ì‰ ò«—»— œ— ‰Ÿ— „Ì êÌ—œ" & Chr(10) & "¬Ì« „Ì ŒÊ«ÂÌœ «œ«„Â œÂÌœ", vbQuestion + vbYesNo, "") = vbYes Then

DataUser.Refresh
DataUser.RecordSource = "select * from user1 where userid like ('" & labeluserid.Caption & "')"
DataUser.Refresh


X = Int(Rnd(1) * 251 * 100)
X = Int(Rnd(1) * 251 * 100)
X = Int(Rnd(1) * 251 * 100)
X = Int(Rnd(1) * 251 * 100)
X = Int(Rnd(1) * 251 * 100)


DataUser.Recordset.Fields("password") = X
DataUser.Recordset.Update
MsgBox "ò·„Â ⁄»Ê— ÃœÌœ" & Chr(10) & X, vbInformation, "»«“Ì«»Ì ò·„Â ⁄»Ê—"
End If
End Sub

Private Sub Form_Load()
bakhshhatable.Refresh
bakhshhatable.RecordSource = " select * from bakhshhatable where name like ('%" & "" & "%')"
bakhshhatable.Refresh


Dim tekrar As String
tekrar = bakhshhatable.Recordset.Fields("goroh")
bakhshhatable.Recordset.Sort = "sort"
Combo1.AddItem (bakhshhatable.Recordset.Fields("goroh"))
For I = 1 To bakhshhatable.Recordset.RecordCount

List1.AddItem (bakhshhatable.Recordset.Fields("name"))

If tekrar <> bakhshhatable.Recordset.Fields("goroh") Then
tekrar = bakhshhatable.Recordset.Fields("goroh")
Combo1.AddItem (bakhshhatable.Recordset.Fields("goroh"))
End If


bakhshhatable.Recordset.MoveNext
Next I

End Sub

Private Sub Label18_Click()
End Sub

Private Sub Label34_Click()
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Me.bedoonepassword.Text = "password" Then
End
End If


Entekhab.Show

End Sub

Private Sub labeluserid_Change()
userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('%" & labeluserid.Caption & "%')"
userprofiletable.Refresh
userprofiletable.Recordset.Sort = "goroh"

End Sub

Private Sub labeluserid_Click()
userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('%" & labeluserid.Caption & "%')"
userprofiletable.Refresh
End Sub


Private Sub List1_Click()
userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where farsiname like ('" & List1.Text & "') and userid like ('%" & labeluserid.Caption & "%')"
userprofiletable.Refresh
End Sub

Private Sub mnudelete_Click()

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ò«—»— —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–› ò«—»—") = vbYes Then
'bakhshhatable.Refresh
'bakhshhatable.RecordSource = " select * from bakhshhatable where name like ('%" & "" & "%')"
'bakhshhatable.Refresh
If labeluserid.Caption = "" Then Exit Sub

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('%" & labeluserid.Caption & "%')"
userprofiletable.Refresh




For I = 1 To userprofiletable.Recordset.RecordCount

On Error Resume Next

userprofiletable.Recordset.Delete
userprofiletable.Recordset.MoveNext

Next I

DataUser.Recordset.Delete
End If


End Sub

Private Sub Text1_Change()
DataUser.Refresh
DataUser.RecordSource = "select * from user1 where username like ('%" & Text1.Text & "%')"
DataUser.Refresh
If DataUser.Recordset.BOF = True Or DataUser.Recordset.EOF = True Then
Picture2.Visible = True
Else
 Picture2.Visible = False
 End If
 

End Sub


Private Sub Text2_Change()
If Text2.Text = "" Or Text3.Text = "" Then
Picture4.Visible = False
Picture3.Visible = False
Exit Sub
End If
End Sub

Private Sub Text3_Change()
If Text2.Text = "" Or Text3.Text = "" Then
Picture4.Visible = False
Picture3.Visible = False
Exit Sub
End If


If Text2.Text = Text3.Text Then
Picture4.Visible = True
Picture3.Visible = True


Else
Picture4.Visible = False
Picture3.Visible = False


End If

End Sub

Private Sub Text5_Change()
DataUser.Refresh
DataUser.RecordSource = "select * from user1 where username like ('%" & Text5.Text & "%') or op like ('%" & Text5.Text & "%') or userid like ('%" & Text5.Text & "%')"
DataUser.Refresh
End Sub

Private Sub Text6_Change()
DataUser.Refresh
DataUser.RecordSource = "select * from user1 where op like ('%" & Text6.Text & "%')"
DataUser.Refresh
If DataUser.Recordset.BOF = True Or DataUser.Recordset.EOF = True Then
Picture5.Visible = True
Else
 Picture5.Visible = False
 End If
 
End Sub


