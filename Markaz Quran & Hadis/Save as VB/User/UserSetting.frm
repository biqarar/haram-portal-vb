VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SettingF 
   Caption         =   " ‰ŸÌ„«  ò«—»—"
   ClientHeight    =   9105
   ClientLeft      =   9090
   ClientTop       =   7935
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UserSetting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   15960
   Begin VB.Frame Frame5 
      Caption         =   "»«“Ì«»Ì ò·«”  „«„ ‘œÂ"
      Height          =   2055
      Left            =   12240
      TabIndex        =   52
      Top             =   8280
      Width           =   2535
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "‘—Ê⁄ ⁄„·Ì«  »«“Ì«»Ì"
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   435
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
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
         Left            =   1800
         TabIndex        =   55
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "À»  »Œ‘ Â«Ì ‰—„ «›“«—"
      Height          =   2055
      Left            =   8880
      TabIndex        =   42
      Top             =   3480
      Width           =   2535
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
         Left            =   120
         Picture         =   "UserSetting.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "ÃÂ  «÷«›Â ò—œ‰ —òÊ—œ «Ì‰ œò„Â —« »“‰Ìœ"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox ttT 
         Height          =   435
         Left            =   960
         TabIndex        =   46
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "À» "
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "¬œ—”"
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
         Left            =   1800
         TabIndex        =   44
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "À»  ‰ ŸÌ„« "
      Height          =   1815
      Left            =   120
      TabIndex        =   26
      Top             =   3120
      Width           =   8655
      Begin VB.ComboBox Combo4 
         Height          =   435
         Left            =   4680
         TabIndex        =   41
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ã«Ìê“Ì‰Ì"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÃœÌœ"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H0080FF80&
         Caption         =   "À» "
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Õ–›"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox TE5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   840
         TabIndex        =   30
         Top             =   1320
         Width           =   6600
      End
      Begin VB.TextBox TE4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   840
         Width           =   6600
      End
      Begin VB.TextBox TE2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   3360
         TabIndex        =   28
         Text            =   "101"
         Top             =   360
         Width           =   720
      End
      Begin VB.ComboBox Combo3 
         Height          =   435
         Left            =   840
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
         Height          =   315
         Left            =   7560
         TabIndex        =   39
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "òœ"
         Height          =   315
         Left            =   7680
         TabIndex        =   37
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â"
         Height          =   315
         Left            =   4200
         TabIndex        =   36
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "⁄‰Ê«‰"
         Height          =   315
         Left            =   2880
         TabIndex        =   35
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "„ ‰"
         Height          =   315
         Left            =   7680
         TabIndex        =   34
         Top             =   840
         Width           =   240
      End
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
      Left            =   8880
      Picture         =   "UserSetting.frx":4327
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "—òÊ—œÌ òÂ „ÌŒÊ«ÂÌœ Õ–› ‰„«ÌÌœ —« «‰ Œ«» Ê «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   9480
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   " ‰ŸÌ„ ”«· Ã«—Ì"
      Height          =   1815
      Left            =   8880
      TabIndex        =   19
      Top             =   120
      Width           =   2655
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
         ItemData        =   "UserSetting.frx":88EC
         Left            =   840
         List            =   "UserSetting.frx":88EE
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1320
         Width           =   855
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
         ItemData        =   "UserSetting.frx":88F0
         Left            =   840
         List            =   "UserSetting.frx":88F2
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "À» "
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   615
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
         ItemData        =   "UserSetting.frx":88F4
         Left            =   120
         List            =   "UserSetting.frx":88F6
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "„«Â"
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
         Left            =   1800
         TabIndex        =   50
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "òœ »Œ‘"
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
         Left            =   1800
         TabIndex        =   24
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "”«· Ã«—Ì"
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
         Left            =   1800
         TabIndex        =   22
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.CommandButton Command9 
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
      Left            =   10200
      Picture         =   "UserSetting.frx":88F8
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "»—«Ì Å—‘ »Â —òÊ—œ »⁄œÌ «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command8 
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
      Left            =   9240
      Picture         =   "UserSetting.frx":C988
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "»—«Ì Å—‘ »Â —òÊ—œ ﬁ»·Ì «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command6 
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
      Left            =   10440
      Picture         =   "UserSetting.frx":108A9
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "«›“Êœ‰ Ê–ŒÌ—Â ò—œ‰"
      Top             =   6360
      Width           =   855
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
      Height          =   4455
      Left            =   11520
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   4215
      Begin MSAdodcLib.Adodc vadie 
         Height          =   330
         Left            =   2400
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
         Connect         =   $"UserSetting.frx":145F9
         OLEDBString     =   $"UserSetting.frx":14682
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
         Left            =   2400
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
         Connect         =   $"UserSetting.frx":1470B
         OLEDBString     =   $"UserSetting.frx":14794
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
         Width           =   2055
         _ExtentX        =   3625
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
         Connect         =   $"UserSetting.frx":1481D
         OLEDBString     =   $"UserSetting.frx":148A6
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
         Width           =   2055
         _ExtentX        =   3625
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
         Connect         =   $"UserSetting.frx":1492F
         OLEDBString     =   $"UserSetting.frx":149B8
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
         Width           =   2055
         _ExtentX        =   3625
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
         Connect         =   $"UserSetting.frx":14A41
         OLEDBString     =   $"UserSetting.frx":14ACA
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
         Connect         =   $"UserSetting.frx":14B53
         OLEDBString     =   $"UserSetting.frx":14BDC
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
         Width           =   2055
         _ExtentX        =   3625
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
         Connect         =   $"UserSetting.frx":14C65
         OLEDBString     =   $"UserSetting.frx":14CEE
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
         Left            =   2400
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
         Connect         =   $"UserSetting.frx":14D77
         OLEDBString     =   $"UserSetting.frx":14E00
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
         Left            =   2400
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
         Connect         =   $"UserSetting.frx":14E89
         OLEDBString     =   $"UserSetting.frx":14F12
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
      Begin MSAdodcLib.Adodc SettingUser 
         Height          =   330
         Left            =   2400
         Top             =   1800
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
         Connect         =   $"UserSetting.frx":14F9B
         OLEDBString     =   $"UserSetting.frx":15024
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
         Left            =   360
         Top             =   2640
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
         Connect         =   $"UserSetting.frx":150AD
         OLEDBString     =   $"UserSetting.frx":15136
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
         Left            =   480
         Top             =   3000
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
         Connect         =   $"UserSetting.frx":151BF
         OLEDBString     =   $"UserSetting.frx":15248
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
         Left            =   480
         Top             =   3360
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
         Connect         =   $"UserSetting.frx":152D1
         OLEDBString     =   $"UserSetting.frx":1535A
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
      Begin MSAdodcLib.Adodc paziresh_table 
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
         Connect         =   $"UserSetting.frx":153E3
         OLEDBString     =   $"UserSetting.frx":1546C
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from paziresh_table"
         Caption         =   "paziresh_table"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "«ÿ·«⁄«  ›«Ì· Œ—ÊÃÌ"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«’·«Õ"
         Height          =   375
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«’·«Õ"
         Height          =   375
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«’·«Õ"
         Height          =   375
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«’·«Õ"
         Height          =   375
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»⁄œ «“  €ÌÌ— «ÿ·«⁄«  ‰—„ «›“«— —« œÊ»«—Â —«Â  «‰œ«“Ì ò‰Ìœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4080
         TabIndex        =   15
         Top             =   2400
         Width           =   4320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nerwork Scan Adress"
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
         TabIndex        =   14
         Top             =   1800
         Width           =   2130
      End
      Begin VB.Label NetScanAdress 
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
         Left            =   2640
         TabIndex        =   13
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scan Adress"
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
         TabIndex        =   11
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label ScanAdress 
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
         Left            =   2640
         TabIndex        =   10
         Top             =   1320
         Width           =   225
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
         Left            =   2640
         TabIndex        =   7
         Top             =   840
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
         TabIndex        =   6
         Top             =   840
         Width           =   855
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
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   225
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
         TabIndex        =   4
         Top             =   360
         Width           =   750
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "UserSetting.frx":154F5
      Height          =   2895
      Left            =   120
      TabIndex        =   38
      Top             =   5040
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   12632319
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "XkodSetting"
         Caption         =   "òœ ”Ì” „"
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
         DataField       =   "XSort"
         Caption         =   "‘„«—Â"
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
         DataField       =   "XName"
         Caption         =   "⁄‰Ê«‰"
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
         DataField       =   "XText"
         Caption         =   "„ ‰"
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
         DataField       =   "Tozih"
         Caption         =   " Ê÷ÌÕ"
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
         DataField       =   "Op"
         Caption         =   "Op"
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
         DataField       =   "D"
         Caption         =   "D"
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
         DataField       =   "Other1"
         Caption         =   "Other1"
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
         DataField       =   "Other2"
         Caption         =   "Other2"
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
         DataField       =   "Other3"
         Caption         =   "Other3"
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
         DataField       =   "Other4"
         Caption         =   "Other4"
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
         DataField       =   "Other5"
         Caption         =   "Other5"
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
            Locked          =   -1  'True
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "UserSetting.frx":1550B
      Height          =   2895
      Left            =   120
      TabIndex        =   47
      Top             =   8040
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   12632319
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
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
      ColumnCount     =   4
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
      BeginProperty Column03 
         DataField       =   "sort"
         Caption         =   "sort"
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
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "UserSetting.frx":15527
      Height          =   2895
      Left            =   11640
      TabIndex        =   51
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   12632319
      HeadLines       =   1
      RowHeight       =   27
      AllowDelete     =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "»—«Ì Ê—Êœ ò·Ìò ò‰Ìœ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Left            =   8880
      TabIndex        =   0
      Top             =   7200
      Width           =   2715
   End
End
Attribute VB_Name = "SettingF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo4_Click()
TE2.Text = "101"
End Sub

Private Sub Command1_Click()
Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "ScanAdressJPG" & "%') "
Setting.Refresh

'Me.AdressLabel.Caption = Setting.Recordset.Fields("xtext")

 Setting.Recordset.Fields("xtext") = InputBox("·ÿ›« ¬œ” ›Ê·œ— »—‰«„Â —« Ê«—œ ò‰Ìœ", "Adress")
Setting.Recordset.Update

Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "ScanAdressJPG" & "%') "
Setting.Refresh

Me.ScanAdress.Caption = Setting.Recordset.Fields("xtext")

End Sub

Private Sub Command10_Click()
bakhshhatable.Recordset.Delete

End Sub

Private Sub Command11_Click()
Beep

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ ”«· Ã«—Ì —« »Â ”«· " & Combo1.Text & " €ÌÌ— œÂÌœ", vbQuestion + vbYesNo, " €ÌÌ— ”«· Ã«—Ì") = vbYes Then

SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "SalJari" + "%')"
SettingUser.Refresh
'MsgBox Right(Combo1.Text, 2)
'Exit Sub


'SettingUser.Refresh
'SettingUser.Recordset.AddNew
'SettingUser.Recordset.Fields("Xcode") = "SalJari"
'SettingUser.Recordset.Fields("status") = "On"
SettingUser.Recordset.Fields("value") = Right(Combo1.Text, 2)
SettingUser.Recordset.Fields("Xtext") = Combo5.Text

SettingUser.Recordset.Update
SettingUser.Refresh
MsgBox " ‰ŸÌ„ ”«· Ã«—Ì »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbInformation + vbOKOnly, Combo1.Text



'„—»Êÿ »Â »Œ‘ òœ »Œ‘
'If MsgBox("¬Ì« ‰”»  »Â òœ " & Combo1.Text & " €ÌÌ— œÂÌœ", vbQuestion + vbYesNo, " €ÌÌ— ”«· Ã«—Ì") = vbYes Then

SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "CodeBakhsh" + "%')"
SettingUser.Refresh
'MsgBox Right(Combo1.Text, 2)
'Exit Sub


'SettingUser.Refresh
'SettingUser.Recordset.AddNew
'SettingUser.Recordset.Fields("Xcode") = "CodeBakhsh"
'SettingUser.Recordset.Fields("status") = "On"
SettingUser.Recordset.Fields("value") = Combo2.Text
SettingUser.Recordset.Fields("Xtext") = Combo2.Text
SettingUser.Recordset.Update
SettingUser.Refresh
MsgBox " ‰ŸÌ„ òœ »Œ‘ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbInformation + vbOKOnly, Combo2.Text


'End If
'„—»Êÿ »Â »Œ‘ òœ »Œ‘


End If

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


End Sub

Private Sub Command15_Click()
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «ÿ·«⁄«  À»  ‘Êœ", vbQuestion + vbYesNo, "À»  «ÿ·«⁄« ") = vbNo Then Exit Sub



Setting.Refresh
Setting.Recordset.AddNew
Setting.Recordset.Fields("xkodsetting") = Combo4.Text
Setting.Recordset.Fields("xsort") = TE2.Text
Setting.Recordset.Fields("xname") = Combo3.Text
Setting.Recordset.Fields("xtext") = TE4.Text
Setting.Recordset.Fields("tozih") = TE5.Text
Setting.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
'Setting.Recordset.Fields("d") = "@"
Setting.Recordset.Fields("other1") = "@"
Setting.Recordset.Fields("other2") = "@"
Setting.Recordset.Fields("other3") = "@"
Setting.Recordset.Fields("other4") = "@"
Setting.Recordset.Fields("other5") = "@"
Setting.Recordset.Update
Setting.Refresh
MsgBox "«ÿ·«⁄«  À»  ‘œ", vbInformation + vbOKOnly, "À»  «ÿ·«⁄« "

End Sub

Private Sub Command16_Click()
On Error Resume Next
Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + Combo4.Text + "%')"
Setting.Refresh

Setting.Recordset.Sort = "xsort"

Setting.Recordset.MovePrevious

Setting.Recordset.MoveLast

TE2.Text = Val(Setting.Recordset.Fields("xsort")) + 1


Beep

End Sub

Private Sub Command17_Click()
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" & Text3.Text & "%') and elat  like ('%" & "« „«„ ò·«”" & "%') "
STU2CLASS.Refresh

For I = 1 To STU2CLASS.Recordset.RecordCount
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & STU2CLASS.Recordset.Fields("parvande") & "%')"
Student.Refresh
                If Student.Recordset.Fields("clas1") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas1") = Text3.Text
                Else
                If Student.Recordset.Fields("clas2") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas2") = Text3.Text
                Else
                If Student.Recordset.Fields("clas3") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas3") = Text3.Text
                Else
                If Student.Recordset.Fields("clas4") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas4") = Text3.Text
                Else
                If Student.Recordset.Fields("clas5") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas5") = Text3.Text
                Else
                MsgBox STU2CLASS.Recordset.Fields("parvande")
                MsgBox "«Ì‰ ﬁ—¬‰ ¬„Ê“ œ— Õ«· Õ«Ÿ— œ— 5 ò·«” »Â ’Ê—  Â„“„«‰ ‘—ò  „Ì ò‰œ Ê „Ã«“ »Â ‘—ò  œ— ò·«” œÌê—Ì ‰„Ì »«‘œ. ·ÿ›« Ìò Ì« ç‰œ ò·«” «Ì‘«‰ —«Õ–› ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
                Exit Sub
                End If
                End If
                End If
                End If
                End If
                STU2CLASS.Recordset.Fields("tpayan") = ""
                STU2CLASS.Recordset.Fields("elat") = ""
               STU2CLASS.Recordset.Fields("tozih") = ""
               Student.Recordset.Update
                STU2CLASS.Recordset.Update
     STU2CLASS.Recordset.MoveNext
     
     
Next I

mclass.Refresh
mclass.RecordSource = "select  * from mclass where kodclass like ('%" & Text3.Text & "%')"
mclass.Refresh
mclass.Recordset.Fields("tozih") = ""
mclass.Recordset.Update

MsgBox "»«“Ì«»Ì «ÿ·«⁄«  ò·«” «‰Ã«„ ‘œ"
End Sub

Private Sub Command2_Click()
Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "netScanAdressJPG" & "%') "
Setting.Refresh

'Me.AdressLabel.Caption = Setting.Recordset.Fields("xtext")

 Setting.Recordset.Fields("xtext") = InputBox("·ÿ›« ¬œ” ›Ê·œ— »—‰«„Â —« Ê«—œ ò‰Ìœ", "Adress")
Setting.Recordset.Update

Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "netScanAdressJPG" & "%') "
Setting.Refresh

Me.NetScanAdress.Caption = Setting.Recordset.Fields("xtext")

End Sub

Private Sub Command3_Click()
Dim ASD As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD, ParvandeQuranAmooZ As String


bakhshhatable.Refresh


Set oExcel = GetObject(Text1.Text)

For I = 1 To ttT.Text
bakhshhatable.Refresh

bakhshhatable.Recordset.AddNew

bakhshhatable.Recordset.Fields("bakhshkod") = oExcel.ActiveSheet.Range("a" & I).Value
bakhshhatable.Recordset.Fields("name") = oExcel.ActiveSheet.Range("b" & I).Value
bakhshhatable.Recordset.Fields("goroh") = oExcel.ActiveSheet.Range("c" & I).Value
bakhshhatable.Recordset.Fields("sort") = oExcel.ActiveSheet.Range("d" & I).Value
bakhshhatable.Recordset.Update

Next I

DataUser.Refresh
DataUser.RecordSource = "select * from user1 where username like ('%" & "" & "%')"
DataUser.Refresh




MsgBox "ok"

End Sub

Private Sub Command5_Click()
On Error Resume Next
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ „Ê—œ —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–›") = vbYes Then

Setting.Recordset.Delete
End If

End Sub

Private Sub Command6_Click()


SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "SalJari" + "%')"
SettingUser.Refresh

Combo1.Text = SettingUser.Recordset.Fields("xtext")



Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "AdressXlsx" & "%') "
Setting.Refresh

Me.AdressLabel.Caption = Setting.Recordset.Fields("xtext")

Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "netAdressXlsx" & "%') "
Setting.Refresh

Me.NetAdresslabel.Caption = Setting.Recordset.Fields("xtext")



Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "ScanAdressJPG" & "%') "
Setting.Refresh

Me.ScanAdress.Caption = Setting.Recordset.Fields("xtext")



Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "NetScanAdressJPG" & "%') "
Setting.Refresh

Me.NetScanAdress.Caption = Setting.Recordset.Fields("xtext")

End Sub

Private Sub Command7_Click()

DataUser.Refresh
DataUser.RecordSource = "select * from user1 where username like ('%" & "" & "%')"
DataUser.Refresh
userprofiletable.Refresh

For I = 1 To DataUser.Recordset.RecordCount

'userprofiletable.Refresh
userprofiletable.Recordset.AddNew

userprofiletable.Recordset.Fields("userid") = DataUser.Recordset.Fields("userid")

userprofiletable.Recordset.Fields("commandname") = bakhshhatable.Recordset.Fields("bakhshkod")
userprofiletable.Recordset.Fields("farsiname") = bakhshhatable.Recordset.Fields("name")
userprofiletable.Recordset.Fields("goroh") = bakhshhatable.Recordset.Fields("goroh")



userprofiletable.Recordset.Fields("status") = "off"


userprofiletable.Recordset.Update
userprofiletable.Refresh





DataUser.Recordset.MoveNext

Next I
MsgBox "halle"

End Sub

Private Sub Form_Load()
'À»  ”«· Õ«—” »—«Ì «Ì‰„òÂ ”«· Â«Ì «œ ‘Ê‰œ
For I = 1390 To 1408
Combo1.AddItem (I)
Next I

Combo2.AddItem ("")
For I = 11 To 19
Combo2.AddItem (I)
Next I

For I = 1 To 12 Step 1  ' »—«Ì Ê«—œ ﬂ—œ‰ ‘„«—Â —Ê“ œ— ÃœÊ· «ÿ·«⁄«  €Ì 
If I < 10 Then
Combo5.AddItem ("0" & I)
Else
Combo5.AddItem (I)
End If
Next I

Combo4.AddItem ("User-EmtahanF-Momtahen")
Combo4.AddItem ("User-VadieF-DaryaftKonnande")
Combo4.AddItem ("CodeBakhsh")
Combo4.AddItem ("SalJari")
Combo4.AddItem ("AdressXlsx")
Combo4.AddItem ("netAdressXlsx")
Combo4.AddItem ("ScanAdressJPG")
Combo4.AddItem ("NetScanAdressJPG")
Combo4.AddItem ("stu-tahsilat")
Combo4.AddItem ("stu-karshenas")

If user.OP.Text = "»—‰«„Â ‰ÊÌ”" Then
Label1.Caption = "»—«Ì Ê—Êœ ò·Ìò ò‰Ìœ"
 
 Label1.ForeColor = &HC000&

End If
If user.Text3.Text = "RezaMohiti" Then

Call Command6_Click
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show



End Sub

Private Sub Label1_Click()
If user.OP.Text = "»—‰«„Â ‰ÊÌ”" Then
Form12.Show

Unload Me

End If
Beep

End Sub

