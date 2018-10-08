VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form sabt_nomre_omoomi 
   Caption         =   "À»  ‰„—«  œÊ—Â Â«Ì ﬁ—«∆ "
   ClientHeight    =   8475
   ClientLeft      =   1995
   ClientTop       =   1890
   ClientWidth     =   16650
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "qeybat_gotrei.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   16650
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "€Ì» "
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   1200
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000004&
      Caption         =   "›ﬁÿ ¬“„Ê‰ ‘›«ÂÌ"
      Height          =   315
      Left            =   3600
      TabIndex        =   109
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000004&
      Caption         =   "À»  ŒÊœò«— êÊ«ÂÌ ‰«„Â"
      Height          =   315
      Left            =   7440
      TabIndex        =   106
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Caption         =   "«⁄„«·"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   435
      Left            =   11400
      TabIndex        =   88
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ComboBox Combo5 
      Height          =   435
      Left            =   10200
      TabIndex        =   87
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000002&
      Height          =   3525
      ItemData        =   "qeybat_gotrei.frx":08CA
      Left            =   13080
      List            =   "qeybat_gotrei.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   86
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Frame Frame7 
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
      Height          =   4575
      Left            =   12960
      TabIndex        =   78
      Top             =   0
      Width           =   3615
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "noe"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   345
         Left            =   480
         TabIndex        =   108
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄"
         Height          =   315
         Left            =   2280
         TabIndex        =   107
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   315
         Left            =   2280
         TabIndex        =   105
         Top             =   3600
         Width           =   315
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "tarh"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   104
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "‰„—Â ò »Ì"
         Height          =   315
         Left            =   2280
         TabIndex        =   103
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "‰„—Â ‘›«ÂÌ"
         Height          =   315
         Left            =   2280
         TabIndex        =   102
         Top             =   2520
         Width           =   750
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
         Height          =   315
         Left            =   2280
         TabIndex        =   101
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "‰ ÌÃÂ"
         Height          =   315
         Left            =   2280
         TabIndex        =   100
         Top             =   2880
         Width           =   330
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "tozih"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   99
         Top             =   3240
         Width           =   135
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "natije"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   345
         Left            =   480
         TabIndex        =   98
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "shafahi"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   97
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "katbi"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   96
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "mostamar"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   95
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "varaqe"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   94
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "t_emtahan"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   93
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "kodclass"
         DataSource      =   "emtahan_omomi"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   92
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label46 
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
         TabIndex        =   85
         Top             =   2520
         Width           =   60
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "‰„—Â Ê—ﬁÂ"
         Height          =   315
         Left            =   2280
         TabIndex        =   84
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "‰„—Â ò·«”Ì"
         Height          =   315
         Left            =   2280
         TabIndex        =   83
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Parvande"
         DataSource      =   "emtahan_omomi"
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
         TabIndex        =   82
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «„ Õ«‰"
         Height          =   315
         Left            =   2280
         TabIndex        =   81
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "ò·«”"
         Height          =   315
         Left            =   2280
         TabIndex        =   80
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   79
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "Ã«Ìê“Ì‰Ì"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Õ–›"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   1200
      Width           =   735
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
      Height          =   375
      Left            =   2400
      TabIndex        =   73
      Top             =   4680
      Visible         =   0   'False
      Width           =   8175
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
         Connect         =   $"qeybat_gotrei.frx":08CE
         OLEDBString     =   $"qeybat_gotrei.frx":0957
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
         Connect         =   $"qeybat_gotrei.frx":09E0
         OLEDBString     =   $"qeybat_gotrei.frx":0A69
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
         Connect         =   $"qeybat_gotrei.frx":0AF2
         OLEDBString     =   $"qeybat_gotrei.frx":0B7B
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
         Connect         =   $"qeybat_gotrei.frx":0C04
         OLEDBString     =   $"qeybat_gotrei.frx":0C8D
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
         Connect         =   $"qeybat_gotrei.frx":0D16
         OLEDBString     =   $"qeybat_gotrei.frx":0D9F
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
         Connect         =   $"qeybat_gotrei.frx":0E28
         OLEDBString     =   $"qeybat_gotrei.frx":0EB1
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
         Connect         =   $"qeybat_gotrei.frx":0F3A
         OLEDBString     =   $"qeybat_gotrei.frx":0FC3
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
         Connect         =   $"qeybat_gotrei.frx":104C
         OLEDBString     =   $"qeybat_gotrei.frx":10D5
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
         Connect         =   $"qeybat_gotrei.frx":115E
         OLEDBString     =   $"qeybat_gotrei.frx":11E7
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
         Connect         =   $"qeybat_gotrei.frx":1270
         OLEDBString     =   $"qeybat_gotrei.frx":12F9
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
         Connect         =   $"qeybat_gotrei.frx":1382
         OLEDBString     =   $"qeybat_gotrei.frx":140B
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
      Begin MSAdodcLib.Adodc emtahan_omomi 
         Height          =   375
         Left            =   2640
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
         Connect         =   $"qeybat_gotrei.frx":1494
         OLEDBString     =   $"qeybat_gotrei.frx":151D
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from emtahan_omomi"
         Caption         =   "omomi_"
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
      Begin MSAdodcLib.Adodc dore_table 
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
         Connect         =   $"qeybat_gotrei.frx":15A6
         OLEDBString     =   $"qeybat_gotrei.frx":162F
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from dore_table"
         Caption         =   "dore_table"
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
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "«’·«Õ ‰„—Â"
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "À»  ‰„—Â"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "¬“„Ê‰ „Ãœœ"
      Height          =   1335
      Left            =   13200
      TabIndex        =   67
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0FF&
         Height          =   420
         Left            =   120
         TabIndex        =   69
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0FF&
         Height          =   420
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ò »Ì"
         Height          =   315
         Left            =   1200
         TabIndex        =   71
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "‘›«ÂÌ"
         Height          =   315
         Left            =   1200
         TabIndex        =   70
         Top             =   960
         Width           =   435
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "‰„—« "
      Height          =   1335
      Left            =   3480
      TabIndex        =   62
      Top             =   2760
      Width           =   3615
      Begin VB.TextBox text_katbi 
         BackColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox text_shafahi 
         BackColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox text_varaqe 
         Height          =   420
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox text_kelasi 
         BackColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.Line Line5 
         X1              =   600
         X2              =   600
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   1800
         X2              =   600
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         X1              =   1800
         X2              =   1800
         Y1              =   840
         Y2              =   240
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   1800
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "‘›«ÂÌ"
         Height          =   315
         Left            =   1320
         TabIndex        =   66
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "ò »Ì"
         Height          =   315
         Left            =   1320
         TabIndex        =   65
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "ò·«”Ì"
         Height          =   315
         Left            =   3000
         TabIndex        =   64
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ê—ﬁÂ"
         Height          =   315
         Left            =   3120
         TabIndex        =   63
         Top             =   360
         Width           =   300
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   1920
         Y1              =   360
         Y2              =   1200
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.Frame Frame4 
      Caption         =   "„‘Œ’«  «„ Õ«‰"
      Height          =   2295
      Left            =   7200
      TabIndex        =   54
      Top             =   1800
      Width           =   5655
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FF8080&
         Caption         =   "«⁄„«·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1890
         Width           =   2055
      End
      Begin VB.TextBox text_tozih 
         Alignment       =   1  'Right Justify
         Height          =   660
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox Combo4 
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Text            =   " "
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   435
         Left            =   2400
         TabIndex        =   4
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox Combo6 
         Height          =   435
         Left            =   2400
         TabIndex        =   6
         Text            =   "«„Ì— ê«∆Ì‰Ì"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox Combo7 
         BackColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2400
         TabIndex        =   57
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox Combo8 
         BackColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   3240
         TabIndex        =   56
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox Combo9 
         BackColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   3840
         TabIndex        =   55
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox Combo10 
         Height          =   435
         Left            =   2400
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
         Height          =   315
         Left            =   1680
         TabIndex        =   74
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "‰ ÌÃÂ"
         Height          =   315
         Left            =   1800
         TabIndex        =   72
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "„ÕœÊÂ «„ Õ«‰"
         Height          =   300
         Left            =   4320
         TabIndex        =   61
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ „„ Õ‰"
         Height          =   300
         Left            =   4560
         TabIndex        =   60
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄ «„ Õ«‰"
         Height          =   300
         Left            =   4560
         TabIndex        =   59
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «„ Õ«‰"
         Height          =   300
         Left            =   4560
         TabIndex        =   58
         Top             =   1920
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   7200
      TabIndex        =   51
      Top             =   480
      Width           =   2895
      Begin VB.ComboBox Combo11 
         Height          =   435
         Left            =   240
         TabIndex        =   113
         Text            =   "œÊ—Â"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   435
         Left            =   1200
         TabIndex        =   2
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   300
         Left            =   2400
         TabIndex        =   53
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "ò·«”"
         Height          =   300
         Left            =   2160
         TabIndex        =   52
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   4095
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   3255
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å«Ì«‰"
         Height          =   345
         Left            =   2040
         TabIndex        =   49
         Top             =   2880
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
         TabIndex        =   48
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ ‘—Ê⁄"
         Height          =   345
         Left            =   2040
         TabIndex        =   47
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "—Ê“ Â«Ì ò·«”"
         Height          =   345
         Left            =   2040
         TabIndex        =   46
         Top             =   3240
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
         TabIndex        =   45
         Top             =   2400
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
         TabIndex        =   44
         Top             =   3120
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
         Height          =   330
         Index           =   0
         Left            =   2040
         TabIndex        =   43
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   330
         Left            =   2040
         TabIndex        =   42
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "„ﬁÿ⁄"
         Height          =   330
         Left            =   2040
         TabIndex        =   41
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
         Height          =   330
         Left            =   2040
         TabIndex        =   40
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄ "
         Height          =   330
         Left            =   2040
         TabIndex        =   39
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "„œ—”"
         Height          =   330
         Left            =   2040
         TabIndex        =   38
         Top             =   2160
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   720
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
         TabIndex        =   35
         Top             =   1080
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
         TabIndex        =   34
         Top             =   1800
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
         TabIndex        =   33
         Top             =   1440
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
         TabIndex        =   32
         Top             =   2040
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
         TabIndex        =   31
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   " «"
         Height          =   330
         Left            =   720
         TabIndex        =   30
         Top             =   1800
         Width           =   120
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
      Height          =   2775
      Left            =   3480
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   28
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "‰«„"
         Height          =   315
         Left            =   2280
         TabIndex        =   27
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         Height          =   315
         Left            =   2280
         TabIndex        =   26
         Top             =   1080
         Width           =   870
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tozih"
         DataSource      =   "Student"
         Height          =   315
         Left            =   480
         TabIndex        =   22
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Mob"
         DataSource      =   "Student"
         Height          =   315
         Left            =   480
         TabIndex        =   21
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tell"
         DataSource      =   "Student"
         Height          =   315
         Left            =   480
         TabIndex        =   20
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Â„—«Â"
         Height          =   315
         Left            =   2280
         TabIndex        =   19
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   " ·›‰ À«» "
         Height          =   315
         Left            =   2280
         TabIndex        =   18
         Top             =   1440
         Width           =   600
      End
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
         TabIndex        =   17
         Top             =   2520
         Width           =   60
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "qeybat_gotrei.frx":16B8
      Height          =   3735
      Left            =   120
      TabIndex        =   50
      ToolTipText     =   "»—«Ì „‘«ÂœÂ «”ò‰ ›«Ì· »— —ÊÌ Å—Ê‰œÂ œÊ »«— ò·Ìò ò‰Ìœ"
      Top             =   4680
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   6588
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
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      Style           =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ﬁ—¬‰ ¬„Ê“«‰"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "‰„—«  À»  ‘œÂ"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "qeybat_gotrei.frx":16CE
      Height          =   3735
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "»—«Ì „‘«ÂœÂ «”ò‰ ›«Ì· »— —ÊÌ Å—Ê‰œÂ œÊ »«— ò·Ìò ò‰Ìœ"
      Top             =   4680
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12640511
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
      Caption         =   "‰„—«  À»  ‘œÂ œÊ—Â Â«Ì ⁄„Ê„Ì"
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "parvande"
         Caption         =   "parvande"
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
         DataField       =   "kodclass"
         Caption         =   "kodclass"
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
         DataField       =   "t_emtahan"
         Caption         =   "t_emtahan"
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
         DataField       =   "varaqe"
         Caption         =   "varaqe"
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
         DataField       =   "mostamar"
         Caption         =   "mostamar"
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
         DataField       =   "katbi"
         Caption         =   "katbi"
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
         DataField       =   "shafahi"
         Caption         =   "shafahi"
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
         DataField       =   "natije"
         Caption         =   "natije"
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
         DataField       =   "tozih"
         Caption         =   "tozih"
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
         DataField       =   "op"
         Caption         =   "op"
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
         DataField       =   "d"
         Caption         =   "d"
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
         DataField       =   "noe"
         Caption         =   "noe"
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
         DataField       =   "momtahen"
         Caption         =   "momtahen"
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
         DataField       =   "tarh"
         Caption         =   "tarh"
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
         DataField       =   "mahdode"
         Caption         =   "mahdode"
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
         DataField       =   "id"
         Caption         =   "id"
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
      EndProperty
   End
   Begin VB.Label Label60 
      BackColor       =   &H000000FF&
      Caption         =   "Label60"
      Height          =   255
      Left            =   5520
      TabIndex        =   111
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label59 
      Caption         =   "Label59"
      Height          =   375
      Left            =   6600
      TabIndex        =   110
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      Caption         =   "ò·«”"
      Height          =   315
      Left            =   12600
      TabIndex        =   89
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label check_lable 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "check"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   12120
      TabIndex        =   75
      Top             =   1680
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Menu mnmnm 
      Caption         =   "#"
   End
   Begin VB.Menu mid 
      Caption         =   "?"
      Begin VB.Menu mnusabt_nomre 
         Caption         =   "À»  ‰„—Â"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "sabt_nomre_omoomi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Combo1.Text + "%')"
mclass.Refresh
End Sub

Private Sub Command1_Click()

If Combo7.Text = "" Or Combo8.Text = "" Or Combo9.Text = "" Then
MsgBox " «—ÌŒ ¬“„Ê‰ —« »Â ’Ê—  ’ÕÌÕ Ê«—œ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If
If Combo4.Text = "" Then
MsgBox "‰ ÌÃÂ —« „‘Œ’ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If
If Combo11.Text = "" Then
MsgBox "œÊ—Â —« „‘Œ’ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If

'check nomre baraye sabt>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>..
'baraye check nahaee va ersal

If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ"
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"


        ' sabt nomre baraye govahi name
                If Combo2.Text = "¬“„Ê‰ «’·Ì" Or Combo2.Text = "¬“„Ê‰ „Ãœœ" Then

If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ"
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ‘›«ÂÌ"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"

 End If
               
        If Combo2.Text = "¬“„Ê‰ „Ãœœ ò »Ì" Or Combo2.Text = "¬“„Ê‰ ò »Ì" Then
        
If Val(text_katbi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"
If Val(text_katbi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"

 End If
                If Combo2.Text = "¬“„Ê‰ „Ãœœ ‘›«ÂÌ" Or Combo2.Text = "¬“„Ê‰ ‘›«ÂÌ" Then
                
If Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ‘›«ÂÌ"
If Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"
  End If
'end check nomre for then


Dim RaNdOm_id As Single
2: RaNdOm_id = Int(Rnd(100) * 100)
emtahan_omomi.Refresh
emtahan_omomi.RecordSource = "select * from emtahan_omomi where id like ('%" & Label8.Caption & RaNdOm_id & "%') "
emtahan_omomi.Refresh
If emtahan_omomi.Recordset.BOF = True Or emtahan_omomi.Recordset.EOF = True Then
GoTo 1
Else
GoTo 2
End If
MsgBox "Err", vbCritical, "vbP"

Exit Sub
1:

emtahan_omomi.Refresh
emtahan_omomi.RecordSource = "select * from emtahan_omomi where parvande like ('%" & Label8.Caption & "%')and kodclass like ('%" & lkodclass.Caption & "%') and noe like ('%" & Combo2.Text & "%')"
emtahan_omomi.Refresh
If emtahan_omomi.Recordset.EOF = False Or emtahan_omomi.Recordset.BOF = False Then
 MsgBox "shoma qablan een nomre ra sabt karde eid", vbCritical + vbOKOnly, "khata"
 Exit Sub
 End If
 
 
emtahan_omomi.Refresh
emtahan_omomi.Recordset.AddNew
emtahan_omomi.Recordset.Fields("id") = Label8.Caption & RaNdOm_id
emtahan_omomi.Recordset.Fields("parvande") = Label8.Caption
emtahan_omomi.Recordset.Fields("kodclass") = lkodclass.Caption
emtahan_omomi.Recordset.Fields("t_emtahan") = Combo7.Text & "/" & Combo8.Text & "/" & Combo9.Text
emtahan_omomi.Recordset.Fields("varaqe") = text_varaqe.Text
emtahan_omomi.Recordset.Fields("mostamar") = text_kelasi.Text
emtahan_omomi.Recordset.Fields("katbi") = text_katbi.Text
emtahan_omomi.Recordset.Fields("shafahi") = text_shafahi.Text
emtahan_omomi.Recordset.Fields("natije") = Combo4.Text
emtahan_omomi.Recordset.Fields("tozih") = text_tozih.Text
emtahan_omomi.Recordset.Fields("noe") = Combo2.Text
emtahan_omomi.Recordset.Fields("tarh") = Combo3.Text
emtahan_omomi.Recordset.Fields("kod_dore") = Combo11.Text

emtahan_omomi.Recordset.Fields("d") = Entekhab.SB.Panels(3).Text
emtahan_omomi.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
emtahan_omomi.Recordset.Update
emtahan_omomi.Refresh
MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbInformation + vbOKOnly, "À»  ‰„—«  œÊ—Â Â«Ì ⁄„Ê„Ì"

If Check1.Value = 1 And Combo4.Text = "ﬁ»Ê·" Then
Govahi.Show
Govahi.Text1.Text = Me.Label8.Caption
        ' sabt nomre baraye govahi name
                If Combo2.Text = "¬“„Ê‰ «’·Ì" Or Combo2.Text = "¬“„Ê‰ „Ãœœ" Then
                Govahi.Text2.Text = ((Val(text_katbi.Text) + Val(text_shafahi.Text)) / 2)
                'payane check baraye en mored
                End If
        If Combo2.Text = "¬“„Ê‰ „Ãœœ ò »Ì" Or Combo2.Text = "¬“„Ê‰ ò »Ì" Then
        Govahi.Text2.Text = (Val(text_katbi.Text))
        'payane check baraye en mored
        End If
                If Combo2.Text = "¬“„Ê‰ „Ãœœ ‘›«ÂÌ" Or Combo2.Text = "¬“„Ê‰ ‘›«ÂÌ" Then
                Govahi.Text2.Text = (Val(text_shafahi.Text))
                'payane check baraye en mored
                End If
End If

' sabt nomre baraye govagi namne





text_katbi.Text = ""
text_kelasi.Text = ""
text_varaqe.Text = ""
text_shafahi.Text = ""
Combo4.Text = ""



End Sub



Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'baraye check nahaee va ersal
Exit Sub

If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ"
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"


        ' sabt nomre baraye govahi name
                If Combo2.Text = "¬“„Ê‰ «’·Ì" Or Combo2.Text = "¬“„Ê‰ „Ãœœ" Then

If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ"
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"

 End If
               
        If Combo2.Text = "¬“„Ê‰ „Ãœœ ò »Ì" Or Combo2.Text = "¬“„Ê‰ ò »Ì" Then
        
If Val(text_katbi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"
If Val(text_katbi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"

 End If
                If Combo2.Text = "¬“„Ê‰ „Ãœœ ‘›«ÂÌ" Or Combo2.Text = "¬“„Ê‰ ‘›«ÂÌ" Then
                
If Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"
  End If
              

End Sub

Private Sub Command2_Click()
         '   emtahan_omomi.Refresh
          '  emtahan_omomi.RecordSource = "select * from emtahan_omomi where parvande like ('%" & emtahan_omomi.Recordset.Fields("parvande") & "%') "
          '  emtahan_omomi.Refresh
On Error Resume Next

 text_varaqe.Text = emtahan_omomi.Recordset.Fields("varaqe")
text_kelasi.Text = emtahan_omomi.Recordset.Fields("mostamar")
 text_katbi.Text = emtahan_omomi.Recordset.Fields("katbi")
 text_shafahi.Text = emtahan_omomi.Recordset.Fields("shafahi")
Combo4.Text = emtahan_omomi.Recordset.Fields("natije")
Combo3.Text = emtahan_omomi.Recordset.Fields("tarh")
text_tozih.Text = emtahan_omomi.Recordset.Fields("tozih")
Combo2.Text = emtahan_omomi.Recordset.Fields("noe")
Combo11.Text = emtahan_omomi.Recordset.Fields("kod_dore")

Beep

'MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbInformation + vbOKOnly, "À»  ‰„—«  œÊ—Â Â«Ì ⁄„Ê„Ì"

End Sub

Private Sub Command3_Click()
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ‰„—Â —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Â‘œ«—") = vbYes Then
emtahan_omomi.Recordset.Delete

End If

End Sub

Private Sub Command4_Click()

'emtahan_omomi.Refresh
'emtahan_omomi.Recordset.AddNew
'emtahan_omomi.Recordset.Fields("parvande") = Label8.Caption
'emtahan_omomi.Recordset.Fields("kodclass") = lkodclass.Caption
'emtahan_omomi.Recordset.Fields("t_emtahan") = Combo7.Text & "/" & Combo8.Text & "/" & Combo8.Text
emtahan_omomi.Recordset.Fields("varaqe") = text_varaqe.Text
emtahan_omomi.Recordset.Fields("mostamar") = text_kelasi.Text
emtahan_omomi.Recordset.Fields("katbi") = text_katbi.Text
emtahan_omomi.Recordset.Fields("shafahi") = text_shafahi.Text
emtahan_omomi.Recordset.Fields("natije") = Combo4.Text
emtahan_omomi.Recordset.Fields("tozih") = text_tozih.Text
emtahan_omomi.Recordset.Fields("noe") = Combo2.Text
' Combo11.Text = emtahan_omomi.Recordset.Fields("kod_dore")
emtahan_omomi.Recordset.Fields("kod_dore") = Combo11.Text

'emtahan_omomi.Recordset.Fields("d") = Entekhab.SB.Panels(3).Text
'emtahan_omomi.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
emtahan_omomi.Recordset.Update
emtahan_omomi.Refresh
MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ã«Ìê“Ì‰ ‘œ", vbInformation + vbOKOnly, "Ã«Ìê“Ì‰Ì ‰„—«  œÊ—Â Â«Ì ⁄„Ê„Ì"



End Sub

Private Sub Command5_Click()
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + Text4.Text + "%') and elat like ('%" + Combo5.Text + "%') "
STU2CLASS.Refresh
'Label37.Caption = STU2CLASS.Recordset.RecordCount
List1.Clear

For I = 1 To STU2CLASS.Recordset.RecordCount

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh

List1.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
STU2CLASS.Recordset.MoveNext
Next I
End Sub

Private Sub Command7_Click()
'Command7.Default = True



If Combo7.Text = "" Or Combo8.Text = "" Or Combo9.Text = "" Then
MsgBox " «—ÌŒ ¬“„Ê‰ —« »Â ’Ê—  ’ÕÌÕ Ê«—œ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If

If Combo11.Text = "" Then
MsgBox "œÊ—Â —« „‘Œ’ ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If

'check nomre baraye sabt>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>..
'baraye check nahaee va ersal


'end check nomre for then


Dim RaNdOm_id As Single
2: RaNdOm_id = Int(Rnd(100) * 100)
emtahan_omomi.Refresh
emtahan_omomi.RecordSource = "select * from emtahan_omomi where id like ('%" & Label8.Caption & RaNdOm_id & "%') "
emtahan_omomi.Refresh
If emtahan_omomi.Recordset.BOF = True Or emtahan_omomi.Recordset.EOF = True Then
GoTo 1
Else
GoTo 2
End If
MsgBox "Err", vbCritical, "vbP"

Exit Sub
1:

emtahan_omomi.Refresh
emtahan_omomi.RecordSource = "select * from emtahan_omomi where parvande like ('%" & Label8.Caption & "%')and kodclass like ('%" & lkodclass.Caption & "%') and noe like ('%" & Combo2.Text & "%')"
emtahan_omomi.Refresh
If emtahan_omomi.Recordset.EOF = False Or emtahan_omomi.Recordset.BOF = False Then
 MsgBox "shoma qablan een nomre ra sabt karde eid", vbCritical + vbOKOnly, "khata"
 Exit Sub
 End If
 
 
emtahan_omomi.Refresh
emtahan_omomi.Recordset.AddNew
emtahan_omomi.Recordset.Fields("id") = Label8.Caption & RaNdOm_id
emtahan_omomi.Recordset.Fields("parvande") = Label8.Caption
emtahan_omomi.Recordset.Fields("kodclass") = lkodclass.Caption
emtahan_omomi.Recordset.Fields("t_emtahan") = Combo7.Text & "/" & Combo8.Text & "/" & Combo9.Text
emtahan_omomi.Recordset.Fields("varaqe") = "0"
emtahan_omomi.Recordset.Fields("mostamar") = "0"
emtahan_omomi.Recordset.Fields("katbi") = "0"
emtahan_omomi.Recordset.Fields("shafahi") = "0"
emtahan_omomi.Recordset.Fields("natije") = "€Ì»  œ— ¬“„Ê‰"
emtahan_omomi.Recordset.Fields("tozih") = text_tozih.Text
emtahan_omomi.Recordset.Fields("noe") = Combo2.Text
emtahan_omomi.Recordset.Fields("tarh") = Combo3.Text
emtahan_omomi.Recordset.Fields("kod_dore") = Combo11.Text

emtahan_omomi.Recordset.Fields("d") = Entekhab.SB.Panels(3).Text
emtahan_omomi.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
emtahan_omomi.Recordset.Update
emtahan_omomi.Refresh
MsgBox "€Ì»  œ— ¬“„Ê‰ À»  ‘œ", vbInformation + vbOKOnly, "À»  ‰„—«  œÊ—Â Â«Ì ﬁ—«∆ "





text_katbi.Text = ""
text_kelasi.Text = ""
text_varaqe.Text = ""
text_shafahi.Text = ""
Combo4.Text = ""


End Sub

Private Sub Form_Load()

'Me.Combo3.SetFocus
'Combo1.SetFocus

Combo2.AddItem ("¬“„Ê‰ «’·Ì")
Combo2.AddItem ("¬“„Ê‰ ò »Ì")
Combo2.AddItem ("¬“„Ê‰ ‘›«ÂÌ")

Combo2.AddItem ("¬“„Ê‰ „Ãœœ ò »Ì")
Combo2.AddItem ("¬“„Ê‰ „Ãœœ ‘›«ÂÌ")
Combo2.AddItem ("¬“„Ê‰ „Ãœœ")


Combo2.AddItem ("")
Combo5.AddItem ("€Ì» ")
Combo5.AddItem ("« „«„ ò·«”")



Combo4.AddItem ("¬“„Ê‰ „Ãœœ ò»Ì Ê ‘›«ÂÌ")
Combo4.AddItem ("¬“„Ê‰ „Ãœœ ò »Ì")
Combo4.AddItem ("¬“„Ê‰ „Ãœœ ‘›«ÂÌ")
Combo4.AddItem ("ﬁ»Ê·")
Combo4.AddItem ("êÊ«ÂÌ ‰«„Â")
Combo4.AddItem ("„—œÊœ")
Combo4.AddItem ("‘—ò  „Ãœœ œ— œÊ—Â")


For I = 101 To 251
Combo11.AddItem (I)
Next I


For I = 1390 To 1408
Combo7.AddItem (I)
Next I

For I = 1 To 12
If I < 10 Then
Combo8.AddItem ("0" & I)
Else
Combo8.AddItem (I)
End If
Next I

For I = 1 To 31
If I < 10 Then
Combo9.AddItem ("0" & I)
Else
Combo9.AddItem (I)
End If

Next I
'Combo3.Text = "01"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show

End Sub

Private Sub Label39_Click()
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Label39.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label49_Change()
If Me.Label49.Caption <> "ﬁ»Ê·" Then
'&H000000FF&
'QERMEZ
Label49.ForeColor = &HFF&

Else
'&H00FF0000&
'ABI
Label49.ForeColor = &HFF0000
End If
End Sub

Private Sub Label58_Change()
If Me.Label49.Caption <> "¬“„Ê‰ «’·Ì" Then


'&H000000FF&
'QERMEZ
Label49.ForeColor = &HFF&

Else
'&H00FF0000&
'ABI
Label49.ForeColor = &H8000&
End If
End Sub

Private Sub Label59_Click()
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
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "allgovahi.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject("\\yafatemeh2-pc\F\Markaz Quran & Hadis\FORMXLS\allgovahi.xlsx")
End If



'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\allgovahi.xlsx")
emtahan_omomi.Refresh
emtahan_omomi.RecordSource = "select * from emtahan_omomi"
Dim nofr As Integer
nofr = emtahan_omomi.Recordset.RecordCount
For J = 3 To nofr + 2

'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("b" & J).Value = emtahan_omomi.Recordset.Fields("parvande")
oExcel.ActiveSheet.Range("c" & J).Value = emtahan_omomi.Recordset.Fields("kodclass")
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & emtahan_omomi.Recordset.Fields("parvande") & "%')"
Student.Refresh
oExcel.ActiveSheet.Range("n" & J).Value = Student.Recordset.Fields("tavalod")
oExcel.ActiveSheet.Range("o" & J).Value = Student.Recordset.Fields("mob") & " - " & Student.Recordset.Fields("tell")
 



oExcel.ActiveSheet.Range("d" & J).Value = emtahan_omomi.Recordset.Fields("katbi")
oExcel.ActiveSheet.Range("e" & J).Value = emtahan_omomi.Recordset.Fields("shafahi")
oExcel.ActiveSheet.Range("f" & J).Value = emtahan_omomi.Recordset.Fields("noe")
oExcel.ActiveSheet.Range("g" & J).Value = emtahan_omomi.Recordset.Fields("natije")
oExcel.ActiveSheet.Range("h" & J).Value = emtahan_omomi.Recordset.Fields("tozih")
oExcel.ActiveSheet.Range("i" & J).Value = emtahan_omomi.Recordset.Fields("t_emtahan")
oExcel.ActiveSheet.Range("j" & J).Value = emtahan_omomi.Recordset.Fields("tarh")


oExcel.ActiveSheet.Range("k" & J).Value = Student.Recordset.Fields("name")
oExcel.ActiveSheet.Range("l" & J).Value = Student.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("m" & J).Value = Student.Recordset.Fields("namepedar")





 emtahan_omomi.Recordset.MoveNext
 Next J
 


MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", vbInformation + vbOKOnly, "«ÿ·«⁄«  êÊ«ÂÌ ‰«„Â"
X = "All" & "Govahi"

'oExcel.SaveAs X
'oExcel.Application.Visible = True
'oExcel.Parent.Windows(1).Visible = True
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
oExcel.Application.Visible = True
On Error GoTo 722


oExcel.Parent.Windows(2).Visible = True
GoTo 910
722:

oExcel.Parent.Windows(1).Visible = True
910:
''''''

oExcel.SaveAs X
'oExcel.Close
End Sub

Private Sub Label60_Click()
Exit Sub

Dim RaNdOm_id As Single
emtahan_omomi.Refresh
emtahan_omomi.RecordSource = "select * from emtahan_omomi " 'where parvande like ('%" & Label8.Caption & RaNdOm_id & "%') "
emtahan_omomi.Refresh
For I = 1 To emtahan_omomi.Recordset.RecordCount

2: RaNdOm_id = Int(Rnd(100) * 100)
emtahan_omomi.Recordset.Fields("id") = emtahan_omomi.Recordset.Fields("parvande") & RaNdOm_id

emtahan_omomi.Recordset.Update
'emtahan_omomi.Refresh
emtahan_omomi.Recordset.MoveNext
Next I

'emtahan_omomi.Refresh
'emtahan_omomi.RecordSource = "select * from emtahan_omomi where parvande like ('%" & Label8.Caption & RaNdOm_id & "%') "
'emtahan_omomi.Refresh

End Sub

Private Sub Label8_Change()
Exit Sub
On Error Resume Next

Combo1.Clear


Combo1.AddItem (Me.Student.Recordset.Fields("clas1"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas2"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas3"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas4"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas5"))

Combo1.Text = Combo1.List(0)
End Sub

Private Sub List1_Click()
'On Error Resume Next
Dim A As String
'A = mid(List1.Text, 1, 7)


Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + A + "%')"
Student.Refresh



End Sub

Private Sub mnmnm_Click()
Entekhab.Show

End Sub

Private Sub mnusabt_nomre_Click()
Call Command7_Click

End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Caption

            Case "ﬁ—¬‰ ¬„Ê“«‰"
            DataGrid2.Visible = True
            DataGrid1.Visible = False
      
            
            
            
            Case "‰„—«  À»  ‘œÂ"
            On Error Resume Next
            DataGrid2.Visible = False
            DataGrid1.Visible = True
         
            emtahan_omomi.Refresh
            emtahan_omomi.RecordSource = "select * from emtahan_omomi where parvande like ('%" & Student.Recordset.Fields("parvande") & "%') "
            emtahan_omomi.Refresh
          
                
End Select
End Sub

Private Sub text_katbi_Change()
'If Val(text_katbi.Text) > 20 Then text_katbi.Text = "20"

If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ"
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"



End Sub

Private Sub text_kelasi_Change()
If Val(text_kelasi.Text) > 3 Then text_kelasi.Text = "3"
text_katbi = Val(text_varaqe) + Val(text_kelasi)
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ"
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"



End Sub

Private Sub text_shafahi_Change()
If Val(text_shafahi.Text) > 20 Then text_shafahi.Text = "20"

If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ"
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"


If Check2.Value = 1 Then
If Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"
End If

End Sub

Private Sub text_varaqe_Change()
If Val(text_varaqe.Text) > 17 Then text_varaqe.Text = "17"
text_katbi = Val(text_varaqe) + Val(text_kelasi)
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ"
If Val(text_katbi.Text) < 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "¬“„Ê‰ „Ãœœ ò »Ì"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) < 16 Then Combo4.Text = "¬“„Ê‰ ‘›«ÂÌ"
If Val(text_katbi.Text) >= 16 And Val(text_shafahi.Text) >= 16 Then Combo4.Text = "ﬁ»Ê·"



End Sub

Private Sub Text2_Change()
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + Text2.Text + "%')or parvande like ('%" + Text2.Text + "%') or clas2 like ('%" + Text2.Text + "%')or nf like ('%" + Text2.Text + "%') or  clas3 like ('%" + Text2.Text + "%') or clas4 like ('%" + Text2.Text + "%') or clas5 like ('%" + Text2.Text + "%')"
Student.Refresh
End Sub

Private Sub Text4_Change()
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Text4.Text + "%')"
mclass.Refresh
End Sub
