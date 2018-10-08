VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Govahi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’œÊ— êÊ«ÂÌ ‰«„Â"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Govahi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3480
      TabIndex        =   73
      Text            =   "«‰ Œ«» ò‰Ìœ"
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      DisabledPicture =   "Govahi.frx":08CA
      DownPicture     =   "Govahi.frx":25544
      DragIcon        =   "Govahi.frx":4A1BE
      Height          =   330
      Left            =   6480
      Picture         =   "Govahi.frx":6EE38
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "À»  êÊ«ÂÌ ‰«„Â"
      Height          =   420
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
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
      Height          =   350
      Left            =   8640
      TabIndex        =   28
      Top             =   3000
      Width           =   1560
   End
   Begin VB.Frame Frame4 
      Caption         =   "«ÿ·«⁄«  À» "
      Height          =   2295
      Left            =   3480
      TabIndex        =   23
      Top             =   1080
      Width           =   2895
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   2040
         TabIndex        =   49
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
         Begin VB.OptionButton Option7 
            Alignment       =   1  'Right Justify
            Caption         =   "òœ „·Ì"
            Height          =   330
            Left            =   600
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option6 
            Alignment       =   1  'Right Justify
            Caption         =   "‘.‘"
            Height          =   330
            Left            =   1680
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.OptionButton Option9 
         Alignment       =   1  'Right Justify
         Caption         =   "ç«Å ‘œÂ"
         Height          =   330
         Left            =   1560
         TabIndex        =   13
         Top             =   1260
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Alignment       =   1  'Right Justify
         Caption         =   "ç«Å ‰‘œÂ"
         Height          =   360
         Left            =   240
         TabIndex        =   14
         Top             =   1260
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Height          =   420
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text17 
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
         Height          =   420
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ À» "
         Height          =   330
         Left            =   1920
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "À»  ò‰‰œÂ"
         Height          =   330
         Left            =   1920
         TabIndex        =   26
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Height          =   330
         Left            =   4320
         TabIndex        =   25
         Top             =   1920
         Width           =   45
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "„‘Œ’«  êÊ«ÂÌ"
      Height          =   1815
      Left            =   6480
      TabIndex        =   20
      Top             =   1080
      Width           =   3015
      Begin VB.ComboBox Combo2 
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
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
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
         Left            =   1200
         TabIndex        =   8
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   1200
         TabIndex        =   9
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "”ÿÕ"
         Height          =   330
         Left            =   480
         TabIndex        =   33
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "„⁄œ·"
         Height          =   330
         Left            =   2520
         TabIndex        =   30
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   300
         Left            =   2640
         TabIndex        =   22
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "„Õ›ÊŸ« "
         Height          =   330
         Left            =   2280
         TabIndex        =   21
         Top             =   960
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "‰Ê⁄ êÊ«ÂÌ ‰«„Â"
      Height          =   1815
      Left            =   9600
      TabIndex        =   19
      Top             =   1080
      Width           =   1935
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Caption         =   " —Ã„Â Ê „›«ÂÌ„"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   " ÃÊÌœ"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "—Ê ŒÊ«‰Ì Ê —Ê«‰ ŒÊ«‰Ì"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ›Ÿ"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   615
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
      Left            =   6480
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   3255
      Begin MSAdodcLib.Adodc Govahi 
         Height          =   330
         Left            =   360
         Top             =   2880
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
         Connect         =   $"Govahi.frx":93AB2
         OLEDBString     =   $"Govahi.frx":93B3B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from govahi"
         Caption         =   "Govahi"
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
      Begin MSAdodcLib.Adodc Qeybat 
         Height          =   330
         Left            =   360
         Top             =   2520
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
         Connect         =   $"Govahi.frx":93BC4
         OLEDBString     =   $"Govahi.frx":93C4D
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
         Connect         =   $"Govahi.frx":93CD6
         OLEDBString     =   $"Govahi.frx":93D5F
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
         Connect         =   $"Govahi.frx":93DE8
         OLEDBString     =   $"Govahi.frx":93E71
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
         Connect         =   $"Govahi.frx":93EFA
         OLEDBString     =   $"Govahi.frx":93F83
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
         Connect         =   $"Govahi.frx":9400C
         OLEDBString     =   $"Govahi.frx":94095
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
         Connect         =   $"Govahi.frx":9411E
         OLEDBString     =   $"Govahi.frx":941A7
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
         Connect         =   $"Govahi.frx":94230
         OLEDBString     =   $"Govahi.frx":942B9
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
         Connect         =   $"Govahi.frx":94342
         OLEDBString     =   $"Govahi.frx":943CB
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
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   69
      Top             =   8235
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
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
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   ""
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
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   615
      Left            =   3480
      TabIndex        =   70
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1085
      Style           =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ﬁ—¬‰ ¬„Ê“«‰"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "êÊ«ÂÌ ‰«„Â"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "êÊ«ÂÌ ‰«„Â Â«Ì À»  ‘œÂ »—«Ì ﬁ—¬‰ ¬„Ê“"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   7320
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton OptionHEFZ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "êÊ«ÂÌ ‰«ÂÂ Â«"
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ﬁ—¬‰ ¬„Ê“«‰"
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "„‘Œ’«  êÊ«ÂÌ ‰«„Â"
      Height          =   3375
      Left            =   120
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "MOADEL"
         DataSource      =   "Govahi"
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
         Left            =   480
         TabIndex        =   72
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "„⁄œ·"
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
         TabIndex        =   71
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄ êÊ«ÂÌ"
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
         TabIndex        =   66
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "noe"
         DataSource      =   "Govahi"
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
         TabIndex        =   65
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   " ÕÊÌ·"
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
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "TAHVIL"
         DataSource      =   "Govahi"
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
         Left            =   480
         TabIndex        =   61
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "sader"
         DataSource      =   "Govahi"
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
         Left            =   480
         TabIndex        =   60
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷⁄Ì  ’œÊ—"
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
         TabIndex        =   59
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â"
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
         TabIndex        =   58
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Ã“¡"
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
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "kodG"
         DataSource      =   "Govahi"
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
         TabIndex        =   56
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "JOZE"
         DataSource      =   "Govahi"
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
         TabIndex        =   55
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "”ÿÕ"
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
         TabIndex        =   54
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "SATH"
         DataSource      =   "Govahi"
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
         Left            =   480
         TabIndex        =   53
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "CHAP"
         DataSource      =   "Govahi"
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
         Left            =   480
         TabIndex        =   52
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "ç«Å"
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
         TabIndex        =   51
         Top             =   2040
         Width           =   300
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "„‘Œ’«  ﬁ—¬‰ ¬„Ê“"
      Height          =   3495
      Left            =   120
      TabIndex        =   34
      Top             =   -120
      Width           =   3255
      Begin VB.Label Label41 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   64
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label40 
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
         TabIndex        =   63
         Top             =   2880
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "‘.‘"
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
         TabIndex        =   48
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Shsh"
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   47
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Kodmeli"
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   46
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "òœ „·Ì"
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
         Width           =   465
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Namepedar"
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
         TabIndex        =   44
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tavalod"
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
         TabIndex        =   43
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label13 
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   480
         TabIndex        =   42
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   480
         TabIndex        =   41
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label9 
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
         TabIndex        =   40
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ  Ê·œ"
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
         TabIndex        =   39
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ Åœ—"
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
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label20 
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
         TabIndex        =   37
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label21 
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   360
         Width           =   840
      End
   End
   Begin MSDataGridLib.DataGrid DataGridSTUDENT 
      Bindings        =   "Govahi.frx":94454
      Height          =   4695
      Left            =   120
      TabIndex        =   68
      Top             =   3480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Govahi.frx":9446A
      Height          =   4695
      Left            =   120
      TabIndex        =   31
      ToolTipText     =   "»—«Ì Ã«Ìê“Ì‰Ì Å—Ê‰œÂ œÊ »« — »— —ÊÌ ‰«„ «Ê ò·Ìò ò‰Ìœ"
      Top             =   3480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
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
      Caption         =   "êÊ«ÂÌ ‰«„Â"
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "kodG"
         Caption         =   "òœ êÊ«ÂÌ ‰«„Â"
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
      BeginProperty Column02 
         DataField       =   "NAME"
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
      BeginProperty Column03 
         DataField       =   "FAMIL"
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
      BeginProperty Column04 
         DataField       =   "NAMEPEDAR"
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
      BeginProperty Column05 
         DataField       =   "TTAVALOD"
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
      BeginProperty Column06 
         DataField       =   "SHSH"
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
      BeginProperty Column07 
         DataField       =   "SADERE"
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
      BeginProperty Column08 
         DataField       =   "MOADEL"
         Caption         =   "„⁄œ· òÊ«ÂÌ ‰«„Â"
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
         DataField       =   "JOZE"
         Caption         =   "Ã“¡"
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
         DataField       =   "SATH"
         Caption         =   "”ÿÕ"
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
         DataField       =   "TSABT"
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
      BeginProperty Column12 
         DataField       =   "CHAP"
         Caption         =   "Ê÷⁄Ì  ç«Å"
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
         DataField       =   "TAHVIL"
         Caption         =   "Ê÷⁄Ì   ÕÊÌ·"
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
         DataField       =   "TTAHVIL"
         Caption         =   " «—ÌŒ  ÕÊÌ·"
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
         DataField       =   "TOZIH"
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
         DataField       =   "op"
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
      BeginProperty Column17 
         DataField       =   "NOE"
         Caption         =   "‰Ê⁄ êÊ«ÂÌ ‰«„Â"
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
         DataField       =   "sader"
         Caption         =   "Ê÷⁄Ì  ’œÊ—"
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
         DataField       =   "kodmeli"
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
      EndProperty
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   330
      Left            =   5760
      TabIndex        =   74
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "‘„«—Â êÊ«ÂÌ ‰«„Â"
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
      Left            =   10320
      TabIndex        =   29
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ã” ÃÊ"
      Height          =   330
      Left            =   10920
      TabIndex        =   18
      Top             =   120
      Width           =   420
   End
   Begin VB.Menu mnuhome 
      Caption         =   "#"
   End
   Begin VB.Menu MNUGOVAHI 
      Caption         =   "êÊ«ÂÌ ‰«„Â"
      Begin VB.Menu mnusodoor 
         Caption         =   "À»  ’œÊ— êÊ«ÂÌ ‰«„Â"
      End
      Begin VB.Menu d 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MNUTAHVIL 
         Caption         =   " ÕÊÌ· êÊ«ÂÌ ‰«„Â »Â ﬁ—¬‰ ¬„Ê“"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuschap 
         Caption         =   "À»  ç«Å êÊ«ÂÌ ‰«„Â"
         Shortcut        =   ^P
      End
      Begin VB.Menu MNQE 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuedit 
         Caption         =   " €ÌÌ— „‘Œ’« "
      End
      Begin VB.Menu MUNDELEI 
         Caption         =   "Õ–› êÊ«ÂÌ ‰«„Â"
      End
   End
End
Attribute VB_Name = "Govahi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo4_Click()



If Combo4.Text = "’«œ— ‘œÂ" Then
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where sader like ('%" + "’«œ— ‘œÂ" + "%') "
Govahi.Refresh
End If

If Combo4.Text = "’«œ— ‰‘œÂ" Then
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where sader like ('%" + "’«œ— ‰‘œÂ" + "%') "
Govahi.Refresh
End If

If Combo4.Text = "ç«Å ‘œÂ" Then
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where chap like ('%" + "ç«Å ‘œÂ" + "%') "
Govahi.Refresh
End If

If Combo4.Text = "ç«Å ‰‘œÂ" Then
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where chap like ('%" + "ç«Å ‰‘œÂ" + "%') "
Govahi.Refresh
End If

If Combo4.Text = "’«œ— ‘œÂ Ê ç«Å ‘œÂ" Then
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where sader like ('%" + "’«œ— ‘œÂ" + "%') and chap like ('%" + "ç«Å ‘œÂ" + "%') "
Govahi.Refresh
End If

If Combo4.Text = "’«œ— ‘œÂ Ê ç«Å ‰‘œÂ" Then
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where sader like ('%" + "’«œ— ‘œÂ" + "%') and chap like ('%" + "ç«Å ‰‘œÂ" + "%') "
Govahi.Refresh
End If

If Combo4.Text = "ç«Å ‘œÂ  ÕÊÌ· œ«œÂ ‰‘œÂ" Then
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where tahvil like ('%" + " ÕÊÌ· œ«œÂ ‰‘œÂ" + "%') and chap like ('%" + "ç«Å ‘œÂ" + "%') "
Govahi.Refresh
End If

If Combo4.Text = "ç«Å ‘œÂ  ÕÊÌ· œ«œÂ ‘œÂ" Then
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where tahvil like ('%" + " ÕÊÌ· œ«œÂ ‘œÂ" + "%') and chap like ('%" + "ç«Å ‘œÂ" + "%') "
Govahi.Refresh
End If


End Sub

Private Sub Command1_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "govahi-newsabt" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513




If Combo1.Text = "«‰ Œ«» ò‰Ìœ" Or Combo2.Text = "«‰ Œ«» ò‰Ìœ" Or Combo3.Text = "«‰ Œ«» ò‰Ìœ" Or Text2.Text = "" Or Val(Text2.Text) > 20 Or Val(Text2.Text) < 10 Then
MsgBox "„‘Œ’«  êÊ«ÂÌ ‰«„Â —« »——”Ì ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If


If MsgBox("À»  êÊ«ÂÌ ‰«„Â «‰Ã«„ ŒÊ«Âœ ‘œ ¬Ì« „ÿ„∆‰ Â” Ìœ", vbQuestion + vbYesNo, "À»  êÊ«ÂÌ ‰«„Â") = vbYes Then
GoTo 1

Else
Exit Sub
End If

1:

'êÊ«ÂÌ ‰«„Â Õ›Ÿ ﬁ—¬‰ ò—Ì„
If Option1.Value = True Then
Dim Dore As String


If Combo3.Text = "œÊ ”«·Â" Then Dore = "02"
If Combo3.Text = "çÂ«— ”«·Â" Then Dore = "04"
If Combo3.Text = "‘‘ ”«·Â" Then Dore = "06"
If Combo3.Text = " À»Ì " Then Dore = "08"
If Combo3.Text = "¬“«œ" Then Dore = "09"

TEP.Text = Student.Recordset.Fields("parvande") & "-" & Dore & Combo1.Text

End If
If Option2.Value = True Then
TEP.Text = Student.Recordset.Fields("parvande") & "-" & "20"
End If
If Option3.Value = True Then
If Combo2.Text = "”ÿÕ 1" Then TEP.Text = Student.Recordset.Fields("parvande") & "-" & "2301"
If Combo2.Text = "”ÿÕ 2" Then TEP.Text = Student.Recordset.Fields("parvande") & "-" & "2302"

End If
If Option4.Value = True Then
If Combo2.Text = "”ÿÕ 1" Then TEP.Text = Student.Recordset.Fields("parvande") & "-" & "2601"
If Combo2.Text = "”ÿÕ 2" Then TEP.Text = Student.Recordset.Fields("parvande") & "-" & "2602"

End If


Govahi.Refresh
Govahi.RecordSource = "select * from govahi where kodg like ('%" + TEP.Text + "%')"
Govahi.Refresh
If Govahi.Recordset.BOF = True Or Govahi.Recordset.EOF = True Then
GoTo 2
Else
If MsgBox("‘„« ﬁ»·« «Ì‰ êÊ«ÂÌ —« À»  ò—œÂ «Ìœ" & Chr(10) & "¬Ì« „Ì ŒÊ«ÂÌœ «œ«„Â œÂÌœ", vbCritical + vbYesNo, "Œÿ«") = vbYes Then
GoTo 2
Else

TEP.Text = ""
Exit Sub
End If
End If

2:

Govahi.Refresh
Govahi.Recordset.AddNew
Govahi.Recordset.Fields("kodg") = TEP.Text
Govahi.Recordset.Fields("parvande") = Student.Recordset.Fields("parvande")
Govahi.Recordset.Fields("name") = Student.Recordset.Fields("name")
Govahi.Recordset.Fields("famil") = Student.Recordset.Fields("famil")
Govahi.Recordset.Fields("namepedar") = Student.Recordset.Fields("namepedar")
Govahi.Recordset.Fields("ttavalod") = Student.Recordset.Fields("tavalod")

'If Option6.Value = True Then Govahi.Recordset.Fields("shsh") = Student.Recordset.Fields("shsh")
'If Option7.Value = True Then Govahi.Recordset.Fields("shsh") = Student.Recordset.Fields("kodmeli")
Govahi.Recordset.Fields("shsh") = Student.Recordset.Fields("shsh")

Govahi.Recordset.Fields("kodmeli") = Student.Recordset.Fields("kodmeli")



Govahi.Recordset.Fields("sadere") = Student.Recordset.Fields("sadere")
Govahi.Recordset.Fields("moadel") = Text2.Text
Govahi.Recordset.Fields("joze") = Combo1.Text
Govahi.Recordset.Fields("tsabt") = Text4.Text
If Option8.Value = True Then Govahi.Recordset.Fields("chap") = "ç«Å ‰‘œÂ"
If Option9.Value = True Then Govahi.Recordset.Fields("chap") = "ç«Å ‘œÂ"

Govahi.Recordset.Fields("tahvil") = " ÕÊÌ· œ«œÂ ‰‘œÂ"
Govahi.Recordset.Fields("ttahvil") = ""
Govahi.Recordset.Fields("tozih") = ""
Govahi.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text

Govahi.Recordset.Fields("sader") = "’«œ— ‰‘œÂ"





If Option1.Value = True Then Govahi.Recordset.Fields("noe") = Option1.Caption
If Option2.Value = True Then Govahi.Recordset.Fields("noe") = Option2.Caption
If Option3.Value = True Then Govahi.Recordset.Fields("noe") = Option3.Caption
If Option4.Value = True Then Govahi.Recordset.Fields("noe") = Option4.Caption






Govahi.Recordset.Update
Govahi.Refresh



MsgBox "À»  êÊ«ÂÌ ‰«„Â »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbInformation + vbOKOnly, "À»  êÊ«ÂÌ ‰«„Â"



















End Sub

Private Sub Command11_Click()
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

Dim nofr As Integer
nofr = Govahi.Recordset.RecordCount
For J = 3 To nofr + 2

'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("b" & J).Value = Govahi.Recordset.Fields("kodg")
oExcel.ActiveSheet.Range("c" & J).Value = Govahi.Recordset.Fields("parvande")
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & Govahi.Recordset.Fields("parvande") & "%')"
Student.Refresh
oExcel.ActiveSheet.Range("n" & J).Value = Student.Recordset.Fields("tell")
oExcel.ActiveSheet.Range("o" & J).Value = Student.Recordset.Fields("mob")




oExcel.ActiveSheet.Range("d" & J).Value = Govahi.Recordset.Fields("name")
oExcel.ActiveSheet.Range("e" & J).Value = Govahi.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("f" & J).Value = Govahi.Recordset.Fields("noe")
oExcel.ActiveSheet.Range("g" & J).Value = Govahi.Recordset.Fields("joze")
oExcel.ActiveSheet.Range("h" & J).Value = Govahi.Recordset.Fields("moadel")
oExcel.ActiveSheet.Range("i" & J).Value = Govahi.Recordset.Fields("sath")
oExcel.ActiveSheet.Range("j" & J).Value = Govahi.Recordset.Fields("tsabt")



oExcel.ActiveSheet.Range("k" & J).Value = Govahi.Recordset.Fields("sader")
oExcel.ActiveSheet.Range("l" & J).Value = Govahi.Recordset.Fields("chap")
oExcel.ActiveSheet.Range("m" & J).Value = Govahi.Recordset.Fields("tahvil")
oExcel.ActiveSheet.Range("p" & J).Value = Govahi.Recordset.Fields("tozih")


 Govahi.Recordset.MoveNext
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
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''
'Shell "Explorer.exe " & "c:\" & KodEnhesariPrint & ".xlsx"

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
Govahi.Recordset.Update

End Sub

Private Sub DataGridSTUDENT_DblClick()
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


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Private Sub Form_Load()
Text4.Text = Taqvim.Tarikh.Caption
Combo3.AddItem ("œÊ ”«·Â")
Combo3.AddItem ("çÂ«— ”«·Â")
Combo3.AddItem ("‘‘ ”«·Â")
Combo3.AddItem (" À»Ì ")
Combo3.AddItem ("¬“«œ")
For I = 5 To 30 Step 5
If I < 10 Then
Combo1.AddItem ("0" & I)
Else
Combo1.AddItem (I)
End If
Next I
Combo2.AddItem ("”ÿÕ 1")
Combo2.AddItem ("”ÿÕ 2")




Me.stb1.Panels(1).Text = user.OP.Text
Me.stb1.Panels(3).Text = Taqvim.Label1.Caption


Combo4.AddItem ("’«œ— ‘œÂ")

Combo4.AddItem ("’«œ— ‰‘œÂ")

Combo4.AddItem ("ç«Å ‘œÂ")

Combo4.AddItem ("ç«Å ‰‘œÂ")

Combo4.AddItem ("’«œ— ‘œÂ Ê ç«Å ‘œÂ")

Combo4.AddItem ("’«œ— ‘œÂ Ê ç«Å ‰‘œÂ")


Combo4.AddItem ("ç«Å ‘œÂ  ÕÊÌ· œ«œÂ ‰‘œÂ")

Combo4.AddItem ("ç«Å ‘œÂ  ÕÊÌ· ‘ œ«œÂ ‘œÂ")






End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show

End Sub

Private Sub mnuedit_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "govahi-edit" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513

If DataGrid1.AllowUpdate = False Then
DataGrid1.AllowUpdate = True

mnuedit.Checked = True

Else
DataGrid1.AllowUpdate = False
mnuedit.Checked = False


End If

End Sub

Private Sub mnuhome_Click()
Entekhab.Show

End Sub

Private Sub mnuschap_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "govahi-chap" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub

14082513
If MsgBox("À»  ç«Å êÊ«ÂÌ ‰«„Â «‰Ã«„ ŒÊ«Âœ ‘œ ¬Ì« „ÿ„∆‰ Â” Ìœ", vbQuestion + vbYesNo, "À»  ç«Å") = vbYes Then

Govahi.RecordSource = "select * from govahi where kodg like ('%" + Label31.Caption + "%') "
Govahi.Refresh
Govahi.Recordset.Fields("chap") = "ç«Å ‘œÂ"
'Govahi.Recordset.Fields("tahvil") = " ÕÊÌ· œ«œÂ ‘œ"
'Govahi.Recordset.Fields("ttahvil") = Text4.Text

Govahi.Recordset.Update
Govahi.Refresh

End If

End Sub

Private Sub mnusodoor_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "govahi-sabt-sodor" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513

If MsgBox("À»  ’œÊ— êÊ«ÂÌ ‰«„Â «‰Ã«„ ŒÊ«Âœ ‘œ ¬Ì« „ÿ„∆‰ Â” Ìœ", vbQuestion + vbYesNo, "À»  ’œÊ—") = vbYes Then

Govahi.RecordSource = "select * from govahi where kodg like ('%" + Label31.Caption + "%') "
Govahi.Refresh
Govahi.Recordset.Fields("sader") = "’«œ— ‘œÂ"
'Govahi.Recordset.Fields("tahvil") = " ÕÊÌ· œ«œÂ ‘œ"
'Govahi.Recordset.Fields("ttahvil") = Text4.Text

Govahi.Recordset.Update
Govahi.Refresh
Beep

End If
End Sub

Private Sub MNUTAHVIL_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "govahi-tahvil" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub

14082513
'If Govahi.Recordset.Fields("tahvil") = " ÕÊÌ· œ«œÂ ‘œ" Then
'MsgBox "«Ì‰ êÊ«ÂÌ ﬁ»·« »Â ﬁ—¬‰ ¬„Ê“  ÕÊÌ· œ«œÂ ‘œÂ «” ", vbCritical + vbOKOnly, "Œÿ«"
'Exit Sub
'End If

If MsgBox("À»   ÕÊÌ· êÊ«ÂÌ ‰«„Â »Â ﬁ—¬‰ ¬„Ê“ ... ¬Ì« „ÿ„∆‰ Â” Ìœ", vbQuestion + vbYesNo, " ÕÊÌ· êÊ«ÂÌ ‰«„Â") = vbYes Then

Govahi.RecordSource = "select * from govahi where kodg like ('%" + Label31.Caption + "%') "
Govahi.Refresh
Govahi.Recordset.Fields("chap") = "ç«Å ‘œÂ"
Govahi.Recordset.Fields("tahvil") = " ÕÊÌ· œ«œÂ ‘œ"
Govahi.Recordset.Fields("ttahvil") = Text4.Text

Govahi.Recordset.Update
Govahi.Refresh


Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String
On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:


Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "rgovahi.xlsx")
'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("c5").Value = Govahi.Recordset.Fields("kodg")
oExcel.ActiveSheet.Range("c6").Value = Govahi.Recordset.Fields("noe")
oExcel.ActiveSheet.Range("c7").Value = Govahi.Recordset.Fields("name") & " " & Govahi.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("c8").Value = Govahi.Recordset.Fields("shsh")
oExcel.ActiveSheet.Range("c9").Value = Govahi.Recordset.Fields("ttahvil")



MsgBox "—”Ìœ œ—Ì«›  êÊ«ÂÌ ‰«„Â ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "œ—Ì«›  êÊ«ÂÌ ‰«„Â"
X = Govahi.Recordset.Fields("kodg")

oExcel.SaveAs X
oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True






End If







End Sub

Private Sub MUNDELEI_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "govahi-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513


If Govahi.Recordset.RecordCount = 0 Then
o = MsgBox("‘„« ÂÌç ê“Ì‰Â «Ì  »—«Ì Õ–› ‰œ«—Ìœ ", vbCritical, "Œÿ«")
Else
o = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ êÊ«ÂÌ ‰«„Â —« «“ ”Ì” „ ŒÊœ Õ–› ò‰Ìœ ", vbYesNo + vbQuestion, "Õ–› êÊ«ÂÌ ‰«„Â")

If o = vbYes Then
Govahi.Recordset.Delete
End If
End If
End Sub

Private Sub Option1_Click()
Combo3.Enabled = True
Combo3.Text = "«‰ Œ«» ò‰Ìœ"
Combo2.Enabled = False
Combo2.Text = ""

Combo1.Enabled = True
Combo1.Text = "«‰ Œ«» ò‰Ìœ"

End Sub

Private Sub Option2_Click()
Combo3.Enabled = False
Combo3.Text = ""
Combo2.Enabled = False
Combo2.Text = ""
Combo1.Enabled = False
Combo1.Text = ""
End Sub

Private Sub Option3_Click()
Combo3.Enabled = False
Combo3.Text = ""
Combo2.Enabled = True
Combo2.Text = "«‰ Œ«» ò‰Ìœ"

Combo1.Enabled = False
Combo1.Text = ""
End Sub

Private Sub Option4_Click()
Combo3.Enabled = False
Combo3.Text = ""
Combo2.Enabled = True
Combo2.Text = "«‰ Œ«» ò‰Ìœ"

Combo1.Enabled = False
Combo1.Text = ""
End Sub

Private Sub Option5_Click()
DataGrid1.Visible = False
DataGridSTUDENT.Visible = True
Frame5.Visible = True
Frame8.Visible = False

End Sub

Private Sub OptionHEFZ_Click()
DataGrid1.Visible = True
DataGridSTUDENT.Visible = False
Frame5.Visible = False
Frame8.Visible = True
Govahi.Refresh
Govahi.RecordSource = "select * from govahi where kodg like ('%" + "" + "%') "
Govahi.Refresh
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Caption

            Case "ﬁ—¬‰ ¬„Ê“«‰"
            'DataGrid2.Visible = True
            'DataGrid1.Visible = False
            'Command1.Enabled = False
            Option5.Value = True
            
            
            Case "êÊ«ÂÌ ‰«„Â"
            'On Error Resume Next
            'DataGrid2.Visible = False
            'DataGrid1.Visible = True
            'Command1.Enabled = True
            OptionHEFZ.Value = True
                      Govahi.Refresh
Govahi.RecordSource = "select * from govahi where parvande like ('%" & "" & "%') "

Govahi.Refresh
                      Case "êÊ«ÂÌ ‰«„Â Â«Ì À»  ‘œÂ »—«Ì ﬁ—¬‰ ¬„Ê“"
            'On Error Resume Next
            'DataGrid2.Visible = False
            'DataGrid1.Visible = True
            'Command1.Enabled = True
            OptionHEFZ.Value = True
          Govahi.Refresh
Govahi.RecordSource = "select * from govahi where parvande like ('%" & Student.Recordset.Fields("parvande") & "%') "

Govahi.Refresh
                
End Select
End Sub

Private Sub Text1_Change()
On Error Resume Next

Student.Refresh
Student.RecordSource = "select * from student where famil like ('%" + Text1.Text + "%') or name like ('%" + Text1.Text + "%') or parvande like ('%" + Text1.Text + "%')or nf like ('%" + Text1.Text + "%')"
Student.Refresh

Govahi.Refresh
Govahi.RecordSource = "select * from govahi where kodg like ('%" + Text1.Text + "%') or famil like ('%" + Text1.Text + "%')"
Govahi.Refresh
Label29.Caption = Govahi.Recordset.RecordCount

End Sub


