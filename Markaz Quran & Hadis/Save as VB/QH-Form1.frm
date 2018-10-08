VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form user 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "„—ﬂ“ ﬁ—¬‰ Ê ÕœÌÀ ﬂ—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„"
   ClientHeight    =   1890
   ClientLeft      =   4620
   ClientTop       =   5085
   ClientWidth     =   4755
   DrawStyle       =   3  'Dash-Dot
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "QH-Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Motor"
      Height          =   2655
      Left            =   3840
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   4335
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
         Connect         =   $"QH-Form1.frx":08CA
         OLEDBString     =   $"QH-Form1.frx":0953
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
         Connect         =   $"QH-Form1.frx":09DC
         OLEDBString     =   $"QH-Form1.frx":0A65
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
         Connect         =   $"QH-Form1.frx":0AEE
         OLEDBString     =   $"QH-Form1.frx":0B77
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
         Connect         =   $"QH-Form1.frx":0C00
         OLEDBString     =   $"QH-Form1.frx":0C89
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
         Connect         =   $"QH-Form1.frx":0D12
         OLEDBString     =   $"QH-Form1.frx":0D9B
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
         Connect         =   $"QH-Form1.frx":0E24
         OLEDBString     =   $"QH-Form1.frx":0EAD
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
         Connect         =   $"QH-Form1.frx":0F36
         OLEDBString     =   $"QH-Form1.frx":0FBF
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
         Left            =   2880
         Top             =   960
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
         Connect         =   $"QH-Form1.frx":1048
         OLEDBString     =   $"QH-Form1.frx":10D1
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
         Height          =   570
         Left            =   2760
         Top             =   360
         Visible         =   0   'False
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   1005
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
         Connect         =   $"QH-Form1.frx":115A
         OLEDBString     =   $"QH-Form1.frx":11E3
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
      Begin MSAdodcLib.Adodc paziresh_table 
         Height          =   330
         Left            =   2760
         Top             =   1320
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
         Connect         =   $"QH-Form1.frx":126C
         OLEDBString     =   $"QH-Form1.frx":12F5
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
      Begin MSAdodcLib.Adodc Setting 
         Height          =   330
         Left            =   2760
         Top             =   1680
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
         Connect         =   $"QH-Form1.frx":137E
         OLEDBString     =   $"QH-Form1.frx":1407
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
   End
   Begin VB.TextBox USERIDTEXT 
      Height          =   1815
      Left            =   6240
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   5760
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ê—Êœ"
      Default         =   -1  'True
      DisabledPicture =   "QH-Form1.frx":1490
      DragMode        =   1  'Automatic
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
      Left            =   3240
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox OP 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      Picture         =   "QH-Form1.frx":F72B
      ScaleHeight     =   1455
      ScaleWidth      =   5280
      TabIndex        =   3
      Top             =   2880
      Width           =   5280
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "ﬂ·„Â ⁄»Ê—"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc DataUser 
      Height          =   330
      Left            =   1080
      Top             =   2400
      Width           =   3000
      _ExtentX        =   5292
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
      Connect         =   $"QH-Form1.frx":11736
      OLEDBString     =   $"QH-Form1.frx":117BF
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·ÿ›« ‰«„ ò«—»—Ì Ê ò·„Â ⁄»Ê— ŒÊœ —« Ê«—œ ò‰Ìœ"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   3450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò·„Â ⁄»Ê—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò«—»—Ì"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   4
      Top             =   720
      Width           =   795
   End
   Begin VB.Menu mnumeue 
      Caption         =   "ò«—»—"
      Visible         =   0   'False
      Begin VB.Menu mnuOzv 
         Caption         =   "⁄÷ÊÌ  œ— ‰—„ «›“«—"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "œ— »«—Â ‰—„ «›“«—"
   End
End
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  


Private Sub Command1_Click()


If Text1 = "" Or Text2 = "" Then
Exit Sub
End If


DataUser.Refresh
st1 = """"
st2 = """"
DataUser.RecordSource = "select * from user1 where username like ('" + Text1.Text + "') and password like ('" + Text2.Text + "') "
'DataUser.RecordSource = "select * from user1 where username = (" & st1 & Text1 & st1 & ") and password = (" & st2 & Text2 & st2 & ")"
'DataUser.RecordSource = "select * from user1 where username = ('%" + Text1.Text + "%') and password = ('%" + Text2.Text + "%') "


DataUser.Refresh
If DataUser.Recordset.EOF = True Then
MsgBox "‰«„ ﬂ«—»—Ì Ê Ì« ﬂ·„Â ⁄»Ê— «‘ »«Â «” ", vbCritical + vbOKOnly, "Œÿ«"
Else



'MsgBox "  Ã‰«» ¬ﬁ«Ì  " & DataUser.Recordset.Fields("op") & "∫" & Chr(10) & "»Â ‰—„ «›“«— „œÌ—Ì  ¬„Ê“‘Ì „—ò“ ﬁ—¬‰ Ê ÕœÌÀ ò—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„ ŒÊ‘ ¬„œÌœ " & Chr(10) & " ÊÃÂ œ«‘ Â »«‘Ìœ  „«„  €ÌÌ—« Ì òÂ «“ «Ì‰ Å” œ— Å«Ìê«Â ’Ê—  „Ì êÌ—œ »Â ‰«„ ‘„« À»  „Ì ‘Êœ", vbInformation + vbOKOnly, "„—ò“ ﬁ—¬‰ Ê ÕœÌÀ ò—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„"
OP.Text = DataUser.Recordset.Fields("op")
Entekhab.useridtext.Text = DataUser.Recordset.Fields("userid")
useridtext.Text = DataUser.Recordset.Fields("userid")
Entekhab.Show
user.Hide
Text2.Text = ""

'Text3.Text = "RezaMohiti"
'SettingF.Show

'SettingF.Hide

End If

End Sub


Private Sub MNU1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub mnuhelp_Click()
WE.Show

End Sub

Private Sub mnuOzv_Click()
MsgBox "»—«Ì ⁄÷ÊÌ  œ— ‰—„ «›“«— »Â „œÌ— „—«Ã⁄Â ò‰Ìœ", vbInformation, "⁄÷ÊÌ "
'ozviat.Show
'user.Hide

End Sub

Private Sub Text1_Click()
Text1 = ""
End Sub

Private Sub Text2_Change()
If Text2.Text = "setting_edit_month" Then
SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "SalJari" + "%')"
SettingUser.Refresh
'MsgBox Right(Combo1.Text, 2)
'Exit Sub

Dim DsS As String

'SettingUser.Refresh
'SettingUser.Recordset.AddNew
'SettingUser.Recordset.Fields("Xcode") = "SalJari"
'SettingUser.Recordset.Fields("status") = "On"
'SettingUser.Recordset.Fields("value") = Right(Combo1.Text, 2)

SettingUser.Recordset.Fields("Xtext") = InputBox("Enter month... For Exampel:" & Chr(10) & "01,02" & Chr(10) & "10,12", "Setting")

SettingUser.Recordset.Update
SettingUser.Refresh
MsgBox "Month is Changed.", vbInformation + vbOKOnly, "Setting_Edit_month"


End If




If Text2.Text = "setting_User" Then
UserProfile.bedoonepassword.Text = "password"
UserProfile.Show
End If



If Text2.Text = "140825131360" Then
OP.Text = "»—‰«„Â ‰ÊÌ”"

'SettingF.Show

Entekhab.Show
user.Hide
Text2.Text = ""

End If

End Sub

Private Sub Text2_Click()
Text2 = ""
End Sub

