VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form SettingEmtahan 
   Caption         =   " ‰ŸÌ„« "
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SettingEmtahan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
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
      Left            =   10560
      TabIndex        =   15
      Top             =   0
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
         Connect         =   $"SettingEmtahan.frx":08CA
         OLEDBString     =   $"SettingEmtahan.frx":0953
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
         Connect         =   $"SettingEmtahan.frx":09DC
         OLEDBString     =   $"SettingEmtahan.frx":0A65
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
         Connect         =   $"SettingEmtahan.frx":0AEE
         OLEDBString     =   $"SettingEmtahan.frx":0B77
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
         Connect         =   $"SettingEmtahan.frx":0C00
         OLEDBString     =   $"SettingEmtahan.frx":0C89
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
         Connect         =   $"SettingEmtahan.frx":0D12
         OLEDBString     =   $"SettingEmtahan.frx":0D9B
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
         Connect         =   $"SettingEmtahan.frx":0E24
         OLEDBString     =   $"SettingEmtahan.frx":0EAD
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
         Connect         =   $"SettingEmtahan.frx":0F36
         OLEDBString     =   $"SettingEmtahan.frx":0FBF
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
         Connect         =   $"SettingEmtahan.frx":1048
         OLEDBString     =   $"SettingEmtahan.frx":10D1
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
         Connect         =   $"SettingEmtahan.frx":115A
         OLEDBString     =   $"SettingEmtahan.frx":11E3
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
         Left            =   2640
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
         Connect         =   $"SettingEmtahan.frx":126C
         OLEDBString     =   $"SettingEmtahan.frx":12F5
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "À»  ‰ ŸÌ„« "
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   1560
         TabIndex        =   24
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   5400
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   7440
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox TE1 
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
         Left            =   6120
         TabIndex        =   8
         Text            =   "Karname-Comment-Xlsx"
         Top             =   360
         Width           =   2760
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
         Left            =   4800
         TabIndex        =   7
         Text            =   "101"
         Top             =   360
         Width           =   720
      End
      Begin VB.TextBox TE4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1365
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   7320
      End
      Begin VB.TextBox TE5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   1560
         TabIndex        =   5
         Top             =   2760
         Width           =   7320
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Õ–›"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ã«Ìê“Ì‰Ì"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FF80&
         Caption         =   "À» "
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÃœÌœ"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   " « ‰„—Â"
         Height          =   300
         Left            =   3000
         TabIndex        =   23
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "«“ ‰„—Â"
         Height          =   300
         Left            =   4920
         TabIndex        =   21
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   " « Ã“¡"
         Height          =   300
         Left            =   6840
         TabIndex        =   19
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«“ Ã“¡"
         Height          =   300
         Left            =   9000
         TabIndex        =   17
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "„ ‰"
         Height          =   315
         Left            =   9120
         TabIndex        =   13
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "⁄‰Ê«‰"
         Height          =   315
         Left            =   4320
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â"
         Height          =   315
         Left            =   5640
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ"
         Height          =   315
         Left            =   9120
         TabIndex        =   10
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
         Height          =   315
         Left            =   8880
         TabIndex        =   9
         Top             =   2760
         Width           =   585
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SettingEmtahan.frx":137E
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7011
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
End
Attribute VB_Name = "SettingEmtahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + "Karname-Comment-Xlsx" + "%')"
Setting.Refresh

Setting.Recordset.Sort = "xsort"

Setting.Recordset.MovePrevious

Setting.Recordset.MoveLast

TE2.Text = Val(Setting.Recordset.Fields("xsort")) + 1


Beep


End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ „Ê—œ —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–›") = vbYes Then

Setting.Recordset.Delete
End If

End Sub

Private Sub Command5_Click()
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «ÿ·«⁄«  À»  ‘Êœ", vbQuestion + vbYesNo, "À»  «ÿ·«⁄« ") = vbNo Then Exit Sub



Setting.Refresh
Setting.Recordset.AddNew
Setting.Recordset.Fields("xkodsetting") = TE1.Text
Setting.Recordset.Fields("xsort") = TE2.Text
Setting.Recordset.Fields("xname") = Combo1.Text
Setting.Recordset.Fields("xtext") = TE4.Text
Setting.Recordset.Fields("tozih") = TE5.Text
Setting.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
'Setting.Recordset.Fields("d") = "@"
Setting.Recordset.Fields("other1") = Text1.Text
Setting.Recordset.Fields("other2") = Text2.Text
Setting.Recordset.Fields("other3") = Text3.Text
Setting.Recordset.Fields("other4") = Text4.Text
Setting.Recordset.Fields("other5") = "@"
Setting.Recordset.Update
Setting.Refresh
MsgBox "«ÿ·«⁄«  À»  ‘œ", vbInformation + vbOKOnly, "À»  «ÿ·«⁄« "

End Sub

Private Sub Command6_Click()

Setting.Recordset.Fields("xkodsetting") = TE1.Text
Setting.Recordset.Fields("xsort") = TE2.Text
Setting.Recordset.Fields("xname") = Combo1.Text
Setting.Recordset.Fields("xtext") = TE4.Text
Setting.Recordset.Fields("tozih") = TE5.Text
Setting.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
'Setting.Recordset.Fields("d") = "@"
Setting.Recordset.Fields("other1") = Text1.Text
Setting.Recordset.Fields("other2") = Text2.Text
Setting.Recordset.Fields("other3") = Text3.Text
Setting.Recordset.Fields("other4") = Text4.Text
Setting.Recordset.Fields("other5") = "@"
Setting.Recordset.Update
Setting.Refresh
'MsgBox "«ÿ·«⁄«  À»  ‘œ", vbInformation + vbOKOnly, "À»  «ÿ·«⁄« "

MsgBox "«ÿ·«⁄«  Ã«Ìê“Ì‰ ‘œ", vbInformation + vbOKOnly, "Ã«Ìê“Ì‰Ì «ÿ·«⁄« "
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next

Setting.Recordset.Update
End Sub

Private Sub DataGrid1_DblClick()
'Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + Setting.Recordset.Fields("xkodsetting") + "%') and xsort like ('%" + Setting.Recordset.Fields("xsort") + "%')"
Setting.Refresh
 TE1.Text = Setting.Recordset.Fields("xkodsetting")
 TE2.Text = Setting.Recordset.Fields("xsort")
 Combo1.Text = Setting.Recordset.Fields("xname")
TE4.Text = Setting.Recordset.Fields("xtext")
TE5.Text = Setting.Recordset.Fields("tozih")
 Text1.Text = Setting.Recordset.Fields("other1")
Text2.Text = Setting.Recordset.Fields("other2")
 Text3.Text = Setting.Recordset.Fields("other3")
  Text4.Text = Setting.Recordset.Fields("other4")
End Sub

Private Sub Form_Load()
Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + "Karname-Comment-Xlsx" + "%')"
Setting.Refresh


Combo1.AddItem ("⁄«·Ì")
Combo1.AddItem ("ŒÊ»")
Combo1.AddItem ("„ Ê”ÿ")
Combo1.AddItem ("÷⁄Ì›")



End Sub

Private Sub Form_Unload(Cancel As Integer)
Karname.Show

End Sub

