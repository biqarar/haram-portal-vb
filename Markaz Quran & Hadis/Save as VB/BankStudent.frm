VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form BankStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "À»  «ÿ·«⁄«  ﬁ—¬‰ ¬„Ê“«‰"
   ClientHeight    =   10050
   ClientLeft      =   5595
   ClientTop       =   1755
   ClientWidth     =   11025
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BankStudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Motor"
      Height          =   255
      Left            =   2280
      TabIndex        =   93
      Top             =   480
      Visible         =   0   'False
      Width           =   735
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
         Connect         =   $"BankStudent.frx":08CA
         OLEDBString     =   $"BankStudent.frx":0953
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
         Connect         =   $"BankStudent.frx":09DC
         OLEDBString     =   $"BankStudent.frx":0A65
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
         Connect         =   $"BankStudent.frx":0AEE
         OLEDBString     =   $"BankStudent.frx":0B77
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
         Connect         =   $"BankStudent.frx":0C00
         OLEDBString     =   $"BankStudent.frx":0C89
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
         Connect         =   $"BankStudent.frx":0D12
         OLEDBString     =   $"BankStudent.frx":0D9B
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
         Connect         =   $"BankStudent.frx":0E24
         OLEDBString     =   $"BankStudent.frx":0EAD
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
         Connect         =   $"BankStudent.frx":0F36
         OLEDBString     =   $"BankStudent.frx":0FBF
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
         Connect         =   $"BankStudent.frx":1048
         OLEDBString     =   $"BankStudent.frx":10D1
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
         Connect         =   $"BankStudent.frx":115A
         OLEDBString     =   $"BankStudent.frx":11E3
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
         Connect         =   $"BankStudent.frx":126C
         OLEDBString     =   $"BankStudent.frx":12F5
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
         Connect         =   $"BankStudent.frx":137E
         OLEDBString     =   $"BankStudent.frx":1407
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
   Begin MSDataGridLib.DataGrid DataGridSTUDENT 
      Bindings        =   "BankStudent.frx":1490
      Height          =   4215
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "»—«Ì Ã«Ìê“Ì‰Ì Å—Ê‰œÂ œÊ »« — »— —ÊÌ ‰«„ «Ê ò·Ìò ò‰Ìœ"
      Top             =   5400
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
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
   Begin VB.Frame Frame10 
      Height          =   735
      Left            =   3240
      TabIndex        =   104
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton Command11 
         Caption         =   "‰„«Ì‘  ’ÊÌ— ﬁ—¬‰ ¬„Ê“"
         Height          =   375
         Left            =   5640
         TabIndex        =   115
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080FF80&
         Caption         =   "ﬁ—¬‰ ¬„Ê“ ÃœÌœ"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox Check_of_sabt_entezar 
         Alignment       =   1  'Right Justify
         Caption         =   "À»  ‰«„ ﬁ—¬‰ ¬„Ê“ œ— ·Ì”  «‰ Ÿ«— »⁄œ «“ À»  «ÿ·«⁄« "
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
         Left            =   6960
         TabIndex        =   110
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080FF80&
         Caption         =   "ﬁ—¬‰ ¬„Ê“ ÃœÌœ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   360
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Frame Frame8 
      Height          =   5295
      Left            =   11040
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ComboBox Combo17 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         TabIndex        =   41
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   435
         Left            =   120
         TabIndex        =   45
         Top             =   4800
         Width           =   2760
      End
      Begin VB.ComboBox Combo15 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "BankStudent.frx":14A6
         Left            =   840
         List            =   "BankStudent.frx":14B9
         TabIndex        =   44
         Text            =   "›⁄·Ì"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.ComboBox Combo16 
         BackColor       =   &H00FFFFFF&
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
         Left            =   840
         TabIndex        =   42
         Top             =   3840
         Width           =   735
      End
      Begin VB.ComboBox Combo14 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   43
         Top             =   3840
         Width           =   735
      End
      Begin VB.ComboBox Combo13 
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
         Height          =   435
         Left            =   2280
         TabIndex        =   108
         Text            =   "00"
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080FF80&
         Caption         =   "’œÊ— —”Ìœ „’«Õ»Â"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3840
         Width           =   735
      End
      Begin VB.ComboBox Combo12 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   38
         Top             =   3360
         Width           =   735
      End
      Begin VB.ComboBox Combo11 
         BackColor       =   &H00FFFFFF&
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
         Left            =   840
         TabIndex        =   39
         Top             =   3360
         Width           =   615
      End
      Begin VB.ComboBox Combo10 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   40
         Top             =   3360
         Width           =   615
      End
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   37
         Top             =   2880
         Width           =   2775
      End
      Begin VB.ComboBox Combo9 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo8 
         BackColor       =   &H00FFFFFF&
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
         Left            =   960
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo7 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "’»Õ"
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄’—"
         Height          =   375
         Left            =   960
         TabIndex        =   30
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "ç—Œ‘Ì"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "„ Ê”ÿ"
         Height          =   330
         Left            =   1560
         TabIndex        =   32
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "÷⁄Ì›"
         Height          =   330
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "ò«„·"
         Height          =   330
         Left            =   1320
         TabIndex        =   34
         Top             =   2040
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«ﬁ’"
         Height          =   330
         Left            =   240
         TabIndex        =   35
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   36
         Text            =   "200,000"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label label_vadie_sabt 
         BackColor       =   &H000000FF&
         Caption         =   "vadie"
         Height          =   375
         Left            =   3600
         TabIndex        =   114
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label sabtentezarlabel 
         BackColor       =   &H000000FF&
         Caption         =   "sabtenrezar"
         Height          =   375
         Left            =   3120
         TabIndex        =   113
         Top             =   4200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label38 
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
         Left            =   3240
         TabIndex        =   112
         Top             =   4800
         Width           =   585
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "œÊ—Â"
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
         Left            =   3480
         TabIndex        =   111
         Top             =   4320
         Width           =   315
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄  „—«Ã⁄Â"
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
         Left            =   3000
         TabIndex        =   107
         Top             =   3840
         Width           =   885
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ „’«Õ»Â"
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
         Left            =   3000
         TabIndex        =   106
         Top             =   3360
         Width           =   870
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "ò«—‘‰«”"
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
         Left            =   3240
         TabIndex        =   105
         Top             =   3000
         Width           =   585
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å–Ì—‘"
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
         Left            =   3000
         TabIndex        =   103
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "“„«‰ ÅÌ‘‰Â«œÌ"
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
         Left            =   2880
         TabIndex        =   102
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "”ÿÕ òÌ›Ì"
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
         Left            =   3120
         TabIndex        =   101
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷⁄Ì  Å—œ«Œ  ÊœÌ⁄Â"
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
         Left            =   2400
         TabIndex        =   100
         Top             =   2040
         Width           =   1410
      End
      Begin VB.Label Label30 
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
         Left            =   3000
         TabIndex        =   99
         Top             =   2520
         Width           =   810
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   97
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TE17 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   350
      Left            =   120
      TabIndex        =   95
      Top             =   3720
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "À»  ﬁ—¬‰ ¬„Ê“ œ— ·Ì”  ò·«”Ì"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "À»  Å—œ«Œ  ÊœÌ⁄Â"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   12360
      TabIndex        =   94
      Top             =   8040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   92
      Top             =   9675
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
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
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ã«Ìê“Ì‰Ì Å—Ê‰œÂ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   4680
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar PB1 
      Height          =   135
      Left            =   240
      TabIndex        =   90
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Å«ò ò—œ‰ ›—„"
      Height          =   660
      Left            =   10080
      TabIndex        =   48
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   480
      TabIndex        =   88
      Top             =   10320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   "¬„«—"
      Height          =   735
      Left            =   4080
      TabIndex        =   85
      Top             =   4560
      Width           =   2895
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   300
         Left            =   240
         TabIndex        =   87
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò· ﬁ—¬‰ ¬„Ê“«‰"
         Height          =   300
         Left            =   1200
         TabIndex        =   86
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   120
      TabIndex        =   83
      Top             =   3840
      Width           =   10815
      Begin VB.TextBox TE16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   9720
      End
      Begin VB.Label Label17 
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
         Left            =   9960
         TabIndex        =   84
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "„‘Œ’«  œÊ—Â"
      Height          =   735
      Left            =   3360
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
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
         ForeColor       =   &H00004000&
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   435
         ItemData        =   "BankStudent.frx":14E2
         Left            =   3360
         List            =   "BankStudent.frx":14E4
         TabIndex        =   1
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   240
         Width           =   1575
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
         Left            =   2760
         TabIndex        =   81
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "·ÿ›« œÊ—Â „Ê—œ ‰Ÿ— —« «‰Œ«» ò‰Ìœ"
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
         Left            =   5160
         TabIndex        =   80
         Top             =   240
         Width           =   2145
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "„‘Œ’«  «”ò‰ ›«Ì·"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   11760
      TabIndex        =   76
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CheckBox Check3 
         Caption         =   "»·Ì"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   77
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "⁄„·Ì«  «”ﬂ‰ «‰Ã«„ ‘œÂø"
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
         Left            =   4800
         TabIndex        =   78
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   1335
      Left            =   11760
      TabIndex        =   72
      Top             =   7680
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   420
         Left            =   120
         TabIndex        =   82
         Text            =   "«‰ Œ«» ò‰Ìœ"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TE15 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         Left            =   120
         TabIndex        =   75
         Text            =   "‰œ«—œ"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "òœ ò·«”"
         Height          =   255
         Left            =   1680
         TabIndex        =   74
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
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
         Left            =   1680
         TabIndex        =   73
         Top             =   480
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Õ’Ì·«  ° ‘„«—Â  „«”"
      Height          =   1215
      Left            =   3960
      TabIndex        =   68
      Top             =   2640
      Width           =   6975
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3960
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox TE14 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2400
      End
      Begin VB.TextBox TE13 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2400
      End
      Begin VB.TextBox TE12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   3960
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â  ·›‰ Â„—«Â"
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
         Left            =   2640
         TabIndex        =   71
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   " Õ’Ì·« "
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
         TabIndex        =   70
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â  ·›‰ À«» "
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
         Left            =   2640
         TabIndex        =   69
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "„‘Œ’«  ›—œÌ"
      Height          =   1935
      Left            =   120
      TabIndex        =   55
      Top             =   720
      Width           =   10815
      Begin VB.CheckBox chj 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ã” ÃÊ"
         Height          =   300
         Left            =   2400
         TabIndex        =   49
         Top             =   1320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.OptionButton Moj 
         Caption         =   "„Ã—œ"
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
         Left            =   9000
         TabIndex        =   13
         Top             =   1500
         Width           =   735
      End
      Begin VB.OptionButton mot 
         Caption         =   "„ «Â·"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox TE11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   5520
         TabIndex        =   15
         Text            =   "0"
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox TE10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   4560
         TabIndex        =   12
         Top             =   1080
         Width           =   2160
      End
      Begin VB.TextBox TE9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   7800
         TabIndex        =   11
         Top             =   1080
         Width           =   2160
      End
      Begin VB.TextBox TE8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   240
         TabIndex        =   10
         Text            =   "‘Ì⁄Â"
         Top             =   720
         Width           =   1440
      End
      Begin VB.TextBox TE7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   2760
         TabIndex        =   9
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox TE6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   5520
         TabIndex        =   8
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox TE5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   7800
         TabIndex        =   7
         Top             =   720
         Width           =   1680
      End
      Begin VB.TextBox TE4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1440
      End
      Begin VB.TextBox TE3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox TE2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   2160
      End
      Begin VB.TextBox TE1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   350
         Left            =   7800
         TabIndex        =   3
         Top             =   360
         Width           =   2400
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   3240
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "œ— Â— òœ«„ «“ „Ê«—œ »«·« Ã” ÃÊ ò‰Ìœ"
         Height          =   300
         Left            =   120
         TabIndex        =   89
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label2 
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
         Left            =   10320
         TabIndex        =   67
         Top             =   360
         Width           =   195
      End
      Begin VB.Label Label3 
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
         Left            =   6840
         TabIndex        =   66
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label4 
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
         Left            =   3960
         TabIndex        =   65
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label5 
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
         Left            =   1920
         TabIndex        =   64
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â ‘‰«”‰«„Â"
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
         Left            =   9600
         TabIndex        =   63
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "’«œ—Â «“"
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
         Left            =   7200
         TabIndex        =   62
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "„·Ì "
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
         Left            =   4080
         TabIndex        =   61
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "„–Â»"
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
         Left            =   2160
         TabIndex        =   60
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ﬂœ „·Ì"
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
         Left            =   10080
         TabIndex        =   59
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â ê–—‰«„Â"
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
         Left            =   6840
         TabIndex        =   58
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷⁄Ì   «Â·"
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
         Left            =   9840
         TabIndex        =   57
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ›—“‰œ«‰"
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
         Left            =   6840
         TabIndex        =   56
         Top             =   1440
         Width           =   930
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "À»  «ÿ·«⁄&« "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H00C0FFC0&
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
      Left            =   8280
      TabIndex        =   54
      Top             =   11760
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H00C0FFC0&
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
      Left            =   8280
      TabIndex        =   53
      Top             =   11280
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox TEP 
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
      Height          =   350
      Left            =   240
      TabIndex        =   50
      Text            =   "À»   Ê”ÿ ”Ì” „"
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "¬œ—” ›«Ì· "
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
      Left            =   3000
      TabIndex        =   96
      Top             =   3600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "»—«Ì «’·«Õ „‘Œ’«  ﬁ—¬‰ ¬„Ê“ »— —ÊÌ ‰«„ «Ê œÊ »«— ò·Ìœ ò‰Ìœ Ê ”Å” Ã«Ìê“Ì‰Ì Å—Ê‰œÂ —« »“‰Ìœ"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   7080
      TabIndex        =   91
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   2520
      X2              =   2520
      Y1              =   4680
      Y2              =   5280
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "‰«„ Ê«—œ ﬂ‰‰œÂ «ÿ·«⁄« "
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
      Left            =   11520
      TabIndex        =   52
      Top             =   11760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "ﬂœ ﬂ·«”"
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
      Left            =   11520
      TabIndex        =   51
      Top             =   11280
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
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
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
   Begin VB.Menu mnuhome 
      Caption         =   "#"
   End
   Begin VB.Menu mnufile 
      Caption         =   "Å—Ê ‰œÂ"
      Begin VB.Menu mnusabte 
         Caption         =   "À»  «ÿ·«⁄« "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "Ã«Ìê“Ì‰Ì Å—Ê‰œÂ"
      End
      Begin VB.Menu sssss 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnu_sabl_entezar 
         Caption         =   "À»  ﬁ—¬‰ ¬„Ê“ œ— ·Ì”  «‰ Ÿ«—"
      End
   End
   Begin VB.Menu mnuoption 
      Caption         =   " ‰ŸÌ„« "
      Begin VB.Menu mnudelete 
         Caption         =   "Õ–› ﬁ—¬‰ ¬„Ê“"
      End
   End
End
Attribute VB_Name = "BankStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EYVAL As Integer
Dim Xzaman_pishnahad_str As String


Private Sub chj_Click()
If chj.Value = 1 Then
Label28.Visible = True

Else
Label28.Visible = False
End If
End Sub

Private Sub chsave_Click()
If chsave.Value = 1 Then
Save.Visible = True

Else
Save.Visible = False
End If


End Sub

Private Sub Combo1_Click()
On Error Resume Next




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
Tarhha.RecordSource = "select * from tarhha where NAME like ('%" + Combo1.Text + "%')" ' „«„ Å—Ê‰œÂ Â« —« „Ì ê—œœ Ê ò”«‰Ì òÂ ÿ— Õ «‰ Œ«» ‘œÂ —« œ«—‰œ „Ì ¬Ê—œÅ
Tarhha.Refresh
    KODTARH = Tarhha.Recordset.Fields("XkodDORE")
    A = Val(SaljariSTR & CodeBakhsh & KODTARH & "000")
    B = Val(SaljariSTR & CodeBakhsh & KODTARH & "999")
    KK = SaljariSTR & CodeBakhsh & KODTARH
    Text1.Text = KK
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + Text1.Text + "%')"
Student.Refresh
Student.Recordset.MoveFirst
Student.Refresh

K = Student.Recordset.Fields("parvande")
PB1.Visible = True
PB1.Value = 1

PB1.Max = Student.Recordset.RecordCount

For J = 1 To Student.Recordset.RecordCount
F = Val(Student.Recordset.Fields("parvande"))
    If F > A Then '»—«Ì  «Ì‰òÂ ⁄œœ œÌê—Ì ﬁ«ÿÌ òœ ‰‘Êœ
    If F < B Then
   
   
             If F > SSS Then
            SSS = F ' »“—ê —Ì‰ —« ÅÌœ« „Ì ò‰œ
             Else
             GoTo 14
             End If
    End If
    End If
14     Student.Recordset.MoveNext
PB1.Value = PB1.Value + 1
Next J
If SSS = 0 Then
TEP.Text = SaljariSTR & CodeBakhsh & KODTARH & "001"
Else

TEP = SSS + 1 ''‰ ÌçÂ ‰Â«ÌÌ
End If

PB1.Visible = False

'MnusabtTanha.Enabled = True
End Sub

Private Sub Combo2_Click()
On Error GoTo 10
Combo1.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where goroh like ('" & Combo2.Text & "')"
Tarhha.Refresh

Tarhha.Recordset.Sort = "sortname"

For I = 1 To Tarhha.Recordset.RecordCount
Combo1.AddItem (Tarhha.Recordset.Fields("name"))
Tarhha.Recordset.MoveNext
Next I
Combo1.Text = Combo1.List(0)


10 Exit Sub

If Combo2.Text = "⁄„Ê„Ì" Then
Combo1.Clear

Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "1" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo1.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo1.Text = Combo1.List(0)
End If


If Combo2.Text = "ò«—ê«Â Â«" Then
Combo1.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "3" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo1.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo1.Text = Combo1.List(0)
End If

If Combo2.Text = " —»Ì  „—»Ì" Then
Combo1.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "4" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo1.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo1.Text = Combo1.List(0)
End If


If Combo2.Text = "Õ›Ÿ ﬁ—¬‰ ò—Ì„" Then

Combo1.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "2" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo1.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo1.Text = Combo1.List(0)
End If

If Combo2.Text = "„ÃÂÊ·" Then
Combo1.Clear
Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where dore like ('%" + "0" + "%')"
Tarhha.Refresh
For I = 1 To Tarhha.Recordset.RecordCount
Combo1.AddItem (Tarhha.Recordset.Fields("tarhname"))
Tarhha.Recordset.MoveNext
Next I
Combo1.Text = Combo1.List(0)
End If

End Sub

Private Sub Combo3_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where ostad like ('%" + Combo3.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub Combo5_KeyUp(KeyCode As Integer, Shift As Integer)
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where tahsilat like ('%" + Combo5.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub Command1_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "bank-newsabt" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub

14082513



'If Check_of_sabt_entezar.Value = 1 Then

'If Combo12.Text = "" Or Combo4.Text = "" Or Combo6.Text = "" Or Combo15.Text = "" Or Combo13.Text = "" Or Combo14.Text = "" Or Combo16.Text = "" Or Combo17.Text = "" Or Combo10.Text = "" Or Combo11.Text = "" Then
'MsgBox "Œÿ« œ— Ê—ÊœÌ «ÿ·«⁄« " & Chr(10) & "»⁄÷Ì «“ ﬁ”„  Â« Å— ‰‘œÂ «” ", vbCritical + vbOKOnly, "Œÿ«"
'Exit Sub
'End If

'End If



'If Entekhab.SB.Panels(1).Text = "„ÌÂ„«‰" Then Exit Sub
If TE8.Text <> "‘Ì⁄Â" Then
MsgBox "«„ò«‰ À»  ‰«„ ﬁ—¬‰ ¬„Ê“ »« €Ì— «“ „–Â» ‘Ì⁄Â ÊÃÊœ ‰œ«—œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If

If TEP.Text = "À»   Ê”ÿ ”Ì” „" Or TEP.Text = "" Or TE1.Text = "" Or TE2.Text = "" Then  'Œÿ« œ— Ê—ÊœÌ ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ Ê Ì« ‘„«—Â Å—Ê‰œÂ ’ÕÌÕ ‰”Ì 

MsgBox "Œÿ« œ— Ê—ÊœÌ: ·ÿ›« ‘„«—Â Å—Ê‰œÂ° ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ —« »——”Ì òÌ‰œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If
If Moj.Value = False And mot.Value = False Then
MsgBox "Ê÷⁄Ì   «Â· —« „‘Œ’ òÌ‰œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If

If TEP = "" Then '«ê— ‘„«—Â Å—Ê‰œÂ Œ«·Ì »Êœ
MsgBox "‘„«—Â Å—Ê‰œÂ »«Ìœ Ê«—œ ‘Êœ", vbExclamation, "Œÿ«"
Else '«ê— ‘„«—Â Å—Ê‰œÂ Œ«·Ì ‰»Êœ Ê ﬂ«—»— ¬‰ —« Å— ﬂ—œÂ »Êœ
'⁄œ„  ò—«—Ì »Êœ‰ ‘„«—Â  Å—Ê‰œÂ
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + TEP.Text + "%')"
Student.Refresh
If Val(Student.Recordset.RecordCount) >= 1 Then
MsgBox "‘„«—Â Å—Ê‰œÂ  ò—«—Ì «” ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
Else
GoTo 1
End If
1 'Å«Ì«‰ ⁄œ„  ò—«—Ì »Êœ‰ ‘„«—Â Å—Ê‰œÂ

'   ‘—Ê⁄  ”   ò—«—Ì »Êœ‰ ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ
' „«„ ﬁ—¬‰ ¬„Ê“«‰ «‰ Œ«» „Ì ‘Êœ
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + "" + "%')"
Student.Refresh
'Å«Ì«‰ «‰ Œ«»  „«„ ﬁ—¬‰ ¬„Ê“«‰
Student.RecordSource = "select * from student where name like ('%" + TE1.Text + "%') and famil like ('%" + TE2.Text + "%') "
Student.Refresh
If Val(Student.Recordset.RecordCount) >= 1 Then
MsgBox "‰«„ «Ì‰ ﬁ—¬‰ ¬„Ê“ ﬁ»·« À»  ‘œÂ «”  " & Chr(10) & "  „Ê—œ Ì«›  ‘œ  " & Student.Recordset.RecordCount, vbExclamation + vbOKOnly, "Â‘œ«—"
Beep

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ œ— Â— ’Ê—  ‰«„ ﬁ—¬‰ ¬„Ê“ À»  ‘Êœ ", vbQuestion + vbYesNo, "À»  «ÿ·«⁄«  ﬁ—¬‰ ¬„Ê“«‰") = vbYes Then
GoTo 2
Else
Exit Sub
End If ' »—«Ì ÅÌ«„Ì òÂ „Ì Å—”œ ¬Ì« „Ì ŒÊ«ÂÌœ À»  ò‰Ìœ


End If
2
chj.Value = 0
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + "" + "%')"
Student.Refresh



'   Å«Ì«‰  ”   ò—«—Ì »Êœ‰ ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ
    




Student.Refresh
Student.Recordset.AddNew
' «ÿ·«⁄«  Ê«—œ ”Ì” „ „Ì ‘Ê‰œ


Student.Recordset.Fields("Parvande") = TEP
Student.Recordset.Fields("name") = TE1
Student.Recordset.Fields("famil") = TE2
Student.Recordset.Fields("NF") = TE1.Text & " " & TE2.Text

Student.Recordset.Fields("NamePedar") = TE3
Student.Recordset.Fields("tavalod") = TE4
If TE5 = "" Then
Student.Recordset.Fields("shsh") = "0"
Else
Student.Recordset.Fields("shsh") = TE5
End If

Student.Recordset.Fields("sadere") = TE6
Student.Recordset.Fields("meliyat") = TE7
Student.Recordset.Fields("mazhab") = TE8
Student.Recordset.Fields("Kodmeli") = TE9
Student.Recordset.Fields("gozarname") = TE10
If Moj.Value = True Then Student.Recordset.Fields("taahol") = Moj.Caption

If mot.Value = True Then Student.Recordset.Fields("taahol") = mot.Caption
If TE11 = "" Then TE11 = "0"

Student.Recordset.Fields("farzand") = TE11

Student.Recordset.Fields("tahsilat") = Combo5

Student.Recordset.Fields("ostad") = Combo3.Text
Student.Recordset.Fields("tozih") = TE16
If TE13 = "" Then TE13 = "0"

Student.Recordset.Fields("tell") = TE13
If TE14 = "" Then TE14 = "0"

Student.Recordset.Fields("mob") = TE14


Student.Recordset.Fields("scan") = TE17.Text
Student.Recordset.Fields("XSELECT") = "0"


Student.Recordset.Fields("xkard") = "Kard-New"




Student.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
Student.Recordset.Fields("d") = Taqvim.Tarikh.Caption


Student.Recordset.Fields("clas1") = "‰œ«—œ"
Student.Recordset.Fields("clas2") = "‰œ«—œ"
Student.Recordset.Fields("clas3") = "‰œ«—œ"
Student.Recordset.Fields("clas4") = "‰œ«—œ"
Student.Recordset.Fields("clas5") = "‰œ«—œ"

'«ÿ·«⁄«  Ê«—œ ”Ì” „ „Ì ‘Ê‰œ
Student.Recordset.Update


Student.Refresh

'Barcode.Text = TEP.Text
'Dim DF As String
'DF = "F:\Markaz Quran & Hadis\Bar Cod\" & "RM" & TEP.Text & ".png"

'Barcode.SaveAsPNGBySize DF, 1425, 730












MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbInformation, "À»  «ÿ·«⁄«  ﬁ—¬‰ ¬„Ê“«‰"
TE1.Text = ""
TE2.Text = ""
TE3.Text = ""
TE4.Text = ""
TE5.Text = ""
TE6.Text = ""
TE7.Text = ""
'TE8.Text = ""
TE9.Text = ""
TE10.Text = ""
TE11.Text = ""
Combo5.Text = ""
TE13.Text = ""
TE14.Text = ""
TE15.Text = ""
TE16.Text = ""
TE17.Text = ""

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & "" & "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount
'Combo1.SetFocus
chj.Value = 1
'jostojo baraye shomare parvadne

'Call TEP_Change

'If Check_of_sabt_entezar.Value = 1 Then Call sabtentezarlabel_Click



End If
chj.Value = 1

End Sub

Private Sub Command10_Click()
On Error Resume Next




Dim A, B, KK, K, F, T, SSS As Long

Dim SaljariSTR, CodeBakhsh, Mahjari As String


SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "SalJari" + "%')"
SettingUser.Refresh

SaljariSTR = SettingUser.Recordset.Fields("value")
Mahjari = SettingUser.Recordset.Fields("xtext")
SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "CodeBakhsh" + "%')"
SettingUser.Refresh
CodeBakhsh = SettingUser.Recordset.Fields("value")

'Tarhha.Refresh
'Tarhha.RecordSource = "select * from tarhha where NAME like ('%" + Combo1.Text + "%')" ' „«„ Å—Ê‰œÂ Â« —« „Ì ê—œœ Ê ò”«‰Ì òÂ ÿ— Õ «‰ Œ«» ‘œÂ —« œ«—‰œ „Ì ¬Ê—œÅ
'Tarhha.Refresh
   ' KODTARH = Tarhha.Recordset.Fields("XkodDORE")
    'A = Val(SaljariSTR & CodeBakhsh & KODTARH & "000")
    'B = Val(SaljariSTR & CodeBakhsh & KODTARH & "999")
    
    A = Val(SaljariSTR & CodeBakhsh & Mahjari & "000")
    B = Val(SaljariSTR & CodeBakhsh & Mahjari & "999")
    KK = SaljariSTR & CodeBakhsh & Mahjari
    'serfan vaseye jostojo
    Text1.Text = KK
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + Text1.Text + "%')"
Student.Refresh
Student.Recordset.MoveFirst
Student.Refresh

K = Student.Recordset.Fields("parvande")
PB1.Visible = True
PB1.Value = 1

PB1.Max = Student.Recordset.RecordCount

For J = 1 To Student.Recordset.RecordCount
F = Val(Student.Recordset.Fields("parvande"))
    If F > A Then '»—«Ì  «Ì‰òÂ ⁄œœ œÌê—Ì ﬁ«ÿÌ òœ ‰‘Êœ
    If F < B Then
   
   
             If F > SSS Then
            SSS = F ' »“—ê —Ì‰ —« ÅÌœ« „Ì ò‰œ
             Else
             GoTo 14
             End If
    End If
    End If
14     Student.Recordset.MoveNext
PB1.Value = PB1.Value + 1
Next J
If SSS = 0 Then
TEP.Text = SaljariSTR & CodeBakhsh & Mahjari & "001"
Else

TEP = SSS + 1 ''‰ ÌçÂ ‰Â«ÌÌ
End If

PB1.Visible = False

'MnusabtTanha.Enabled = True

End Sub

Private Sub Command11_Click()
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

Private Sub Command3_Click()
If chj.Value = 0 Then

TE1.Text = ""
TE2.Text = ""
TE3.Text = ""
TE4.Text = ""
TE5.Text = ""
TE6.Text = ""
TE7.Text = ""
'TE8.Text = ""
TE9.Text = ""
TE10.Text = ""
TE11.Text = ""
Combo5.Text = ""
TE13.Text = ""
TE14.Text = ""
TE15.Text = ""
TE16.Text = ""
TE17.Text = ""
End If

If chj.Value = 1 Then
chj.Value = 0
TE1.Text = ""
TE2.Text = ""
TE3.Text = ""
TE4.Text = ""
TE5.Text = ""
TE6.Text = ""
TE7.Text = ""
'TE8.Text = ""
TE9.Text = ""
TE10.Text = ""
TE11.Text = ""
Combo5.Text = ""
TE13.Text = ""
TE14.Text = ""
TE15.Text = ""
TE16.Text = ""
TE17.Text = ""
chj.Value = 1
End If

End Sub

Private Sub Command2_Click()
Dim DF As String
DF = "D:\" & "RM" & Student.Recordset.Fields("PARVANDE") & ".PNG"

Barcode.SaveAsPNG DF






End Sub

Private Sub Command4_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "bank-edit" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub


14082513


If TE8.Text <> "‘Ì⁄Â" Then
MsgBox "«„ò«‰ À»  ‰«„ ﬁ—¬‰ ¬„Ê“ »« €Ì— «“ „–Â» ‘Ì⁄Â ÊÃÊœ ‰œ«—œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If


If EYVAL = 0 Then
MsgBox "«» œ« »«Ìœ »« œÊ »«— ò·Ìò »— —ÊÌ ‰«„ ﬁ—¬‰ ¬„Ê“ Å—Ê‰œÂ «Ì‘«‰ —« »—«Ì Ã«Ìê“Ì‰Ì ¬„«œÂ òÌ‰œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
Else
GoTo 1
End If
Exit Sub
1:
EYVAL = 0



Beep
'Student.Refresh

'Student.RecordSource = "select * from student where name like ('%" + TE1.Text + "%') and famil like ('%" + TE2.Text + "%') "
'Student.Refresh
Student.Refresh

Student.RecordSource = "select * from student where parvande like ('%" + TEP.Text + "%') "
Student.Refresh
'If Student.Recordset.RecordCount > 1 Then
'MsgBox "»—«Ì «Ì‰ ‰«„ »Ì‘ — «“ Ìò „Ê—œ Å—Ê‰œÂ Ì«›  ‘œÂ «”  ·ÿ›« œ— Ã«Ìê“Ì‰Ì Å—Ê‰œÂ œﬁ  ò‰Ìœ", vbExclamation + vbOKOnly, "Â‘œ«—"
'End If


If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ Å—Ê‰œÂ ¬ﬁ«Ì  " & Student.Recordset.Fields("famil") & " Ã«Ìê“Ì‰ ‘Êœ ", vbQuestion + vbYesNo, "Ã«Ìê“Ì‰Ì Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“") = vbYes Then

'Student.RecordSource = "select * from student where parvande like ('%" + PS.Caption + "%')"
'Student.Refresh



' «ÿ·«⁄«  Ê«—œ ”Ì” „ „Ì ‘Ê‰œ


Student.Recordset.Fields("Parvande") = TEP
Student.Recordset.Fields("name") = TE1
Student.Recordset.Fields("famil") = TE2
Student.Recordset.Fields("NamePedar") = TE3
Student.Recordset.Fields("tavalod") = TE4
If TE5 = "" Then
Student.Recordset.Fields("shsh") = "0"
Else
Student.Recordset.Fields("shsh") = TE5
End If

Student.Recordset.Fields("sadere") = TE6
Student.Recordset.Fields("meliyat") = TE7
Student.Recordset.Fields("mazhab") = TE8
Student.Recordset.Fields("Kodmeli") = TE9
Student.Recordset.Fields("gozarname") = TE10
If Moj.Value = True Then Student.Recordset.Fields("taahol") = Moj.Caption

If mot.Value = True Then Student.Recordset.Fields("taahol") = mot.Caption
If TE11 = "" Then TE11 = "0"

Student.Recordset.Fields("farzand") = TE11

Student.Recordset.Fields("tahsilat") = Combo5

Student.Recordset.Fields("ostad") = Combo3.Text
Student.Recordset.Fields("tozih") = TE16
If TE13 = "" Then TE13 = "0"

Student.Recordset.Fields("tell") = TE13
If TE14 = "" Then TE14 = "0"

Student.Recordset.Fields("mob") = TE14

Student.Recordset.Fields("NF") = TE1.Text & " " & TE2.Text
Student.Recordset.Fields("d") = Taqvim.Tarikh.Caption


Student.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text


Student.Recordset.Fields("scan") = TE17

'Student.Recordset.Fields("clas1") = TE15.Text


'Student.Recordset.Fields("clas2") = "‰œ«—œ"
'Student.Recordset.Fields("clas3") = "‰œ«—œ"
'Student.Recordset.Fields("clas4") = "‰œ«—œ"
'Student.Recordset.Fields("clas5") = "‰œ«—œ"

'«ÿ·«⁄«  Ê«—œ ”Ì” „ „Ì ‘Ê‰œ
Student.Recordset.Update



MsgBox "«ÿ·«⁄«  Ã«Ìê“Ì‰ ‘œ", vbInformation, "Ã«Ìê“Ì‰Ì «ÿ·«⁄« "





chj.Value = 1






End If

End Sub

Private Sub Command5_Click()
VadieF.Show
VadieF.Text1.Text = TEP.Text
VadieF.Combo4.Text = Combo2.Text
VadieF.Combo3.Text = Combo1.Text

End Sub

Private Sub Command6_Click()
FClassroom.Show
FClassroom.Text1.Text = TEP.Text

End Sub

Private Sub Command7_Click()
FFOF.Show

End Sub

Private Sub Command8_Click()
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String
'Student.Recordset.MoveFirst
'On Error GoTo 1
GoTo 2

1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"

Exit Sub

2:



paziresh_table.Refresh
paziresh_table.RecordSource = "select * from paziresh_table where parvande like ('%" & TEP.Text & "%')"
paziresh_table.Refresh


Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + TEP.Text + "%')"
Student.Refresh


If paziresh_table.Recordset.BOF = True Or paziresh_table.Recordset.EOF = True Then
'kasi peyda nashode
MsgBox "‰«„ «Ì‰ ﬁ—¬‰ ¬„Ê“ œ— ·Ì”  «‰Ÿ«— ÊÃÊœ œ«—œ", vbInformation, "·Ì”  «‰ Ÿ«—"
Exit Sub
End If



If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "mosahebe.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "mosahebe.xlsx")
End If

oExcel.ActiveSheet.Range("b5").Value = Student.Recordset.Fields("parvande")
oExcel.ActiveSheet.Range("b3").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("b4").Value = Student.Recordset.Fields("NamePedar")
oExcel.ActiveSheet.Range("b6").Value = paziresh_table.Recordset.Fields("tarikh_morajee")
oExcel.ActiveSheet.Range("b7").Value = paziresh_table.Recordset.Fields("saat_morajee")
oExcel.ActiveSheet.Range("b8").Value = paziresh_table.Recordset.Fields("tahvil_be")

'tarikh_paziresh
oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True
'

End Sub

Private Sub Command9_Click()




Dim A, B, KK, K, F, T, SSS As Long

Dim SaljariSTR, CodeBakhsh, Mahjari As String


SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "SalJari" + "%')"
SettingUser.Refresh

SaljariSTR = SettingUser.Recordset.Fields("value")
Mahjari = SettingUser.Recordset.Fields("xtext")


SettingUser.Refresh
SettingUser.RecordSource = "select * from settinguser where xcode like ('%" + "CodeBakhsh" + "%')"
SettingUser.Refresh
CodeBakhsh = SettingUser.Recordset.Fields("value")
A = SaljariSTR & CodeBakhsh & Mahjari

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & A & "%')"
Student.Refresh
If Student.Recordset.BOF = True Or Student.Recordset.EOF = True Then
'yani nafar avale ke toye ein mah mikhad sabt name kone
TEP.Text = SaljariSTR & Mahjari & "001"
Beep

Exit Sub

End If

Student.Recordset.Sort = "parvande"
Student.Recordset.MovePrevious

Student.Recordset.MoveLast

TEP.Text = Val(Student.Recordset.Fields("parvande")) + 1
Beep


Exit Sub

Tarhha.Refresh
Tarhha.RecordSource = "select * from tarhha where NAME like ('%" + Combo1.Text + "%')" ' „«„ Å—Ê‰œÂ Â« —« „Ì ê—œœ Ê ò”«‰Ì òÂ ÿ— Õ «‰ Œ«» ‘œÂ —« œ«—‰œ „Ì ¬Ê—œÅ
Tarhha.Refresh
    KODTARH = Tarhha.Recordset.Fields("XkodDORE")
    A = Val(SaljariSTR & CodeBakhsh & KODTARH & "000")
    B = Val(SaljariSTR & CodeBakhsh & KODTARH & "999")
    KK = SaljariSTR & CodeBakhsh & KODTARH
    Text1.Text = KK
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + Text1.Text + "%')"
Student.Refresh
Student.Recordset.MoveFirst
Student.Refresh

K = Student.Recordset.Fields("parvande")
PB1.Visible = True
PB1.Value = 1

PB1.Max = Student.Recordset.RecordCount

For J = 1 To Student.Recordset.RecordCount
F = Val(Student.Recordset.Fields("parvande"))
    If F > A Then '»—«Ì  «Ì‰òÂ ⁄œœ œÌê—Ì ﬁ«ÿÌ òœ ‰‘Êœ
    If F < B Then
   
   
             If F > SSS Then
            SSS = F ' »“—ê —Ì‰ —« ÅÌœ« „Ì ò‰œ
             Else
             GoTo 14
             End If
    End If
    End If
14     Student.Recordset.MoveNext
PB1.Value = PB1.Value + 1
Next J
If SSS = 0 Then
TEP.Text = SaljariSTR & CodeBakhsh & KODTARH & "001"
Else

TEP = SSS + 1 ''‰ ÌçÂ ‰Â«ÌÌ
End If

PB1.Visible = False

'MnusabtTanha.Enabled = True












End Sub

Private Sub DataGridSTUDENT_DblClick()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "bank-stu-dbl-edit" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
If Student.Recordset.RecordCount < 1 Then Exit Sub

chj.Value = 0
Beep
EYVAL = 1
On Error Resume Next


TEP.Text = Student.Recordset.Fields("Parvande")
TE1.Text = Student.Recordset.Fields("name")
TE2.Text = Student.Recordset.Fields("famil")


 TE3 = Student.Recordset.Fields("NamePedar")
 TE4 = Student.Recordset.Fields("tavalod")

TE5 = Student.Recordset.Fields("shsh")


 TE6 = Student.Recordset.Fields("sadere")
 TE7 = Student.Recordset.Fields("meliyat")
 TE8 = Student.Recordset.Fields("mazhab")
  TE9 = Student.Recordset.Fields("Kodmeli")
 TE10 = Student.Recordset.Fields("gozarname")
If Student.Recordset.Fields("taahol") = Moj.Caption Then Moj.Value = True

If Student.Recordset.Fields("taahol") = mot.Caption Then mot.Value = True

 TE11 = Student.Recordset.Fields("farzand")

 Combo5.Text = Student.Recordset.Fields("tahsilat")

 Combo3.Text = Student.Recordset.Fields("ostad")
TE16 = Student.Recordset.Fields("tozih")


  TE13 = Student.Recordset.Fields("tell")


 TE14 = Student.Recordset.Fields("mob")


' TE15.Text = Student.Recordset.Fields("clas1")

 TE17.Text = Student.Recordset.Fields("scan")


'Student.Recordset.Fields("clas2") = Student.Recordset.Fields("clas2")
'Student.Recordset.Fields("clas3") = Student.Recordset.Fields("clas3")
'Student.Recordset.Fields("clas4") = Student.Recordset.Fields("clas4")
'Student.Recordset.Fields("clas5") = Student.Recordset.Fields("clas5")





















End Sub

Private Sub Form_Load()
GoTo 1
teacher.Refresh
For I = 1 To teacher.Recordset.RecordCount
Combo3.AddItem (teacher.Recordset.Fields("Famil"))
teacher.Recordset.MoveNext
Next I
MnusabtTanha.Enabled = False
stb1.Panels(1).Text = user.OP.Text

1
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + "" + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount
EYVAL = 0



Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "stu-tahsilat" & "%') "
Setting.Refresh
Combo5.Clear

 For I = 1 To Setting.Recordset.RecordCount
 Combo5.AddItem (Setting.Recordset.Fields("xtext"))
Setting.Recordset.MoveNext
Next I



Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "stu-karshenas" & "%') "
Setting.Refresh
Combo6.Clear

 For I = 1 To Setting.Recordset.RecordCount
 Combo6.AddItem (Setting.Recordset.Fields("xtext"))
Setting.Recordset.MoveNext
Next I

Combo17.AddItem ("‘‰»Â")
Combo17.AddItem ("Ìò ‘‰»Â")
Combo17.AddItem ("œÊ ‘‰»Â")
Combo17.AddItem ("”Â ‘‰»Â")
Combo17.AddItem ("çÂ«— ‘‰»Â")
Combo17.AddItem ("Å‰Ã ‘‰»Â")
Combo17.AddItem ("Ã„⁄Â")



For I = 1 To 20
Combo4.AddItem (I & "0,000")
Next I

stb1.Panels(1).Text = user.OP.Text
stb1.Panels(3).Text = Taqvim.Label1.Caption
 stb1.Panels(5).Text = Taqvim.KKK.Caption
 
 
 Combo7.Text = Mid(Me.stb1.Panels(5).Text, 1, 4)
Combo8.Text = Mid(Me.stb1.Panels(5).Text, 5, 2)
Combo9.Text = Mid(Me.stb1.Panels(5).Text, 7, 2)

Dim E, J As Integer
'halqe haye saat _ tarikh
For I = 1391 To 1408
Combo7.AddItem (I)
Combo12.AddItem (I)
Next I

For J = 1 To 12
If J < 10 Then
Combo8.AddItem ("0" & J)
Combo11.AddItem ("0" & J)
Else
Combo8.AddItem (J)
Combo11.AddItem (J)
End If
Next J


For E = 1 To 31
If E < 10 Then
Combo9.AddItem ("0" & E)
Combo10.AddItem ("0" & E)
Else
Combo8.AddItem (E)
Combo10.AddItem (E)
End If

Next E
For I = 0 To 59
If I < 10 Then
Combo13.AddItem ("0" & I)
Combo14.AddItem ("0" & I)
Else
Combo13.AddItem (I)
Combo14.AddItem (I)
End If

Next I

For I = 0 To 23
If I < 10 Then
Combo16.AddItem ("0" & I)

Else
Combo16.AddItem (I)

End If

Next I






Tarhha.Refresh
Tarhha.Recordset.Sort = "sortgoroh"


For I = 1 To Tarhha.Recordset.RecordCount
Combo2.AddItem (Tarhha.Recordset.Fields("goroh"))
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
'BankStudent.Hide

End Sub

Private Sub mnubank_Click()
Beep

End Sub

Private Sub mnuclasjadid_Click()
ModiriyatCLASS.Show

End Sub

Private Sub label_vadie_sabt_Click()



VadieF.Combo1 = Me.Combo4.Text

VadieF.Combo2 = Me.Combo6.Text

'VadieF.Combo5


If Option3.Value = True Then VadieF.Combo5 = "Å—œ«Œ  ò«„·"

If Option4.Value = True Then VadieF.Combo5 = "»œÂò«—"

'VadieF.Show
'VadieF.Enabled = False
VadieF.Text1.Text = Me.TEP.Text

VadieF.text_auto_sabt.Text = "auto_sabt_bank_stu"
'VadieF.Hide

End Sub


Private Sub mnu_sabl_entezar_Click()
Exit Sub

If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "bank-mnu-sabt-List-entezar" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub

14082513






TEP.Text = Student.Recordset.Fields("parvande")
Call sabtentezarlabel_Click


End Sub

Private Sub mnudelete_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "bank-stu-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513

If Student.Recordset.RecordCount = 0 Then
o = MsgBox("‘„« ÂÌç ê“Ì‰Â «Ì  »—«Ì Õ–› ‰œ«—Ìœ ", vbCritical, " ÊÃÂ")
Else
o = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ﬁ—¬‰ ¬„Ê“ —« Õ–› ò‰Ìœ", vbYesNo + vbQuestion, "Õ–› ﬁ—¬‰ ¬„Ê“")
If o = vbYes Then
            STU2CLASS.Refresh
            STU2CLASS.RecordSource = " select * from stu2class where  parvande like ('%" + Student.Recordset.Fields("parvande") + "%')" ' and kodclass like ('%" + lkodclass.Caption + "%')"
            STU2CLASS.Refresh
            For I = 1 To STU2CLASS.Recordset.RecordCount
            STU2CLASS.Recordset.Delete
            STU2CLASS.Recordset.MoveNext
            Next I
            
Student.Recordset.Delete
End If
End If


End Sub

Private Sub MnusabtTanha_Click()
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ ‘„«—Â Å—Ê‰œÂ »Â  ‰Â«ÌÌ œ— ”Ì” „ À»  ‘Êœ", vbQuestion + vbYesNo, "À»  ‘„«—Â Å—Ê‰œÂ") = vbYes Then
' «÷«›Â ﬂ—œ‰
Student.Refresh
Student.Recordset.AddNew
Student.Recordset.Fields("Parvande") = TEP
Student.Recordset.AddNew
Student.Refresh
Else
Exit Sub
End If
End Sub

Private Sub mnuend_Click()
Entekhab.Show

End Sub

Private Sub mnufclass_Click()
FClassroom.Show

End Sub

Private Sub mnugozaresh_Click()
Gozaresh.Show

End Sub

Private Sub mnuhome_Click()
Entekhab.Show

End Sub

Private Sub mnuplclas_Click()
FClassroom.Show

End Sub

Private Sub mnuqeybat_Click()
QeybatF.Show

End Sub

Private Sub mnusabt_Click()
EmtahanF.Show
End Sub

Private Sub mnusabte_Click()
Call Command1_Click

End Sub

Private Sub mnusodor_Click()




Karname.Show

End Sub

Private Sub mnuupdate_Click()

Call Command4_Click


End Sub

Private Sub Moj_Click()
If Moj.Value = False Then
TE11.Text = "0"
End If
TE11.Enabled = False
TE11.Text = "0"
TE11.BackColor = &HC0C0C0

End Sub

Private Sub Option3_Click()
Combo4.Text = "200,000"
End Sub

Private Sub Option4_Click()
Combo4.Text = ""

End Sub

Private Sub stb1_PanelClick(ByVal Panel As ComctlLib.Panel)
Combo7.Text = Mid(Me.stb1.Panels(5).Text, 1, 4)
Combo8.Text = Mid(Me.stb1.Panels(5).Text, 5, 2)
Combo9.Text = Mid(Me.stb1.Panels(5).Text, 7, 2)
End Sub

Private Sub mot_Click()
TE11.Enabled = True
TE11.Text = ""
TE11.BackColor = &HFFFFFF


End Sub

Private Sub sabtentezarlabel_Click()


If Combo12.Text = "" Or Combo4.Text = "" Or Combo6.Text = "" Or Combo15.Text = "" Or Combo13.Text = "" Or Combo14.Text = "" Or Combo16.Text = "" Or Combo17.Text = "" Or Combo10.Text = "" Or Combo11.Text = "" Then
MsgBox "Œÿ« œ— Ê—ÊœÌ «ÿ·«⁄« " & Chr(10) & "»⁄÷Ì «“ ﬁ”„  Â« Å— ‰‘œÂ «” ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If






paziresh_table.Refresh
paziresh_table.RecordSource = "select * from paziresh_table where parvande like ('%" & TEP.Text & "%')"
paziresh_table.Refresh

If paziresh_table.Recordset.BOF = True Or paziresh_table.Recordset.EOF = True Then
'kasi peyda nashode
GoTo 1
Else

MsgBox "‰«„ «Ì‰ ﬁ—¬‰ ¬„Ê“ œ— ·Ì”  «‰Ÿ«— ÊÃÊœ œ«—œ", vbInformation, "·Ì”  «‰ Ÿ«—"
Exit Sub
End If


1




Call label_vadie_sabt_Click






paziresh_table.Refresh
paziresh_table.Recordset.AddNew
paziresh_table.Recordset.Fields("parvande") = TEP.Text
paziresh_table.Recordset.Fields("tarikh_paziresh") = Combo7.Text & "/" & Combo8.Text & "/" & Combo9.Text
If Check4.Value = 1 Then Xzaman_pishnahad_str = Xzaman_pishnahad_str & "-" & "’»Õ"
If Check2.Value = 1 Then Xzaman_pishnahad_str = Xzaman_pishnahad_str & "-" & "ŸÂ—"
If Check1.Value = 1 Then Xzaman_pishnahad_str = Xzaman_pishnahad_str & "-" & "ç—Œ‘Ì"

paziresh_table.Recordset.Fields("zaman_pishnahad") = Xzaman_pishnahad_str
Xzaman_pishnahad_str = ""
If Option1.Value = True Then paziresh_table.Recordset.Fields("sath_keyfi") = "„ Ê”ÿ"

If Option2.Value = True Then paziresh_table.Recordset.Fields("sath_keyfi") = "÷⁄Ì›"


If Option3.Value = True Then paziresh_table.Recordset.Fields("pardakht_vadie") = "ò«„·"

If Option4.Value = True Then paziresh_table.Recordset.Fields("pardakht_vadie") = "‰«ﬁ’"



paziresh_table.Recordset.Fields("mablaq_daryafti") = Combo4.Text

paziresh_table.Recordset.Fields("tahvil_be") = Combo6.Text
paziresh_table.Recordset.Fields("tarikh_morajee") = Combo12.Text & "/" & Combo11.Text & "/" & Combo10.Text
paziresh_table.Recordset.Fields("saat_morajee") = Combo16.Text & ":" & Combo14.Text & ":" & Combo13.Text
paziresh_table.Recordset.Fields("natije_mosahebe") = "0"
paziresh_table.Recordset.Fields("mosahebe_konande") = "0"
paziresh_table.Recordset.Fields("dore") = Combo15.Text

paziresh_table.Recordset.Fields("tozih") = Text2.Text

paziresh_table.Recordset.Update
paziresh_table.Refresh

Call Command8_Click


'MsgBox ""


End Sub

Private Sub TE1_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where name like ('%" + TE1.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If


End Sub

Private Sub TE10_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where gozarname like ('%" + TE10.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE11_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where farzand like ('%" + TE11.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE11_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub TE13_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where tell like ('%" + TE13.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE13_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub TE14_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where mob like ('%" + TE14.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE14_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub TE15_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + TE15.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE16_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where tozih like ('%" + TE16.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE2_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where famil like ('%" + TE2.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE3_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where namepedar like ('%" + TE3.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE4_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where tavalod like ('%" + TE4.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE5_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where shsh like ('%" + TE5.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub TE6_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where sadere like ('%" + TE6.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE7_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where meliyat like ('%" + TE7.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE8_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where mazhab like ('%" + TE8.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE9_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where kodmeli like ('%" + TE9.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
End Sub

Private Sub TE9_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub TEP_Change()
If chj.Value = 1 Then
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + TEP.Text + "%')"
Student.Refresh
Label27.Caption = Student.Recordset.RecordCount

End If
'Barcode.Text = TEP.Text

End Sub

Private Sub TEP_DblClick()
TEP.Text = ""

End Sub

Private Sub TEP_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
