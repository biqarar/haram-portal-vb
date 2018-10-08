VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Amar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ê—«“‘ êÌ—Ì «“ «ÿ·«⁄« "
   ClientHeight    =   11040
   ClientLeft      =   5055
   ClientTop       =   4275
   ClientWidth     =   19035
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "STU2CLASS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   19035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "ò·«” Â«Ì  ò—«—Ì"
      Height          =   735
      Left            =   14760
      TabIndex        =   157
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   450
      Left            =   13560
      TabIndex        =   156
      Text            =   "Text3"
      Top             =   6360
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "STU2CLASS.frx":08CA
      Height          =   2055
      Left            =   13440
      TabIndex        =   155
      Top             =   4080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "KOD"
         Caption         =   "KOD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Parvande"
         Caption         =   "Parvande"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Mablaq"
         Caption         =   "Mablaq"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "D"
         Caption         =   "D"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Op"
         Caption         =   "Op"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Dore"
         Caption         =   "Dore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Vazeyat"
         Caption         =   "Vazeyat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Terja"
         Caption         =   "Terja"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Elat"
         Caption         =   "Elat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Tozih"
         Caption         =   "Tozih"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "TTasvie"
         Caption         =   "TTasvie"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Bedehkar"
         Caption         =   "Bedehkar"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "MErja"
         Caption         =   "MErja"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "VErja"
         Caption         =   "VErja"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   3360
      Left            =   15840
      TabIndex        =   152
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   17520
      TabIndex        =   151
      Top             =   6600
      Width           =   375
   End
   Begin VB.Frame Frame11 
      Caption         =   "Ê÷⁄Ì  ÊœÌ⁄Â"
      Height          =   3855
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   132
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   975
         Left            =   3000
         TabIndex        =   154
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   120
         TabIndex        =   153
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   120
         TabIndex        =   150
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "»Â —Ê“ —”«‰Ì"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox Combo13 
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
         RightToLeft     =   -1  'True
         TabIndex        =   134
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo12 
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
         RightToLeft     =   -1  'True
         TabIndex        =   133
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label115 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ »Âœò«—«‰"
         Height          =   330
         Left            =   2400
         TabIndex        =   149
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label114 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   148
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label Label113 
         AutoSize        =   -1  'True
         Caption         =   "«“  «—ÌŒ"
         Height          =   330
         Left            =   2400
         TabIndex        =   147
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label109 
         AutoSize        =   -1  'True
         Caption         =   "„»·€ ÊœÌ⁄Â Â«Ì œ—Ì«› Ì"
         Height          =   330
         Left            =   2400
         TabIndex        =   146
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label Label108 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   145
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label107 
         AutoSize        =   -1  'True
         Caption         =   "„·»€ «—Ã«⁄«  ÊœÌ⁄Â"
         Height          =   330
         Left            =   2400
         TabIndex        =   144
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label Label106 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   143
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label105 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò”«‰Ì òÂ ÊœÌ⁄Â ‰œ«œÂ «‰œ"
         Height          =   330
         Left            =   2400
         TabIndex        =   142
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label Label104 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   141
         Top             =   2280
         Width           =   120
      End
      Begin VB.Label Label103 
         AutoSize        =   -1  'True
         Caption         =   " «  «—ÌŒ"
         Height          =   330
         Left            =   2280
         TabIndex        =   140
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label101 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   139
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         Caption         =   "œ—’œ Œÿ« (¬“„«Ì‘Ì"
         Height          =   330
         Left            =   2400
         TabIndex        =   138
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label97 
         AutoSize        =   -1  'True
         Caption         =   "„«Â"
         Height          =   330
         Left            =   5040
         TabIndex        =   137
         Top             =   960
         Width           =   225
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "”«·"
         Height          =   330
         Left            =   5040
         TabIndex        =   136
         Top             =   360
         Width           =   270
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Ê÷⁄Ì  «„ Õ«‰« "
      Height          =   3855
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   103
      Top             =   7080
      Width           =   5775
      Begin VB.ComboBox Combo10 
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
         RightToLeft     =   -1  'True
         TabIndex        =   128
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo9 
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
         RightToLeft     =   -1  'True
         TabIndex        =   126
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "»Â —Ê“ —”«‰Ì"
         Height          =   495
         Left            =   3960
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label95 
         AutoSize        =   -1  'True
         Caption         =   "”«·"
         Height          =   330
         Left            =   5040
         TabIndex        =   130
         Top             =   360
         Width           =   270
      End
      Begin VB.Label Label94 
         AutoSize        =   -1  'True
         Caption         =   "„«Â"
         Height          =   330
         Left            =   5040
         TabIndex        =   127
         Top             =   960
         Width           =   225
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "œ—’œ  ÃœÌœÌ"
         Height          =   330
         Left            =   2400
         TabIndex        =   122
         Top             =   3480
         Width           =   825
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   121
         Top             =   3360
         Width           =   120
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         Caption         =   "œ—’œ  ÃœÌœÌ 2"
         Height          =   330
         Left            =   2400
         TabIndex        =   120
         Top             =   3120
         Width           =   945
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   119
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   117
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label87 
         AutoSize        =   -1  'True
         Caption         =   " ÃœÌœÌ 1"
         Height          =   330
         Left            =   2400
         TabIndex        =   116
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   115
         Top             =   2280
         Width           =   120
      End
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         Caption         =   "œ—’œ ﬁ»Ê·Ì"
         Height          =   330
         Left            =   2400
         TabIndex        =   114
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   113
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "Ã„⁄  ÃœÌœÌ Â«"
         Height          =   330
         Left            =   2400
         TabIndex        =   112
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   111
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   " ÃœÌœÌ 2"
         Height          =   330
         Left            =   2400
         TabIndex        =   110
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   109
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "ﬁ»Ê·Ì"
         Height          =   330
         Left            =   2400
         TabIndex        =   108
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   107
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò· ò«—‰«„Â Â«"
         Height          =   330
         Left            =   2400
         TabIndex        =   106
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   105
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "œ—’œ  ÃœÌœÌ1"
         Height          =   330
         Left            =   2400
         TabIndex        =   104
         Top             =   2760
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   375
      Left            =   12120
      TabIndex        =   102
      Top             =   7320
      Visible         =   0   'False
      Width           =   2775
      Begin MSAdodcLib.Adodc Qeybat 
         Height          =   330
         Left            =   240
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
         Connect         =   $"STU2CLASS.frx":08DE
         OLEDBString     =   $"STU2CLASS.frx":0967
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from qeybat"
         Caption         =   "Qeybat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
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
         Left            =   240
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
         Connect         =   $"STU2CLASS.frx":09F0
         OLEDBString     =   $"STU2CLASS.frx":0A79
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from emtahan"
         Caption         =   "Emtahan"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
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
         Left            =   240
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
         Connect         =   $"STU2CLASS.frx":0B02
         OLEDBString     =   $"STU2CLASS.frx":0B8B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from stu2class"
         Caption         =   "STU2CLASS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
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
         Left            =   240
         Top             =   1440
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
         Connect         =   $"STU2CLASS.frx":0C14
         OLEDBString     =   $"STU2CLASS.frx":0C9D
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mclass"
         Caption         =   "mclass"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc teacher 
         Height          =   375
         Left            =   240
         Top             =   1080
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
         Connect         =   $"STU2CLASS.frx":0D26
         OLEDBString     =   $"STU2CLASS.frx":0DAF
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from teacher"
         Caption         =   "teacher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Student 
         Height          =   330
         Left            =   240
         Top             =   360
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
         Connect         =   $"STU2CLASS.frx":0E38
         OLEDBString     =   $"STU2CLASS.frx":0EC1
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from student"
         Caption         =   "Student"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Tarhha 
         Height          =   375
         Left            =   240
         Top             =   720
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
         Connect         =   $"STU2CLASS.frx":0F4A
         OLEDBString     =   $"STU2CLASS.frx":0FD3
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tarhha"
         Caption         =   "Tarhha"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Govahi 
         Height          =   330
         Left            =   240
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
         Connect         =   $"STU2CLASS.frx":105C
         OLEDBString     =   $"STU2CLASS.frx":10E5
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from govahi"
         Caption         =   "Govahi"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc vadie 
         Height          =   330
         Left            =   240
         Top             =   3240
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
         Connect         =   $"STU2CLASS.frx":116E
         OLEDBString     =   $"STU2CLASS.frx":11F7
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
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "«—”«· »Â »—‰«„Â «ò”·"
      Height          =   495
      Left            =   12000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   7920
      Width           =   3735
   End
   Begin VB.Frame Frame9 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   3495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   82
      Top             =   2880
      Width           =   3735
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å«Ì«‰"
         Height          =   345
         Left            =   2040
         TabIndex        =   100
         Top             =   3120
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
         TabIndex        =   99
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ ‘—Ê⁄"
         Height          =   345
         Left            =   2040
         TabIndex        =   98
         Top             =   2760
         Width           =   735
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
         TabIndex        =   97
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
         Height          =   330
         Index           =   0
         Left            =   2040
         TabIndex        =   96
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   330
         Left            =   2040
         TabIndex        =   95
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "„ﬁÿ⁄"
         Height          =   330
         Left            =   2040
         TabIndex        =   94
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
         Height          =   330
         Left            =   2040
         TabIndex        =   93
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄ "
         Height          =   330
         Left            =   2040
         TabIndex        =   92
         Top             =   1920
         Width           =   405
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "„œ—”"
         Height          =   330
         Left            =   2040
         TabIndex        =   91
         Top             =   2280
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
         Height          =   345
         Left            =   240
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
         Top             =   1920
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
         TabIndex        =   86
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
         TabIndex        =   85
         Top             =   2280
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
         TabIndex        =   84
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   " «"
         Height          =   330
         Left            =   720
         TabIndex        =   83
         Top             =   1920
         Width           =   120
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "⁄„·ò—œ «”« Ìœ"
      Height          =   2175
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   72
      Top             =   4800
      Width           =   4215
      Begin VB.ComboBox Combo7 
         Height          =   450
         Left            =   120
         TabIndex        =   81
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Combo5 
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
         TabIndex        =   73
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "‘—ò  ò‰‰œò«‰ «Ê·ÌÂ"
         Height          =   330
         Left            =   1560
         TabIndex        =   80
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   79
         Top             =   840
         Width           =   120
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "‘—ò  ò‰‰œê«‰ ‰Â«ÌÌ"
         Height          =   330
         Left            =   1560
         TabIndex        =   78
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   77
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "œ— ’œ— —Ì“‘"
         Height          =   330
         Left            =   1560
         TabIndex        =   76
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   75
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label9 
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
         Left            =   3000
         TabIndex        =   74
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Ê÷⁄Ì  €Ì»  Â«"
      Height          =   3855
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   7080
      Width           =   5775
      Begin VB.ComboBox Combo11 
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
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   129
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo8 
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
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   124
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "»Â —Ê“ —”«‰Ì"
         Height          =   495
         Left            =   4080
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label96 
         AutoSize        =   -1  'True
         Caption         =   "”«·"
         Height          =   330
         Left            =   5160
         TabIndex        =   131
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label93 
         AutoSize        =   -1  'True
         Caption         =   "„«Â"
         Height          =   330
         Left            =   5160
         TabIndex        =   125
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   " ⁄Âœ ò »Ì"
         Height          =   330
         Left            =   2400
         TabIndex        =   69
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   68
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò· €Ì»  Â«"
         Height          =   330
         Left            =   2400
         TabIndex        =   67
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   66
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "€Ì»   €Ì— „ÊÃÂ"
         Height          =   330
         Left            =   2400
         TabIndex        =   65
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   64
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   " «ŒÌ—"
         Height          =   330
         Left            =   2400
         TabIndex        =   63
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   62
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "„—Œ’Ì"
         Height          =   330
         Left            =   2400
         TabIndex        =   61
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   60
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   " ⁄Âœ ‘›«ÂÌ"
         Height          =   330
         Left            =   2400
         TabIndex        =   59
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   58
         Top             =   2280
         Width           =   120
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "€Ì»  „ÊÃÂ"
         Height          =   330
         Left            =   2400
         TabIndex        =   57
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   56
         Top             =   1200
         Width           =   120
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "»Â —Ê“ —”«‰Ì"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   120
      Width           =   3735
   End
   Begin VB.Frame Frame6 
      Caption         =   "⁄„·ò—œ «”« Ìœ"
      Height          =   1935
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   2880
      Width           =   4215
      Begin VB.ComboBox Combo6 
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
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Left            =   2400
         TabIndex        =   71
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   53
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "œ— ’œ— —Ì“‘"
         Height          =   330
         Left            =   1680
         TabIndex        =   52
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   51
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "‘—ò  ò‰‰œê«‰ ‰Â«ÌÌ"
         Height          =   330
         Left            =   1680
         TabIndex        =   50
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   49
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "‘—ò  ò‰‰œò«‰ «Ê·ÌÂ"
         Height          =   330
         Left            =   1680
         TabIndex        =   48
         Top             =   840
         Width           =   1290
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "¬„«— ò·«” Â«"
      Height          =   4095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   2880
      Width           =   4695
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
         ItemData        =   "STU2CLASS.frx":1280
         Left            =   2760
         List            =   "STU2CLASS.frx":1293
         TabIndex        =   33
         Text            =   "œÊ—Â —« «‰ Œ«» ò‰Ìœ"
         Top             =   360
         Width           =   1815
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
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "‘—ò  ò‰‰œò«‰ «Ê·ÌÂ"
         Height          =   330
         Left            =   1680
         TabIndex        =   46
         Top             =   2400
         Width           =   1290
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   45
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "‘—ò  ò‰‰œê«‰ ‰Â«ÌÌ"
         Height          =   330
         Left            =   1680
         TabIndex        =   44
         Top             =   2760
         Width           =   1320
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   43
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "œ— ’œ— —Ì“‘"
         Height          =   330
         Left            =   1680
         TabIndex        =   42
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   41
         Top             =   3120
         Width           =   120
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "ò· ò·«” Â«Ì »—ê“«— ‘œÂ œ— «Ì‰ ÿ—Õ"
         Height          =   330
         Left            =   1680
         TabIndex        =   40
         Top             =   960
         Width           =   2265
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” Â«Ì œ—  Õ«· »—ê“«—Ì"
         Height          =   330
         Left            =   1680
         TabIndex        =   38
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   37
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” Â«Ì  „«„ ‘œÂ"
         Height          =   330
         Left            =   1680
         TabIndex        =   36
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   120
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
         Left            =   2160
         TabIndex        =   34
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "«ÿ·«⁄«  ò·«” Â«"
      Height          =   1815
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   600
      Width           =   3735
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   30
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” Â«Ì  „«„ ‘œÂ"
         Height          =   330
         Left            =   1680
         TabIndex        =   29
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” Â«Ì œ—  Õ«· »—ê“«—Ì"
         Height          =   330
         Left            =   1680
         TabIndex        =   27
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "ò· ò·«” Â«Ì »—ê“«— ‘œÂ"
         Height          =   330
         Left            =   1680
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ê—ÊœÌ Ê Œ—ÊÃÌ ò·«”"
      Height          =   2895
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   4695
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò· ò”«‰Ì òÂ «“ ò·«” Õ–› ‘œÂ «‰œ"
         Height          =   330
         Left            =   1680
         TabIndex        =   22
         Top             =   1320
         Width           =   2685
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   120
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "”«Ì—"
         Height          =   330
         Left            =   1680
         TabIndex        =   20
         Top             =   2400
         Width           =   270
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Õ–› »Â œ·Ì· « „«„ ò·«”"
         Height          =   330
         Left            =   1680
         TabIndex        =   18
         Top             =   2040
         Width           =   1620
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Õ–› »Â œ·Ì· €Ì» "
         Height          =   330
         Left            =   1680
         TabIndex        =   16
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¬„«— ›⁄·Ì ‘—ò  ò‰‰œê«‰ œ— ò·«”"
         Height          =   330
         Left            =   1680
         TabIndex        =   14
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò· ‘—ò  ò‰‰œê«‰ œ— ò·«”"
         Height          =   330
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "«ÿ·«⁄«  Ê—ÊœÌ ”Ì” „"
      Height          =   2895
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.ComboBox Combo2 
         Height          =   450
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   450
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "»Â  ›òÌò ò‘Ê—"
         Height          =   330
         Left            =   2880
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "»Â  ›òÌò ‘Â—"
         Height          =   330
         Left            =   2880
         TabIndex        =   5
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò· ﬁ—¬‰ ¬„Ê“«‰ À»  ‰«„ ‘œÂ"
         Height          =   330
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   2310
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Å—Ê‰œÂ Â«Ì ‰«ﬁ’"
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label l2 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   120
      End
   End
   Begin VB.Menu mnuhime 
      Caption         =   "#"
   End
End
Attribute VB_Name = "Amar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Student.Refresh
Student.RecordSource = " select * from student where sadere like ('%" + Combo1.Text + "%')"
Student.Refresh
Label24.Caption = Student.Recordset.RecordCount

End Sub

Private Sub Combo1_Click()
Student.Refresh
Student.RecordSource = " select * from student where sadere like ('%" + Combo1.Text + "%')"
Student.Refresh
Label24.Caption = Student.Recordset.RecordCount

End Sub


Private Sub Combo2_Change()
Student.Refresh
Student.RecordSource = " select * from student where meliyat like ('%" + Combo2.Text + "%')"
Student.Refresh
Label22.Caption = Student.Recordset.RecordCount

End Sub

Private Sub Combo2_Click()
Student.Refresh
Student.RecordSource = " select * from student where meliyat like ('%" + Combo2.Text + "%')"
Student.Refresh
Label22.Caption = Student.Recordset.RecordCount

End Sub


Private Sub Combo4_Change()
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


Private Sub Combo5_Click()
Combo7.Clear

mclass.Refresh
mclass.RecordSource = "select * from mclass where ostad like ('%" + Combo5.Text + "%')"
mclass.Refresh


For I = 1 To mclass.Recordset.RecordCount


Combo7.AddItem (mclass.Recordset.Fields("kodclass"))
mclass.Recordset.MoveNext
Next I












End Sub


Private Sub Combo6_Click()
On Error GoTo 1

GoTo 2

1: MsgBox "ò·«” Ì«›  ‰‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub

2:

Label41.Caption = 0

mclass.Refresh
mclass.RecordSource = "select * from mclass where ostad like ('%" + Combo6.Text + "%')"
mclass.Refresh


For I = 1 To mclass.Recordset.RecordCount


STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" + mclass.Recordset.Fields("kodclass") + "%')"
STU2CLASS.Refresh
Label41.Caption = Val(Label41.Caption) + STU2CLASS.Recordset.RecordCount





mclass.Recordset.MoveNext



Next I

Label43.Caption = 0

mclass.Refresh
mclass.RecordSource = "select * from mclass where ostad like ('%" + Combo6.Text + "%')"
mclass.Refresh


For J = 1 To mclass.Recordset.RecordCount


STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where elat like ('%" + "« „«„ ò·«”" + "%') and kodclass like ('%" + mclass.Recordset.Fields("kodclass") + "%')"
STU2CLASS.Refresh
Label43.Caption = Val(Label43.Caption) + STU2CLASS.Recordset.RecordCount





mclass.Recordset.MoveNext



Next J






Label45.Caption = 100 - ((Val(Label43.Caption) * 100) \ (Val(Label41.Caption)))





End Sub

Private Sub Combo7_Click()
On Error GoTo 1

GoTo 2

1: MsgBox "ò·«” Ì«›  ‰‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub

2:

Label50.Caption = 0

mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Combo7.Text + "%')"
mclass.Refresh





STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where kodclass like ('%" + Combo7.Text + "%')"
STU2CLASS.Refresh
Label50.Caption = STU2CLASS.Recordset.RecordCount










Label46.Caption = 0

mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Combo7.Text + "%')"
mclass.Refresh





STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from stu2class where elat like ('%" + "« „«„ ò·«”" + "%') and kodclass like ('%" + Combo7.Text + "%')"
STU2CLASS.Refresh
Label46.Caption = STU2CLASS.Recordset.RecordCount
















Label10.Caption = 100 - ((Val(Label46.Caption) * 100) \ (Val(Label50.Caption)))



End Sub


Private Sub Command1_Click()
Student.Refresh
Student.RecordSource = " select * from student where parvande like ('%" + "" + "%')"
Student.Refresh
l1.Caption = Student.Recordset.RecordCount

Student.Refresh
Student.RecordSource = " select * from student where tozih like ('%" + "‰«ﬁ’" + "%') or tozih like ('%" + "À»  „Êﬁ " + "%') "
Student.Refresh
l2.Caption = Student.Recordset.RecordCount



STU2CLASS.Refresh
STU2CLASS.RecordSource = " select * from STU2CLASS where parvande like ('%" + "" + "%') "
STU2CLASS.Refresh
Label13.Caption = Student.Recordset.RecordCount


mclass.Refresh
For I = 1 To mclass.Recordset.RecordCount

Combo6.AddItem (mclass.Recordset.Fields("ostad"))
Combo5.AddItem (mclass.Recordset.Fields("ostad"))


mclass.Recordset.MoveNext

Next I



End Sub

Private Sub Command2_Click()

Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
Dim AD As String
On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:
Dim Mablaq_kol As Double
Mablaq_kol = 0
Dim Erjaat_ As Double
 Erjaat_ = 0
 

If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "allgovahi.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject("\\yafatemeh2-pc\F\Markaz Quran & Hadis\FORMXLS\allgovahi.xlsx")
End If



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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Dim Vadie_nadeha As Double
 Vadie_nadeha = 0
vadie.Refresh
vadie.RecordSource = "select * from vadie"
vadie.Refresh



vadie.Refresh
vadie.RecordSource = "select * from vadie where d like ('%" & Text2.Text & "%')"
vadie.Refresh
Mablaq_kol = 0
 Erjaat_ = 0
For I = 1 To vadie.Recordset.RecordCount

Mablaq_kol = Mablaq_kol + Val(vadie.Recordset.Fields("mablaq"))
vadie.Recordset.MoveNext

Next I
MsgBox Mablaq_kol, , "„»·€ ò· Ê—ÊœÌ"


''''''erjaar
Dim M_erja As Double
M_erja = 0


vadie.Refresh
vadie.RecordSource = "select * from vadie"
vadie.Refresh



vadie.Refresh
vadie.RecordSource = "select * from vadie where d like ('%" & Text2.Text & "%') and vazeyat like('%" & "«—Ã«⁄" & "%')"
vadie.Refresh
Erjaat_ = 0

M_erja = 0

For I = 1 To vadie.Recordset.RecordCount

M_erja = M_erja + Val(vadie.Recordset.Fields("merja"))

Erjaat_ = Erjaat_ + Val(vadie.Recordset.Fields("mablaq"))
vadie.Recordset.MoveNext

Next I
MsgBox Erjaat_, , "„·»€ ò· «—Ã«⁄ œ«œÂ ‘œÂ „»·€"

MsgBox M_erja, , "„»·€ À»  ‘œÂ œ— „»·€ «—Ã«⁄"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'



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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Command3_Click()
'be rooz resani
'vadie ha
Dim Mablaq_kol As Double
Mablaq_kol = 0
Dim Erjaat_ As Double
 Erjaat_ = 0
 
 Dim Vadie_nadeha As Double
 Vadie_nadeha = 0
vadie.Refresh
vadie.RecordSource = "select * from vadie"
vadie.Refresh



vadie.Refresh
vadie.RecordSource = "select * from vadie where d like ('%" & Text2.Text & "%')"
vadie.Refresh
Mablaq_kol = 0
 Erjaat_ = 0
For I = 1 To vadie.Recordset.RecordCount

Mablaq_kol = Mablaq_kol + Val(vadie.Recordset.Fields("mablaq"))
vadie.Recordset.MoveNext

Next I
MsgBox Mablaq_kol, , "„»·€ ò· Ê—ÊœÌ"


''''''erjaar
Dim M_erja As Double
M_erja = 0


vadie.Refresh
vadie.RecordSource = "select * from vadie"
vadie.Refresh



vadie.Refresh
vadie.RecordSource = "select * from vadie where d like ('%" & Text2.Text & "%') and vazeyat like('%" & "«—Ã«⁄" & "%')"
vadie.Refresh
Erjaat_ = 0

M_erja = 0

For I = 1 To vadie.Recordset.RecordCount

M_erja = M_erja + Val(vadie.Recordset.Fields("merja"))

Erjaat_ = Erjaat_ + Val(vadie.Recordset.Fields("mablaq"))
vadie.Recordset.MoveNext

Next I
MsgBox Erjaat_, , "„·»€ ò· «—Ã«⁄ œ«œÂ ‘œÂ „»·€"

MsgBox M_erja, , "„»·€ À»  ‘œÂ œ— „»·€ «—Ã«⁄"



End Sub

Private Sub Command4_Click()
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where  temtahan like ('" & Combo10.Text & Combo9.Text & "%')"
Emtahan.Refresh
Label78.Caption = Emtahan.Recordset.RecordCount
'ò· «„ Õ«‰« 

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where VAZEYAT like ('%" & "ﬁ»Ê·Ì" & "%') and temtahan like ('" & Combo10.Text & Combo9.Text & "%')"
Emtahan.Refresh
Label80.Caption = Emtahan.Recordset.RecordCount


Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where VAZEYAT like ('%" & " ÃœÌœÌ 1" & "%')and temtahan like ('" & Combo10.Text & Combo9.Text & "%')"
Emtahan.Refresh
Label88.Caption = Emtahan.Recordset.RecordCount

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where VAZEYAT like ('%" & " ÃœÌœÌ 2" & "%')and temtahan like ('" & Combo10.Text & Combo9.Text & "%')"
Emtahan.Refresh
Label82.Caption = Emtahan.Recordset.RecordCount

Label84.Caption = (Val(Label82.Caption) + Val(Label88.Caption))



Label86.Caption = Val(Label80.Caption) * 100 / Val(Label78.Caption)

Label76.Caption = Val(Label88.Caption) * 100 / Val(Label78.Caption)

Label89.Caption = Val(Label82.Caption) * 100 / Val(Label78.Caption)

Label91.Caption = (Val(Label82.Caption) + Val(Label88.Caption)) * 100 / Val(Label78.Caption)



End Sub

Private Sub Command5_Click()

Qeybat.Refresh
Qeybat.RecordSource = "select * from Qeybat where MAH like ('%" & Combo8.Text & "%')and sal like ('%" & Combo11.Text & "%')"
Qeybat.Refresh
Label65.Caption = Qeybat.Recordset.RecordCount




Qeybat.Refresh
Qeybat.RecordSource = "select * from Qeybat where MAH like ('%" & Combo8.Text & "') and sal like ('%" & Combo11.Text & "%') AND NOE LIKE ('%" & "€Ì— „ÊÃÂ" & "%')"
Qeybat.Refresh


Label63.Caption = Qeybat.Recordset.RecordCount

Qeybat.Refresh
Qeybat.RecordSource = "select * from Qeybat where MAH like ('%" & Combo8.Text & "')and sal like ('%" & Combo11.Text & "%') and NOE like ('" & "€Ì»  „ÊÃÂ" & "')"
Qeybat.Refresh
Label55.Caption = Qeybat.Recordset.RecordCount


Qeybat.Refresh
Qeybat.RecordSource = "select * from Qeybat where MAH like ('%" & Combo8.Text & "')and sal like ('%" & Combo11.Text & "%') AND NOE LIKE ('%" & " «ŒÌ—" & "%')"
Qeybat.Refresh
Label61.Caption = Qeybat.Recordset.RecordCount


Qeybat.Refresh
Qeybat.RecordSource = "select * from Qeybat where MAH like ('%" & Combo8.Text & "')and sal like ('%" & Combo11.Text & "%') AND NOE LIKE ('%" & "„—Œ’Ì" & "%')"
Qeybat.Refresh
Label59.Caption = Qeybat.Recordset.RecordCount


Qeybat.Refresh
Qeybat.RecordSource = "select * from Qeybat where MAH like ('%" & Combo8.Text & "')and sal like ('%" & Combo11.Text & "%') AND NOE LIKE ('%" & "‘›«ÂÌ" & "%')"
Qeybat.Refresh
Label57.Caption = Qeybat.Recordset.RecordCount


Qeybat.Refresh
Qeybat.RecordSource = "select * from Qeybat where MAH like ('%" & Combo8.Text & "')and sal like ('%" & Combo11.Text & "%') AND NOE LIKE ('%" & "ò »Ì" & "%')"
Qeybat.Refresh
Label67.Caption = Qeybat.Recordset.RecordCount



End Sub

Private Sub Command6_Click()
vadie.Refresh
For I = 1 To vadie.Recordset.RecordCount
On Error Resume Next

List1.AddItem (vadie.Recordset.Fields("d"))
vadie.Recordset.MoveNext
Next I
MsgBox ":"

End Sub

Private Sub Command7_Click()

MsgBox mid(Text2.Text, 1, 4)
MsgBox mid(Text2.Text, 6, 2)
MsgBox mid(Text2.Text, 9, 2)

End Sub

Private Sub Form_Load()
'ò‘Ê— Â«

Combo2.AddItem ("«Ì—«‰")
Combo2.AddItem ("«›€«‰")
Combo2.AddItem ("⁄—«ﬁ")
Combo2.AddItem ("Â‰œ")
Combo2.AddItem ("Å«ò” «‰")
'Combo2.AddItem ("")
'Combo2.AddItem ("")
'‘Â— Â«
Combo1.AddItem ("ﬁ„")
Combo1.AddItem (" Â—«‰")

'Combo8.AddItem (" „«„Ì „«Â Â«")

Combo8.AddItem ("›—Ê—œÌ‰")
Combo8.AddItem ("«—œÌ»Â‘ ")
Combo8.AddItem ("Œ—œ«œ")
Combo8.AddItem (" Ì—")
Combo8.AddItem ("„—œ«œ")
Combo8.AddItem ("‘Â—ÌÊ—")
Combo8.AddItem ("„Â—")
Combo8.AddItem ("¬»«‰")
Combo8.AddItem ("¬–—")
Combo8.AddItem ("œÌ")
Combo8.AddItem ("»Â„‰")
Combo8.AddItem ("«”›‰œ")



For I = 1390 To 1408
Combo10.AddItem (I & "/")
Combo11.AddItem (I)

Next I


For I = 1 To 12 Step 1  ' »—«Ì Ê«—œ ﬂ—œ‰ ‘„«—Â —Ê“ œ— ÃœÊ· «ÿ·«⁄«  €Ì 
If I < 10 Then
Combo9.AddItem ("0" & I)
Else
Combo9.AddItem (I)
End If
Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show

End Sub


Private Sub mnuhime_Click()
Entekhab.Show

End Sub

