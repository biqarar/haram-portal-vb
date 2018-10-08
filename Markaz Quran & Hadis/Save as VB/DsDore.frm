VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form DsDore 
   Caption         =   "À»  Ê «’·«Õ œÊ—Â Â«"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DsDore.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
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
      Height          =   2655
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   3495
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ ê—ÊÂ"
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
         Left            =   2640
         TabIndex        =   30
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   " — Ì» ê—ÊÂ"
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
         TabIndex        =   29
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ œÊ—Â"
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
         TabIndex        =   28
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Goroh"
         DataSource      =   "Tarhha"
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
         TabIndex        =   27
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "SortGoroh"
         DataSource      =   "Tarhha"
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
         TabIndex        =   26
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Name"
         DataSource      =   "Tarhha"
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
         TabIndex        =   25
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label16 
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
         Left            =   2640
         TabIndex        =   24
         Top             =   2160
         Width           =   585
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "tozih"
         DataSource      =   "Tarhha"
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
         Left            =   480
         TabIndex        =   23
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "xKoddore"
         DataSource      =   "Tarhha"
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
         TabIndex        =   22
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "SortName"
         DataSource      =   "Tarhha"
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
         Left            =   480
         TabIndex        =   21
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "òœ œÊ—Â"
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
         TabIndex        =   20
         Top             =   1800
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   " — Ì» œÊ—Â"
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
         TabIndex        =   19
         Top             =   1440
         Width           =   720
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "À»  œÊ—Â"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H000000FF&
      Caption         =   "Õ–›"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ã«Ìê“Ì‰Ì"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "„‘Œ’«  œÊ—Â"
      Height          =   2415
      Left            =   5160
      TabIndex        =   11
      Top             =   240
      Width           =   4215
      Begin VB.TextBox tozihtext 
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
         Height          =   350
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   3240
      End
      Begin VB.ComboBox Gorohname 
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
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox sortnamedoretext 
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
         Height          =   350
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   480
      End
      Begin VB.TextBox sortgorohtext 
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
         Height          =   350
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   480
      End
      Begin VB.TextBox kodgorohtext 
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
         Height          =   350
         Left            =   2640
         TabIndex        =   5
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox namedoretext 
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
         Height          =   350
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   1680
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "·ÿ›« œ— Ê«—œ ò—œ‰ «ÿ·«⁄«  œﬁ  ›—„«ÌÌœ"
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
         Left            =   720
         TabIndex        =   32
         Top             =   2040
         Width           =   2610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ÊœÌ⁄Â"
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
         TabIndex        =   17
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ê—ÊÂ"
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
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ œÊ—Â"
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
         TabIndex        =   15
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "òœ œÊ—Â"
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
         TabIndex        =   14
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   " — Ì» ê—ÊÂ"
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
         Left            =   840
         TabIndex        =   13
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   " — Ì» ‰«„ œÊ—Â"
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
         Left            =   720
         TabIndex        =   12
         Top             =   840
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   12240
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
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
         Connect         =   $"DsDore.frx":08CA
         OLEDBString     =   $"DsDore.frx":0953
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from stu2class"
         Caption         =   "STU2CLASS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
         Connect         =   $"DsDore.frx":09DC
         OLEDBString     =   $"DsDore.frx":0A65
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mclass"
         Caption         =   "mclass"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Connect         =   $"DsDore.frx":0AEE
         OLEDBString     =   $"DsDore.frx":0B77
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
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "DsDore.frx":0C00
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
      DefColWidth     =   107
      HeadLines       =   1
      RowHeight       =   29
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
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "œÊ—Â Â«"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Goroh"
         Caption         =   "Goroh"
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
         DataField       =   "SortGoroh"
         Caption         =   "SortGoroh"
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
         DataField       =   "Name"
         Caption         =   "Name"
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
         DataField       =   "SortName"
         Caption         =   "SortName"
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
         DataField       =   "xKoddore"
         Caption         =   "xKoddore"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      EndProperty
   End
   Begin MSAdodcLib.Adodc Tarhha 
      Height          =   375
      Left            =   3720
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   $"DsDore.frx":0C15
      OLEDBString     =   $"DsDore.frx":0C9E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tarhha"
      Caption         =   "Tarhha"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar Sb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   7545
      Width           =   9465
      _ExtentX        =   16695
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
End
Attribute VB_Name = "DsDore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Tarhha.Recordset.Fields("goroh") = Gorohname.Text
Tarhha.Recordset.Fields("sortgoroh") = sortgorohtext.Text
Tarhha.Recordset.Fields("name") = namedoretext.Text
Tarhha.Recordset.Fields("sortname") = sortnamedoretext.Text
Tarhha.Recordset.Fields("xkoddore") = kodgorohtext.Text
Tarhha.Recordset.Fields("tozih") = tozihtext.Text
Tarhha.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
Tarhha.Recordset.Fields("d") = Me.Sb1.Panels(3).Text




Tarhha.Recordset.Update
Tarhha.Refresh

End Sub

Private Sub Command5_Click()

Tarhha.Refresh
Tarhha.Recordset.Sort = "kod"
Tarhha.Recordset.MovePrevious
Dim kod_tarh As Integer

kod_tarh = (Tarhha.Recordset.Fields("kod")) + 1


Tarhha.Refresh
Tarhha.Recordset.AddNew
Tarhha.Recordset.Fields("tarh") = kod_tarh
Tarhha.Recordset.Fields("goroh") = Gorohname.Text
Tarhha.Recordset.Fields("sortgoroh") = sortgorohtext.Text
Tarhha.Recordset.Fields("name") = namedoretext.Text
Tarhha.Recordset.Fields("sortname") = sortnamedoretext.Text
Tarhha.Recordset.Fields("xkoddore") = kodgorohtext.Text
Tarhha.Recordset.Fields("tozih") = tozihtext.Text
Tarhha.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
Tarhha.Recordset.Fields("d") = Me.Sb1.Panels(3).Text




Tarhha.Recordset.Update
Tarhha.Refresh

End Sub

Private Sub Command6_Click()
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ œÊ—Â —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–› œÊ—Â") = vbYes Then
On Error Resume Next
Tarhha.Recordset.Delete




End If
End Sub

Private Sub DataGrid1_DblClick()
Gorohname.Text = Tarhha.Recordset.Fields("goroh")
sortgorohtext.Text = Tarhha.Recordset.Fields("sortgoroh")
 namedoretext.Text = Tarhha.Recordset.Fields("name")
 sortnamedoretext.Text = Tarhha.Recordset.Fields("sortname")
 kodgorohtext.Text = Tarhha.Recordset.Fields("xkoddore")
 tozihtext.Text = Tarhha.Recordset.Fields("tozih")



End Sub

Private Sub Form_Load()
Sb1.Panels(1).Text = user.OP.Text
Me.Sb1.Panels(3).Text = Taqvim.Tarikh.Caption
'Sb1.Panels(3).Text = Taqvim.Label1.Caption





Gorohname.AddItem ("⁄„Ê„Ì")
Gorohname.AddItem ("Õ›Ÿ ﬁ—¬‰ ò—Ì„")
Gorohname.AddItem ("ò«—ê«Â Â«")
Gorohname.AddItem (" —Ì»  „—»Ì")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show

End Sub

